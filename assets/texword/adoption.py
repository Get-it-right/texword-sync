"""Adopt an existing Word document while retaining untouched native content."""
from __future__ import annotations
import copy,difflib,json,re
from pathlib import Path
from texword import ConversionError,pandoc,load,dump,tag,run,field,raw,NS,W,ROOT,bookmark_name
from review_syntax import prepare_reviews,format_review

def simplify_latex(value,*,columns=None,column_widths=None):
 from source_spans import _read_group
 # Keep semantic text in TeX; layout is owned by the Word template adapter.
 while True:
  m=re.search(r'\\texorpdfstring\s*\{',value)
  if not m:break
  first=_read_group(value,m.end()-1,len(value),'{','}')
  second=_read_group(value,first[1],len(value),'{','}') if first else None
  if not second:raise ConversionError('Malformed generated heading text')
  value=value[:m.start()]+value[first[0]:first[1]-1].strip()+value[second[1]:]
 if columns:
  m=re.search(r'\\begin\{longtable\}',value)
  if m:
   p=m.end()
   if value[p:p+1]=='[':
    optional=_read_group(value,p,len(value),'[',']');p=optional[1]
   group=_read_group(value,p,len(value),'{','}')
   if not group:raise ConversionError('Malformed generated table columns')
   specification='l'*columns
   if r'\begin{minipage}' in value:
    # Pandoc's block cells use minipage{\linewidth}. In an l column that
    # becomes the entire page width for EACH cell. Give those cells bounded
    # paragraph columns, using the edited Word grid when it is available.
    if column_widths is not None and (len(column_widths)!=columns or any(w<=0 for w in column_widths)):
     raise ConversionError('Invalid native table grid widths; the source was not replaced')
    weights=column_widths if column_widths is not None else [1]*columns
    total=sum(weights)
    specification=''.join('p{\\dimexpr'+format(w/total,'.8f')+'\\linewidth-2\\tabcolsep\\relax}' for w in weights)
    # A natural-width multirow can also overflow a bounded paragraph column.
    value=re.sub(r'(\\multirow(?:\[[^\]]*\])?\{[^{}]+\})\{\*\}',lambda match:match[1]+r'{\linewidth}',value)
   value=value[:m.start()]+'\\begin{longtable}{'+specification+'}'+value[group[1]:]
   if '\\endfirsthead' in value:
    first,tail=value.split('\\endfirsthead',1)
    if '\\endhead' not in tail:raise ConversionError('Malformed generated table header')
    value=first+tail.split('\\endhead',1)[1]
   elif '\\endhead' in value:
    value=value.replace('\\endhead','',1)
  # Labels are structural metadata, outside the caption text.
  value=re.sub(r'(\\caption\{[^\n]*?)\\label\{([^}]+)\}(\})',r'\1\3\\label{\2}',value)
  # Pandoc 2.19 emits empty rule suffixes that its reader turns into cell text.
  # Restrict cleanup to standalone generated rules, preserving real parentheses.
  value=re.sub(r'(?m)^([ \t]*\\(?:toprule|midrule|bottomrule))\(\)(?=[ \t]*\r?$)',r'\1',value)
 # TeX already interprets prose dashes. Do not rewrite identifiers, English
 # titles or mathematical subtraction while simplifying layout commands.
 return value.strip()

def create_project(original,project,*,body_start,body_end):
 from adopt_extract import extract_blocks
 from texword import adapt,digest
 original=Path(original).resolve();project=Path(project).resolve()
 if project.exists():raise ConversionError('Adoption project already exists; reuse its current source instead')
 project.mkdir(parents=True)
 adapter=project/'template_adapter';profile=adapt(original,adapter,body_start=body_start,body_end=body_end)
 result=extract_blocks(original,profile['body_start'],profile['body_end'],project)
 from docx import Document
 original_body=Document(original).element.body
 source='\\documentclass{article}\n\\usepackage{amsmath,booktabs,longtable,array}\n\\newcommand{\\twcomment}[4][Codex]{#3}\n\\newcommand{\\twsuggest}[5][Codex]{#3}\n\\begin{document}\n\n'
 records=[];layouts=[]
 for b in result['blocks']:
  if not b['latex'].strip():layouts.append({'position':len(source),'indices':b['original_indices']});continue
  grid=original_body[b['original_indices'][-1]].findall('w:tblGrid/w:gridCol',NS) if b['kind']=='table' else None
  columns=len(grid) if grid is not None else None
  widths=[int(column.get(W+'w','0')) for column in grid] if grid is not None else None
  start=len(source);source+=simplify_latex(b['latex'],columns=columns,column_widths=widths);end=len(source);source+='\n\n'
  records.append(dict(b,start=start,end=end,latex=source[start:end]))
 source+='\\end{document}\n'
 (project/'report.tex').write_bytes(source.encode('utf8'))
 metadata={'schema':1,'original':str(original),'original_sha256':digest(original),'source':source,'blocks':records,'layouts':layouts,
           'bibliography':result['bibliography'],'source_path':str(project/'report.tex')}
 dump(adapter/'imported.json',metadata)
 return {'project':str(project),'source':str(project/'report.tex'),'adapter':str(adapter),'blocks':len(records),'layout_nodes':sum(len(x['indices']) for x in layouts),'references':len(result['bibliography'])}

def write_workflow(project,word):
 project=Path(project).resolve();word=Path(word).resolve()
 dump(project/'workflow.json',{'source':str(project/'report.tex'),'adapter':str(project/'template_adapter'),'code':str(project/'generate_word.py'),'word':str(word)})
 program='"""Fixed daily Word/LaTeX synchronization commands."""\nfrom pathlib import Path\nimport sys\nsys.path.insert(0, '+repr(str(ROOT))+' )\nfrom workflow_runtime import main\nif __name__ == "__main__":\n raise SystemExit(main(Path(__file__).resolve().parent))\n'
 (project/'workflow.py').write_text(program,encoding='utf8')

def _position(opcodes,p,*,end=False):
 for op,a,b,c,d in opcodes:
  if a<=p<=b:
   if op=='equal':return c+p-a
   if p==a:return c
   if p==b:return d
   return None
 return opcodes[-1][4] if opcodes else p

def source_opcodes(before,after):
 """Match document lines first; do character alignment only inside changed lines."""
 if before==after:return [('equal',0,len(before),0,len(after))]
 left=before.splitlines(keepends=True);right=after.splitlines(keepends=True)
 la=[0];ra=[0]
 for line in left:la.append(la[-1]+len(line))
 for line in right:ra.append(ra[-1]+len(line))
 out=[]
 for op,a,b,c,d in difflib.SequenceMatcher(a=[x.rstrip('\r\n') for x in left],b=[x.rstrip('\r\n') for x in right],autojunk=False).get_opcodes():
  if op=='equal':
   for i,j in zip(range(a,b),range(c,d)):
    n=len(left[i].rstrip('\r\n'));out.append(('equal',la[i],la[i]+n,ra[j],ra[j]+n))
    if left[i][n:]==right[j][n:]:out.append(('equal',la[i]+n,la[i+1],ra[j]+n,ra[j+1]))
    else:out.append(('replace',la[i]+n,la[i+1],ra[j]+n,ra[j+1]))
  elif op=='replace':
   for kind,i,j,k,l in difflib.SequenceMatcher(a=before[la[a]:la[b]],b=after[ra[c]:ra[d]],autojunk=False).get_opcodes():out.append((kind,la[a]+i,la[a]+j,ra[c]+k,ra[c]+l))
  else:out.append((op,la[a],la[b],ra[c],ra[d]))
 return out

def bibliography_entries(latex):
 from source_spans import _read_group
 entries=[];matches=list(re.finditer(r'\\bibitem\s*\{',latex))
 for i,m in enumerate(matches):
  group=_read_group(latex,m.end()-1,len(latex),'{','}')
  if group is None:raise ConversionError('Malformed bibliography key')
  # source_spans returns (inner_start, after_closing_brace).
  a,b=group;key=latex[a:b-1].strip()
  stop=matches[i+1].start() if i+1<len(matches) else latex.find(r'\end{thebibliography}',b)
  if stop<0:raise ConversionError('Bibliography has no closing environment')
  if not re.fullmatch(r'[A-Za-z][A-Za-z0-9:_-]*',key):raise ConversionError('Unsupported bibliography key '+key)
  entries.append({'key':key,'latex':latex[b:stop].strip(),'number':len(entries)+1})
 if not entries:raise ConversionError('Bibliography must contain explicit bibitem entries')
 if len({e['key'] for e in entries})!=len(entries):raise ConversionError('Duplicate bibliography keys')
 return entries

def prepare_adopted(adapter,source):
 """Reuse exact baseline blocks; parse edits through the existing LaTeX frontend."""
 from source_spans import split_source
 from roundtrip import _split_math
 from roundtrip_model import sha
 adapter=Path(adapter);source=Path(source);base=load(adapter/'imported.json');current=source.read_bytes().decode('utf8')
 opcodes=source_opcodes(base['source'],current)
 spans=[];covered=[]
 section_mode=load(adapter/'profile.json').get('source_mapping_granularity')=='section'
 if section_mode:
  # Paragraph boundaries may legitimately disappear inside a section. Parse
  # the current source, retaining exact native originals wherever they match.
  originals={}
  for old in base['blocks']:originals.setdefault(old['latex'],[]).append(old)
  used=set()
  for s in split_source(current):
   if spans and re.fullmatch(r'\\addcontentsline(?:\{[^{}]*\}){3}',s['latex'].strip()):
    spans[-1]['end']=s['end'];spans[-1]['latex']=current[spans[-1]['start']:s['end']];spans[-1]['original_indices']=[]
    continue
   old=next((r for r in originals.get(s['latex'],[]) if r['start'] not in used),None)
   if old:used.add(old['start'])
   spans.append(dict(s,original_indices=old['original_indices'] if old and s['kind']!='header' else [],old_indices=old['original_indices'] if old else []))
 for old in ([] if section_mode else base['blocks']):
  a=_position(opcodes,old['start']);b=_position(opcodes,old['end'],end=True)
  if a is None or b is None:raise ConversionError('A source edit crossed adopted block boundaries; split the edit into the existing logical blocks')
  if b<=a:continue
  latex=current[a:b]
  spans.append({'start':a,'end':b,'latex':latex,'kind':old['kind'],'original_indices':old['original_indices'] if latex==old['latex'] else [],'old_indices':old['original_indices']})
  covered.append((a,b))
 begin=re.search(r'\\begin\{document\}',current);finish=re.search(r'\\end\{document\}',current)
 if not begin or not finish:raise ConversionError('Adopted source must retain its document environment')
 # Explicit additions in the gaps are ordinary LaTeX blocks.
 cursor=begin.end()
 for a,b in ([] if section_mode else sorted(covered)+[(finish.start(),finish.start())]):
  if a>cursor and current[cursor:a].strip():
   for s in split_source(current[cursor:a]):spans.append(dict(s,start=s['start']+cursor,end=s['end']+cursor,original_indices=[],old_indices=[]))
  cursor=max(cursor,b)
 spans.sort(key=lambda s:s['start'])
 bibliography=[]
 for s in spans:
  if '\\begin{thebibliography}' in s['latex']:s['kind']='bibliography';bibliography=bibliography_entries(s['latex'])
 registry={b['key']:b for b in bibliography}
 if not registry and base.get('bibliography'):raise ConversionError('Adopted source lost its bibliography')
 events=[];reviews=[];reuse={};layout={};review_keys=set()
 for i,s in enumerate(spans):
  s['id']='s'+str(i+1).zfill(4);s['math_sources']=[];sid=s['id']
  fragment,definitions=prepare_reviews(s['latex'])
  for d in definitions:
   if d['key'] in review_keys:raise ConversionError('Duplicate review id '+d['key'])
   review_keys.add(d['key']);d['compiled_id']=str(len(reviews));d['rendered_text']=format_review(d);reviews.append(d)
  if s['original_indices']:reuse[sid]=s['original_indices']
  if s['kind']=='bibliography':
   fragment='\n\n'.join('TWADOPTBIB'+str(e['number']).zfill(4)+'TOKEN '+e['latex'] for e in bibliography)
  events.append((s['start'],1,'\n\nTWRT_BEGIN_'+sid+'\n\n'+fragment+'\n\nTWRT_END_'+sid+'\n\n'))
 for i,l in enumerate(base['layouts']):
  position=_position(opcodes,l['position'])
  if position is None:
   if not section_mode:raise ConversionError('A source edit crossed a preserved section boundary')
   position=next(c for op,a,b,c,d in opcodes if a<=l['position']<=b)
  # Layout is a node boundary, never a new paragraph inserted into prose.
  containing=next((s for s in spans if s['start']<position<s['end']),None)
  if containing:position=containing['end']
  name='TWADOPTLAYOUT'+str(i).zfill(4);layout[name]=l['indices'];events.append((position,0,'\n\n'+name+'\n\n'))
 # Block-boundary markers are rendered temporarily, then lowered to Python calls.
 prepared=current[:begin.end()]+'\n'+''.join(e[2] for e in sorted(events,key=lambda e:(e[0],e[1])))+'\n'+current[finish.start():]
 ast_raw,_=pandoc(['-f','latex','-t','json'],cwd=source.parent,input=prepared);ast=json.loads(ast_raw)
 active=None
 def collect_math(o,s):
  if isinstance(o,dict):
   if o.get('t')=='Math':s['math_sources'].append({'latex':o['c'][1],'display':o['c'][0]['t']=='DisplayMath'})
   else:
    for v in o.values():collect_math(v,s)
  elif isinstance(o,list):
   for v in o:collect_math(v,s)
 byid={s['id']:s for s in spans}
 for block in ast['blocks']:
  text=''.join(x.get('c','') for x in block.get('c',[]) if isinstance(x,dict) and x.get('t')=='Str') if block.get('t')=='Para' else ''
  m=re.fullmatch(r'TWRT_(BEGIN|END)_(s\d+)',text)
  if m:active=m.group(2) if m.group(1)=='BEGIN' else None
  elif active:collect_math(block,byid[active])
 bibkeys={str(e['number']).zfill(4):e for e in bibliography};bid=300000
 def convert(o):
  nonlocal bid
  if isinstance(o,list):return [convert(x) for x in o]
  if not isinstance(o,dict):return o
  if o.get('t')=='Cite':
   if len(o['c'][0])>3:raise ConversionError('Each citation group supports at most three references')
   xml=''
   for ci in o['c'][0]:
    key=ci['citationId']
    if key not in registry:raise ConversionError('Unknown adopted bibliography key '+key)
    xml+=run('[')+field('REF '+key+' \\h',registry[key]['number'])+run(']')
   return raw(xml)
  if o.get('t')=='Para' and o['c'] and o['c'][0].get('t')=='Str':
   m=re.match(r'^TWADOPTBIB(\d+)TOKEN$',o['c'][0]['c'])
   if m:
    entry=bibkeys[m.group(1)];bid+=1
    xml=run('[')+__import__('lxml').etree.tostring(tag('bookmarkStart',id=bid,name=entry['key']),encoding='unicode')+field('SEQ AblBibRef \\* ARABIC',entry['number'])+__import__('lxml').etree.tostring(tag('bookmarkEnd',id=bid),encoding='unicode')+run('] ')
    tail=o['c'][1:]
    if tail and tail[0].get('t')=='Space':tail=tail[1:]
    return {'t':'Div','c':[['',[],[['custom-style','Bibliography']]],[{'t':'Para','c':[raw(xml)]+[convert(x) for x in tail]}]]}
  return {k:convert(v) for k,v in o.items()}
 ast=convert(ast)
 source_map={'schema':1,'source_path':str(source.resolve()),'source':current,'source_sha256':sha(current),'spans':spans,'adopted':True,
             'native_references':{k:{'key':k,'kind':'citation','number':entry['number']} for k,entry in registry.items()}}
 return ast,source_map,reviews,reuse,layout
