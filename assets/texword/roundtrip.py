"""Import mapped Word content edits as bounded LaTeX source patches."""
from __future__ import annotations
import copy,difflib,io,json,os,re,tempfile
from pathlib import Path
from lxml import etree as E
from docx import Document
from docx.oxml import OxmlElement
from texword import ConversionError,pandoc,dump,run
from roundtrip_model import NS,W,M,sha,signature,doc_relationships,read_manifest,field_parts,math_model,equivalent_math

def _split_math(blocks):
 out=[]
 for b in blocks:
  if b.get('t')=='Para':
   current=[]
   for x in b['c']:
    if x.get('t')=='Math' and x['c'][0]['t']=='DisplayMath':
     if any(i['t'] not in ('Space','SoftBreak') for i in current):out.append({'t':'Para','c':current})
     out.append({'t':'Para','c':[x]});current=[]
    else:current.append(x)
   if any(i['t'] not in ('Space','SoftBreak') for i in current):out.append({'t':'Para','c':current})
  else:out.append(b)
 return out

def _normalize_imported_math(obj):
 # Word leaves an empty math run after an invisible closing delimiter; Pandoc
 # writes it as one or more terminal control spaces. Keep internal spacing and
 # visible delimiters intact, and normalize only inside parsed math nodes.
 if isinstance(obj,dict):
  if obj.get('t')=='Math':
   latex=re.sub(r'(\\right\.)(?:\\\s+)+$',r'\1',obj['c'][1])
   # Pandoc wraps Word's Unicode em spaces in text{quad}, which its LaTeX
   # reader rejects. Preserve the spaces as math commands; leave prose intact.
   obj['c'][1]=re.sub(r'\\text\{((?:\\(?:qquad|quad)(?![A-Za-z])|\\[,;:! ]|\s)+)\}',
    lambda match:match[1].replace(r'\!','')+' ',latex)
  else:
   for value in obj.values():_normalize_imported_math(value)
 elif isinstance(obj,list):
  for value in obj:_normalize_imported_math(value)

def _parse_comparison(obj):
 if isinstance(obj,list):return [_parse_comparison(x) for x in obj]
 if not isinstance(obj,dict):return obj
 obj={k:_parse_comparison(v) for k,v in obj.items()}
 if obj.get('t') in ('Para','Plain'):
  while obj['c'] and obj['c'][0].get('t') in ('SoftBreak','Space'):obj['c'].pop(0)
  while obj['c'] and obj['c'][-1].get('t') in ('SoftBreak','Space'):obj['c'].pop()
 if obj.get('t')=='Link' and dict(obj['c'][0][2]).get('reference-type'):obj['c'][1]=[]
 return obj

def map_source(path,full_ast):
 from source_spans import split_source
 from review_syntax import prepare_reviews
 path=Path(path);source=path.read_bytes().decode('utf8');spans=split_source(source)
 begin=re.search(r'\\begin\{document\}',source);preamble=source[:begin.start()] if begin else ''
 blocks=[];plain=[];kept=[]
 for s in spans:
  prepared,_=prepare_reviews(s['latex'])
  fragment=preamble+'\n\\begin{document}\n'+prepared+'\n\\end{document}'
  raw,_=pandoc(['-f','latex','-t','json'],cwd=path.parent,input=fragment)
  part=_split_math(json.loads(raw)['blocks'])
  if not part:continue  # Comments and non-rendering text remain untouched in the source.
  s['math_sources']=[]
  def collect_math(o):
   if isinstance(o,dict):
    if o.get('t')=='Math':s['math_sources'].append({'latex':o['c'][1],'display':o['c'][0]['t']=='DisplayMath'})
    else:
     for v in o.values():collect_math(v)
   elif isinstance(o,list):
    for v in o:collect_math(v)
  collect_math(part)
  plain+=part;kept.append(s)
  for marker in ('BEGIN','END'):
   if marker=='END':blocks+=part
   blocks.append({'t':'RawBlock','c':['openxml','<w:p xmlns:w="'+NS['w']+'">'+run('TWRT_'+marker+'_'+s['id'])+'</w:p>']})
 if _parse_comparison(plain)!=_parse_comparison(_split_math(full_ast['blocks'])):
  raise ConversionError('Roundtrip source block coverage differs from full LaTeX parsing; no approximate source map was created. Split dependent macros/environments or use ordinary compile.')
 full_ast=copy.deepcopy(full_ast);full_ast['blocks']=blocks
 return full_ast,{'schema':1,'source_path':str(path.resolve()),'source':source,'source_sha256':sha(source),'spans':kept}

def _new_run(value,fmt=None):
 r=OxmlElement('w:r')
 if fmt and any(fmt):
  pr=OxmlElement('w:rPr');r.append(pr)
  for active,key in zip(fmt,('b','i')):
   if active:pr.append(OxmlElement('w:'+key))
 t=OxmlElement('w:t');t.set('{http://www.w3.org/XML/1998/namespace}space','preserve');t.text=value;r.append(t)
 return r

def _sanitize(elements,record,manifest,review_import=None,images=None,doc=None):
 """Remove generated numbering; retain native math and replace field semantics with tokens."""
 elements=copy.deepcopy(elements);replacements={};counter=0;math_index=0
 for e in elements:
  for node in e.iter():
   if node.tag in (W+'sdt',W+'altChunk',W+'customXml'):
    raise ConversionError('Unsupported changed block container: '+E.QName(node).localname)
  if E._Element.xpath(e,'.//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo',namespaces=NS):
   raise ConversionError('Accept/reject tracked revisions before importing: '+record['id'])
  if e.findall('.//w:hyperlink',NS) or (e.findall('.//w:vMerge',NS) and not manifest.get('allow_merged_table_edits')):
   raise ConversionError('Changed hyperlink/vertically merged blocks require manual reconciliation: '+record['id'])
  for p in ([e] if e.tag==W+'p' else [])+list(e.findall('.//w:p',NS)):
   parts=field_parts(p,include_comments=True);items=[]
   for kind,value,fmt in parts:
    if kind in ('review_start','review_end'):
     spec=(review_import or {}).get(value)
     if spec is None:raise ConversionError('Word comment is not mapped: '+str(value))
     token='TWRTTOKEN'+str(counter)+'END';counter+=1
     replacements[token]=dict(spec,kind='review_open' if kind=='review_start' else 'review_close',review_kind=spec['kind'])
     items.append(('text',token,None))
    elif kind=='field':
     if re.match(r'^SEQ\s+(?:Table|Figure|Equation|AblBibRef)\b',value):token='TWRTSEQ'
     else:
      match=re.match(r'^REF\s+(\S+)',value)
      info=manifest['references'].get(match.group(1)) if match else None
      if info is None:raise ConversionError('Unmapped Word field: '+value)
      token='TWRTTOKEN'+str(counter)+'END';counter+=1
      replacements[token]=info
     items.append(('text',token,None))
    elif kind=='math':
     root=value.find('m:oMath',NS) if value.tag==M+'oMathPara' else value
     originals=record.get('math_sources',[]);baselines=record.get('math_baselines',[])
     model=math_model(root)
     # Match semantics, not a shifting ordinal: prose splits and formula
     # insertions/deletions must not discard intact styled LaTeX sources.
     display=value.tag==M+'oMathPara'
     candidates=[i for i,old in enumerate(baselines) if i<len(originals)
                 and originals[i].get('display',False)==display and equivalent_math(model,old)]
     matched=math_index if math_index in candidates else candidates[0] if candidates else None
     if record['kind']!='equation' and matched is not None:
      token='TWRTTOKEN'+str(counter)+'END';counter+=1
      replacements[token]={'kind':'math','key':originals[matched]['latex'],'display':originals[matched]['display']}
      items.append(('text',token,None))
     else:
      # Pandoc's DOCX reader drops explicit roman math style and rewrites °C
      # as an empty-base superscript. Preserve actual styled runs as tokens;
      # never infer units from letters, parentheses or surrounding prose.
      from review_syntax import encode_plain
      for run in root.iter(M+'r'):
       style=run.find('m:rPr/m:sty',NS);text_nodes=run.findall('m:t',NS)
       payload=''.join(t.text or '' for t in text_nodes)
       if (style is None or style.get(M+'val')!='p' or not text_nodes
           or run.find('m:rPr/m:scr',NS) is not None or run.find('m:rPr/m:nor',NS) is not None
           or not any(c.isalpha() or c=='°' for c in payload)):continue
       token='TWRTMATHRUN'+str(counter)+'END';counter+=1
       replacements[token]={'kind':'math_run','latex':'\\mathrm{'+encode_plain(payload)+'}'}
       text_nodes[0].text=token
       for extra in text_nodes[1:]:run.remove(extra)
      items.append((kind,value,fmt))
     math_index+=1
    elif kind=='drawing' and images is not None:
     token='TWRTTOKEN'+str(counter)+'END';counter+=1
     replacements[token]={'kind':'image','latex':images.inline(value)}
     items.append(('text',token,None))
    elif kind in ('text','drawing','tab','break'):items.append((kind,value,fmt))
    else:raise ConversionError('Unsupported changed Word content in '+record['id']+': '+kind)
   merged=[]
   for item in items:
    if merged and item[0]=='text' and merged[-1][0]=='text' and item[2]==merged[-1][2]:merged[-1]=('text',merged[-1][1]+item[1],item[2])
    else:merged.append(item)
   items=merged
   from section_mapping import _table_caption
   if record['kind']=='figure_caption' or (record['kind']=='table' and doc is not None and _table_caption(doc,p,record)):
    prefix=''.join(value for kind,value,_ in items if kind=='text')
    match=re.match(r'^\s*(?:图|Figure|表|Table)\s*(?:S?TWRTSEQ|S?\d+)[\s\u3000.:：、]*',prefix,re.I)
    remaining=match.end() if match else 0
    for i,(kind,value,fmt) in enumerate(items):
     if kind=='text' and remaining:
      removed=min(remaining,len(value));items[i]=(kind,value[removed:],fmt);remaining-=removed
   # Numbering added by the forward renderer belongs to the template, not source prose.
   text_indexes=[i for i,v in enumerate(items) if v[0]=='text']
   if text_indexes:
    i=text_indexes[0];kind,value,fmt=items[i]
    if record['kind']=='header':value=re.sub(r'^(?:[一二三四五六七八九十]+、|\d+(?:\.\d+)*\s+)','',value)
    items[i]=(kind,value,fmt)
   for x in list(p):
    if x.tag!=W+'pPr':p.remove(x)
   for kind,value,fmt in items:
    if kind=='text':p.append(_new_run(value,fmt))
    elif kind=='math':p.append(value)
    elif kind=='drawing':
     r=OxmlElement('w:r');r.append(value);p.append(r)
    elif kind=='break':
     r=OxmlElement('w:r');r.append(OxmlElement('w:br'));p.append(r)
    elif kind=='tab':
     if record['kind']=='equation':continue
     p.append(_new_run(' '))
    # Layout-only equation tabs are deliberately not semantic text.
 return elements,replacements

def _restore_fields(obj,replacements):
 if isinstance(obj,list):
  out=[]
  for x in obj:
   restored=_restore_fields(x,replacements)
   if isinstance(x,dict) and x.get('t')=='Str' and isinstance(restored,list):out.extend(restored)
   else:out.append(restored)
  return out
 if not isinstance(obj,dict):return obj
 if obj.get('t')=='Math':
  obj=copy.deepcopy(obj)
  obj['c'][1]=re.sub(r'TWRTMATHRUN\d+END',lambda m:replacements[m.group()]['latex'],obj['c'][1])
  return obj
 if obj.get('t')=='Str':
  value=obj['c'];pattern=r'(\[)?(TWRTTOKEN\d+END)(\])?|(?P<seq>TWRTSEQ)'
  out=[];pos=0
  for m in re.finditer(pattern,value):
   if m.start()>pos:out.append({'t':'Str','c':value[pos:m.start()]})
   if not m.group('seq'):
    info=replacements[m.group(2)];command='cite' if info['kind']=='citation' else 'ref'
    # Brackets are generated only for bibliography references. Noncitation brackets are prose.
    if m.group(1) and command!='cite':out.append({'t':'Str','c':'['})
    if info['kind'].startswith('review_'):
     from review_syntax import encode_plain
     if info['kind']=='review_open':
      author=info.get('author','Codex')
      if '[' in author or ']' in author:raise ConversionError('Comment author with brackets requires manual reconciliation')
      name={'suggestion':'twsuggest','revision':'twrevision'}.get(info['review_kind'],'twcomment')
      latex='\\'+name+('['+encode_plain(author)+']' if author!='Codex' else '')+'{'+info['key']+'}{'
     elif info['review_kind'] in ('suggestion','revision'):
      proposal=info.get('proposal_source') if info['review_kind']=='revision' else None
      latex='}{'+(proposal if proposal is not None else encode_plain(info['proposal']))+'}{'+encode_plain(info['reason'])+'}'
     else:latex='}{'+encode_plain(info['note'])+'}'
    elif info['kind']=='image':latex=info['latex']
    else:latex=('\\['+info['key']+'\\]' if info.get('display') else '$'+info['key']+'$') if info['kind']=='math' else '\\'+command+'{'+info['key']+'}'
    out.append({'t':'RawInline','c':['latex',latex]})
    if m.group(3) and command!='cite':out.append({'t':'Str','c':']'})
   pos=m.end()
  if pos==0:return obj
  if pos<len(value):out.append({'t':'Str','c':value[pos:]})
  return out
 return {k:_restore_fields(v,replacements) for k,v in obj.items()}

def block_to_latex(doc,elements,record,manifest,review_import=None,images=None):
 # A source comment is not represented in Word. Refuse a lossy rewrite of that block.
 for match in re.finditer('%',record['latex']):
  prefix=record['latex'][:match.start()];slashes=len(prefix)-len(prefix.rstrip('\\'))
  tail=record['latex'][match.end():].splitlines()[0] if record['latex'][match.end():] else ''
  if slashes%2==0 and tail.strip():raise ConversionError('Changed source block contains LaTeX comments; manual reconciliation preserves them: '+record['id'])
 if images is not None and '\\begin{figure}' in record['latex']:
  from section_mapping import _convert
  return _convert(doc,elements,[record],manifest,review_import,images)
 bibkeys=[]
 if record['kind']=='bibliography':
  for e in elements:
   keys=E._Element.xpath(e,'.//w:bookmarkStart/@w:name',namespaces=NS)
   keys=[k for k in keys if k in manifest.get('native_references',{})]
   if len(keys)!=1:raise ConversionError('Each bibliography entry must retain one native bibliography bookmark')
   bibkeys.append(keys[0])
 sanitized,replacements=_sanitize(elements,record,manifest,review_import,images,doc=doc)
 with tempfile.TemporaryDirectory(prefix='texword_reverse_') as temp:
  # Reuse style/relationship definitions from the edited Word; replace only the body.
  # Saving through the original package retains styles and relationships reliably.
  tempdoc=Path(temp)/'fragment.docx'
  package=getattr(doc,'_texword_fragment_package',None)
  if package is None:
   buffer=io.BytesIO();doc.save(buffer);package=buffer.getvalue()
   doc._texword_fragment_package=package
  mini=Document(io.BytesIO(package))
  body=mini.element.body
  for x in list(body):
   if x.tag!=W+'sectPr':body.remove(x)
  for x in sanitized:body.insert(len(body)-1,x)
  # Review annotations are protected tokens now. Pandoc must not import their
  # package as unanchored note text in the temporary fragment.
  for rid,rel in list(mini.part.rels.items()):
   if not rel.is_external and rel.reltype.endswith(('/comments','/commentsExtended','/commentsIds','/commentsExtensible')):mini.part.drop_rel(rid)
  mini.save(tempdoc)
  raw,_=pandoc([tempdoc,'-f','docx','-t','json']);ast=json.loads(raw)
 _normalize_imported_math(ast['blocks'])
 ast['blocks']=_restore_fields(ast['blocks'],replacements)
 if record['kind']=='bibliography':
  paragraphs=[b for b in ast['blocks'] if b.get('t') in ('Para','Plain')]
  if len(paragraphs)!=len(bibkeys):raise ConversionError('Bibliography entry boundaries changed; reconcile new/deleted references in LaTeX')
  blocks=[{'t':'RawBlock','c':['latex',r'\begin{thebibliography}{999}']}]
  for key,p in zip(bibkeys,paragraphs):
   # The original native SEQ is already removed, leaving only its brackets.
   while p['c'] and p['c'][0].get('t')=='Space':p['c'].pop(0)
   if p['c'] and p['c'][0].get('t')=='Str' and p['c'][0]['c']=='[':
    p['c'].pop(0)
    if p['c'] and p['c'][0].get('t')=='Str' and p['c'][0]['c'].startswith(']'):p['c'][0]['c']=p['c'][0]['c'][1:].lstrip()
   if p['c'] and p['c'][0].get('t')=='Str':p['c'][0]['c']=re.sub(r'^\[\]\s*','',p['c'][0]['c'])
   while p['c'] and (p['c'][0].get('t')=='Space' or p['c'][0].get('c')==''):p['c'].pop(0)
   blocks.extend([{'t':'RawBlock','c':['latex','\\bibitem{'+key+'}']},p])
  blocks.append({'t':'RawBlock','c':['latex',r'\end{thebibliography}']});ast['blocks']=blocks
 # Strip template-generated caption prefixes after Pandoc has combined adjacent runs.
 def strip_number(inlines):
  while inlines and inlines[0].get('t') in ('Space','SoftBreak'):inlines.pop(0)
  if inlines and inlines[0].get('t')=='Str':
   inlines[0]['c']=re.sub(r'^(?:表|图)\s*(?:S?TWRTSEQ|S?\d+)[\s\u3000.:：、]*','',inlines[0]['c'])
   if not inlines[0]['c']:inlines.pop(0)
  while inlines and (inlines[0].get('t') in ('Space','SoftBreak') or inlines[0].get('c')==''):inlines.pop(0)
 if record['kind']=='header':
  command=re.match(r'\\((?:sub)*section)(\*)?',record['latex'])
  level=record.get('heading_level') or (1+command.group(1).count('sub') if command else None)
  if level and len(ast['blocks'])==1 and ast['blocks'][0].get('t') in ('Para','Plain'):
   ast['blocks'][0]={'t':'Header','c':[level,['',[],[]],ast['blocks'][0]['c']]}
  for b in ast['blocks']:
   if b.get('t')=='Header':b['c'][1]=['',(['unnumbered']+([] if '\\addcontentsline' in record['latex'] else ['unlisted'])) if command and command.group(2) else [],[]]
 if record['kind']=='equation':
  maths=[]
  def collect(o):
   if isinstance(o,dict):
    if o.get('t')=='Math':maths.append(o['c'][1])
    else:
     for v in o.values():collect(v)
   elif isinstance(o,list):
    for v in o:collect(v)
  collect(ast['blocks'])
  if len(maths)!=1:raise ConversionError('A mapped equation must contain one Word math object')
  if record['latex'].lstrip().startswith(r'\['):return '\\['+maths[0]+'\\]'
  env=re.match(r'\s*\\begin\{(equation\*?|align\*?|gather\*?|multline\*?|displaymath)\}',record['latex']);env=env.group(1) if env else 'equation*'
  labels=''.join('\\label{'+s+'}' for s in re.findall(r'\\label\{([^}]+)\}',record['latex']))
  return '\\begin{'+env+'}'+labels+'\n'+maths[0]+'\n\\end{'+env+'}'
 if record['kind']=='table':
  table=next((b for b in ast['blocks'] if b.get('t')=='Table'),None)
  captions=[b for b in ast['blocks'] if b.get('t') in ('Para','Plain')]
  if table is None or len(captions)>1:raise ConversionError('Changed table structure needs manual reconciliation')
  if captions:
   strip_number(captions[0]['c']);table['c'][1][1]=[{'t':'Plain','c':captions[0]['c']}]
  ast['blocks']=[table]
 rendered,_=pandoc(['-f','json','-t','latex','--wrap=none'],input=json.dumps(ast,ensure_ascii=False))
 if record['kind']=='table':
  from adoption import simplify_latex
  native_table=next((element for element in elements if element.tag==W+'tbl'),None)
  widths=[int(col.get(W+'w','0')) for col in native_table.findall('w:tblGrid/w:gridCol',NS)] if native_table is not None else None
  rendered=simplify_latex(rendered,columns=len(table['c'][2]),column_widths=widths)
 if record['kind']=='bibliography':rendered=re.sub(r'(\\bibitem\{[^}]+\})\s*\n\s*',r'\1 ',rendered)
 # Preserve semantic labels, never import generated heading/caption identifiers.
 labels=re.findall(r'\\label\{([^}]+)\}',record['latex'])
 rendered=re.sub(r'\\label\{[^}]+\}','',rendered).strip()
 for label in labels:
  if record['kind']=='table':
   caption=re.search(r'\\caption(?:\[[^]]*\])?\{',rendered)
   if not caption:raise ConversionError('Cannot locate regenerated table caption')
   # Structural labels belong after the caption, never in its visible text.
   from review_syntax import _read_group
   group=_read_group(rendered,caption.end()-1,len(rendered),'{','}')
   if group is None:raise ConversionError('Malformed regenerated table caption')
   at=group[2]
   rendered=rendered[:at]+'\\label{'+label+'}'+rendered[at:]
  else:rendered+='\\label{'+label+'}'
 # The original eqref is expressed as parentheses around REF in Word.
 rendered=re.sub(r'\(\\ref\{([^}]+)\}\)',r'\\eqref{\1}',rendered)
 if 'TWRT' in rendered:raise ConversionError('Unresolved reverse-conversion token')
 return rendered

def _mapped_interval(base,current,start,end):
 from adoption import source_opcodes
 delta=0
 for op,a,b,c,d in source_opcodes(base,current):
  if op=='equal':continue
  intersects=(a<end and b>start) or (a==b and start<=a<=end)
  if intersects:raise ConversionError('Both Word and LaTeX changed the same source block')
  if b<=start:delta+=(d-c)-(b-a)
 return start+delta,end+delta


def same_inline_math_notation(left,right):
 r"""Compare inline delimiters and transport newlines; leave literal code intact."""
 from review_syntax import _iter_commands
 def key(value):
  value=value.replace('\r\n','\n').replace('\r','\n')
  parts=[];cursor=0;opening=None
  for command in _iter_commands(value,0,len(value)):
   if command.symbol=='(' and opening is None:opening=command
   elif command.symbol==')' and opening is not None:
    parts.extend((value[cursor:opening.start],'$',value[opening.end:command.start],'$'))
    cursor=command.end;opening=None
  parts.append(value[cursor:]);return ''.join(parts)
 return key(left)==key(right)

def ensure_import_preamble(text):
 """Add only compilation prerequisites introduced by supported Word content."""
 from review_syntax import prepare_reviews,_without_comments
 begin=re.search(r'\\begin\{document\}',text)
 if not begin:return text,[]
 preamble=text[:begin.start()];body=text[begin.end():];active=_without_comments(preamble)
 packages={name.strip() for match in re.finditer(r'\\(?:usepackage|RequirePackage)(?:\s*\[[^\]]*\])?\s*\{([^}]+)\}',active) for name in match[1].split(',')}
 additions=[]
 if r'\begin{longtable}' in body and 'longtable' not in packages:additions.append(r'\usepackage{longtable}')
 if re.search(r'\\(?:toprule|midrule|bottomrule)\b',body) and 'booktabs' not in packages:additions.append(r'\usepackage{booktabs}')
 _,reviews=prepare_reviews(text)
 macros={'comment':('twcomment',4),'suggestion':('twsuggest',5),'revision':('twrevision',5)}
 kinds={review['kind'] for review in reviews}
 for kind in macros:
  if kind not in kinds:continue
  name,count=macros[kind]
  if not re.search(r'\\(?:newcommand|renewcommand|providecommand|DeclareRobustCommand)\s*\{?\s*\\'+name+r'\b',active):
   additions.append('\\providecommand{\\'+name+'}['+str(count)+'][Codex]{#3}')
 if not additions:return text,[]
 newline='\r\n' if '\r\n' in text else '\n'
 return text[:begin.start()]+newline.join(additions)+newline+text[begin.start():],additions


def sync_word(word,source,*,output=None,apply=False,report=None,baseline_word=None,merge_bases=None,capture_baselines=False,allow_structure=True,allow_template_updates=False):
 """Every failure returns evidence for this input, never a stale report."""
 word=Path(word).resolve();source=Path(source).resolve()
 target=Path(report).resolve() if report else word.with_suffix('.sync.json')
 destination=source if apply else Path(output).resolve() if output else None
 if target.suffix.lower()!='.json' or target in (source,word,destination):
  raise ConversionError('Sync report must be a separate .json file')
 before=target.stat().st_mtime_ns if target.exists() else None
 identities={'word_sha256':sha(word.read_bytes()) if word.is_file() else None,
             'source_sha256':sha(source.read_bytes()) if source.is_file() else None}
 try:
  result=_sync_word(word,source,output=output,apply=apply,report=target,baseline_word=baseline_word,merge_bases=merge_bases,capture_baselines=capture_baselines,allow_structure=allow_structure,allow_template_updates=allow_template_updates)
 except (ConversionError,ValueError) as exc:
  changed=target.exists() and target.stat().st_mtime_ns!=before
  prior=json.loads(target.read_text('utf8')) if changed else {}
  result=dict(prior,status='conflict',word=str(word),source=str(source),applied=False,error=str(exc),**identities)
  result.setdefault('changed_blocks',0)
  if not result.get('conflicts'):result['conflicts']=[{'reason':str(exc)}]
  dump(target,result)
  if isinstance(exc,ConversionError):raise
  raise ConversionError(str(exc)) from exc
 result.update(identities)
 full=json.loads(target.read_text('utf8'));full.update(identities);dump(target,full)
 return result


def _sync_word(word,source,*,output=None,apply=False,report=None,baseline_word=None,merge_bases=None,capture_baselines=False,allow_structure=True,allow_template_updates=False):
 word=Path(word).resolve();source=Path(source).resolve()
 if source.suffix.lower()!='.tex':raise ConversionError('Source must be a .tex file')
 if output and apply:raise ConversionError('Choose either --output or --apply')
 destination=source if apply else Path(output).resolve() if output else None
 if destination and destination.suffix.lower()!='.tex':raise ConversionError('Reverse output must use .tex')
 if destination==source and not apply:raise ConversionError('Use --apply to update the source in place')
 if destination and destination!=source and destination.exists():raise ConversionError('Reverse output already exists; choose a new derived output')
 report=Path(report).resolve() if report else word.with_suffix('.sync.json')
 if report.suffix.lower()!='.json' or report in (source,word,destination):raise ConversionError('Sync report must be a separate .json file')
 word_hash=sha(word.read_bytes());doc=Document(word);manifest=read_manifest(doc);current=source.read_bytes().decode('utf8');base=manifest['source']
 if manifest.get('semantic_version',1)<2:
  if baseline_word is None:raise ConversionError('Legacy semantic mapping needs a matching generated baseline; use the project workflow or baseline_word')
  from roundtrip_model import upgrade_semantic_manifest
  manifest=upgrade_semantic_manifest(Document(baseline_word),manifest)
 from tracked_reviews import project_revisions
 from native_comparison import project_reader
 try:project_reader(doc,manifest)
 except ValueError as exc:raise ConversionError(str(exc)) from exc
 try:project_revisions(doc,manifest,project_accepted=True)
 except ValueError as exc:raise ConversionError(str(exc)) from exc
 if sha(base)!=manifest['source_sha256']:raise ConversionError('Source-map baseline integrity failed')
 from image_roundtrip import ImageImport
 images=ImageImport(doc,source,manifest,destination)
 if manifest.get('granularity')=='section':
  from section_mapping import sync_sections
  return sync_sections(word,source,doc,manifest,destination,report,images,merge_bases=merge_bases,capture_baselines=capture_baselines,allow_structure=allow_structure,allow_template_updates=allow_template_updates)
 rels=doc_relationships(doc);mapped={};outside=[];order=[];conflicts=[];patches=[]
 from review_comments import read_comments,review_snapshot,prepare_import_reviews,reconcile_review_ids
 try:comments=read_comments(doc)
 except ValueError as exc:comments={};conflicts.append({'reason':str(exc)})
 for e in doc.element.body:
  tag=e.find('w:sdtPr/w:tag',NS) if e.tag==W+'sdt' else None
  if tag is not None and tag.get(W+'val','').startswith('texword:'):
   sid=tag.get(W+'val').split(':',1)[1];order.append(sid)
   if sid in mapped:conflicts.append({'id':sid,'reason':'Duplicate source block'})
   mapped[sid]=list(e.find('w:sdtContent',NS))
  else:outside.append(e)
 if order!=[s['id'] for s in manifest['spans']]:conflicts.append({'reason':'Source blocks deleted, inserted, moved, or their tags removed; manual reconciliation required'})
 if signature(outside,rels)!=manifest['outside_baseline']:conflicts.append({'reason':'Content outside mapped source blocks changed (template, bibliography, or unanchored insertion)'})
 try:
  outside_reviews,_=reconcile_review_ids(review_snapshot(outside,comments),manifest.get('outside_reviews',[]))
  if outside_reviews!=manifest.get('outside_reviews',[]):conflicts.append({'reason':'Comments outside mapped source blocks changed; manual reconciliation required'})
 except ValueError as exc:conflicts.append({'reason':str(exc)})
 skipped_latex=[];word_bases={r['id']:base[r['start']:r['end']] for r in manifest['spans']}
 for record in manifest['spans']:
  sid=record['id']
  if sid not in mapped:continue
  try:
   snapshot=review_snapshot(mapped[sid],comments)
   normalized,origins=reconcile_review_ids(snapshot,record.get('review_baseline',[]))
   changed=signature(mapped[sid],rels,record.get('math_baselines'))!=record['baseline'] or normalized!=record.get('review_baseline',[])
   if not changed and merge_bases is None and not manifest.get('accepted_review_sources'):continue
   review_import=prepare_import_reviews(snapshot,manifest,origins)
   from adoption import source_opcodes,_position
   from section_mapping import merge_source_edits
   ops=source_opcodes(base,current);a=_position(ops,record['start']);b=_position(ops,record['end'])
   if a is None or b is None:raise ConversionError('Source changes crossed mapped block boundaries')
   replacement=block_to_latex(doc,mapped[sid],record,manifest,review_import,images) if changed else word_bases[sid]
   if same_inline_math_notation(replacement,base[record['start']:record['end']]):replacement=base[record['start']:record['end']]
   word_bases[sid]=replacement
   replacement,skipped=merge_source_edits((merge_bases or {}).get(sid,base[record['start']:record['end']]),current[a:b],replacement,
       accepted_reviews=manifest.get('accepted_review_sources'))
   skipped_latex.extend(dict(item,id=sid) for item in skipped)
   if replacement!=current[a:b]:patches.append({'id':sid,'kind':record['kind'],'start':a,'end':b,'before':current[a:b],'after':replacement})
  except (ConversionError,ValueError) as exc:conflicts.append({'id':sid,'reason':str(exc)})
 result={'status':'conflict' if conflicts else 'preview','word':str(word),'source':str(source),'output':str(destination) if destination else None,
         'changed_blocks':len(patches),'conflicts':conflicts,'patches':patches,'applied':False,
         'skipped_latex_edits':skipped_latex,'merge_policy':'word_wins_overlaps',
         'review_comments':[dict(c,id=cid) for cid,c in comments.items()],
         'approval_policy':'Suggestions remain comments until explicitly confirmed; removal or resolution never applies the proposed text.',
         'format_policy':'Template layout is owned by the template adapter; body bold/italic changes are imported.'}
 if capture_baselines:result['word_bases']=word_bases
 if conflicts:
  dump(report,result);raise ConversionError('Word import stopped without changing LaTeX; see '+str(report))
 updated=current
 for p in sorted(patches,key=lambda p:p['start'],reverse=True):updated=updated[:p['start']]+p['after']+updated[p['end']:]
 if result['changed_blocks']:updated,result['preamble_additions']=ensure_import_preamble(updated)
 try:images.validate(updated)
 except ConversionError as exc:
  result['status']='conflict';result['conflicts'].append({'reason':str(exc)});dump(report,result);raise
 result['image_assets']=images.inventory()
 result['diff']=''.join(difflib.unified_diff(current.splitlines(True),updated.splitlines(True),fromfile=str(source),tofile=str(destination or source)))
 if destination:
  # Validate the patched document syntax before any source file is replaced.
  from review_syntax import prepare_reviews
  validated,_=prepare_reviews(updated)
  pandoc(['-f','latex','-t','json'],cwd=source.parent,input=validated)
  # Check that the source did not change during conversion, then replace atomically.
  if source.read_bytes().decode('utf8')!=current:raise ConversionError('LaTeX changed during import; retry from the saved current source')
  if sha(word.read_bytes())!=word_hash:raise ConversionError('Word changed during import; retry from the newly saved document')
  destination.parent.mkdir(parents=True,exist_ok=True)
  with tempfile.NamedTemporaryFile(dir=destination.parent,suffix='.tex',delete=False) as f:
   temporary=Path(f.name);f.write(updated.encode('utf8'))
  try:images.publish(destination,temporary)
  finally:
   if temporary.exists():temporary.unlink()
  result['status']='applied';result['applied']=True
 dump(report,result)
 return {k:v for k,v in result.items() if k not in ('patches','diff')}
