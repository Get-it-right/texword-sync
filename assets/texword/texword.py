"""Compile LaTeX to inspectable Word-building Python code, using a generated Word-template adapter."""
from __future__ import annotations
import argparse,copy,hashlib,json,os,re,shutil,subprocess,sys,tempfile,time,zipfile,pprint,posixpath
from pathlib import Path
from lxml import etree as E
if __name__ == '__main__':sys.modules['texword']=sys.modules[__name__]
ROOT=Path(__file__).resolve().parent
NS={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main','m':'http://schemas.openxmlformats.org/officeDocument/2006/math','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
W='{'+NS['w']+'}'
ROLES={'body':('BodyText','Body Text'),'heading1':('Heading1','heading 1'),'heading2':('Heading2','heading 2'),'heading3':('Heading3','heading 3'),'caption':('Caption','Caption'),'table_text':('TWTableText','TW Table Text'),'bibliography':('Bibliography','Bibliography')}

class ConversionError(RuntimeError):pass
def digest(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def text(e):return ''.join(e.xpath('.//w:t/text() | .//m:t/text()',namespaces=NS))
def document_text(p):
 with zipfile.ZipFile(p) as z:return text(E.fromstring(z.read('word/document.xml')))
def load(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def dump(p,x):Path(p).write_text(json.dumps(x,ensure_ascii=False,indent=2),encoding='utf8')
def tag(local,**attrs):return E.Element(W+local,{W+k:str(v) for k,v in attrs.items()})
def merge_props(base,extra):
 result=copy.deepcopy(base)
 if extra is not None:
  for x in extra:
   for old in list(result):
    if old.tag==x.tag:result.remove(old)
   result.append(copy.deepcopy(x))
 return result
def strip_props(p):
 for x in list(p):
  if E.QName(x).localname in ['pStyle','sectPr','numPr','rStyle','pageBreakBefore']:p.remove(x)
 return p
def write_zip(source,target,changes):
 with zipfile.ZipFile(source) as zi,zipfile.ZipFile(target,'w',zipfile.ZIP_DEFLATED) as zo:
  for info in zi.infolist():zo.writestr(info,changes.get(info.filename,zi.read(info.filename)))
def xml_bytes(e):return E.tostring(e,encoding='UTF-8',xml_declaration=True,standalone=True)
def pandoc_path():
 path=os.environ.get('PANDOC') or shutil.which('pandoc')
 if not path:raise ConversionError('Pandoc not found; install Pandoc or set PANDOC.')
 return path
def _prepare_panel_figures(source):
 """Keep shared captions when Pandoc's LaTeX reader drops multi-image figures."""
 from source_spans import _read_group
 groups={}
 def replace(match):
  body=match.group(1)
  pictures=list(re.finditer(r'(?:\\twimagecrop\{[^{}]+\}\s*)?\\includegraphics(?:\[[^]]*\])?\{[^{}]+\}',body))
  if len(pictures)<2:return match[0]
  parts={};ranges=[(p.start(),p.end()) for p in pictures]
  for name in ('caption','label'):
   found=list(re.finditer(r'\\'+name+r'\s*(?=\{)',body))
   if len(found)!=1:return match[0]
   m=found[0];group=_read_group(body,m.end(),len(body),'{','}')
   if group is None:return match[0]
   parts[name]=body[m.start():group[1]];ranges.append((m.start(),group[1]))
  rest=body
  for a,b in sorted(ranges,reverse=True):rest=rest[:a]+rest[b:]
  if re.sub(r'\\centering\b','',rest).strip():return match[0]
  marker='TWFIGPANELS'+hashlib.sha256(match[0].encode()).hexdigest()[:20]+str(len(groups))
  if marker in source:raise ConversionError('Reserved figure marker appears in source')
  groups[marker]=len(pictures)
  primary='\\begin{figure}\n'+pictures[0][0]+'\n'+parts['caption']+parts['label']+'\n\\end{figure}'
  return primary+'\n\n'+marker+'BEGIN\n\n'+' '.join(p[0] for p in pictures[1:])+'\n\n'+marker+'END\n\n'
 return re.sub(r'\\begin\{figure\}(?:\[[^]]*\])?(.*?)\\end\{figure\}',replace,source,flags=re.S),groups

def _restore_panel_figures(ast,groups):
 remaining=set(groups)
 def marker(block):
  return ''.join(x.get('c','') for x in block.get('c',[]) if x.get('t')=='Str') if isinstance(block,dict) and block.get('t')=='Para' else ''
 def walk(obj):
  if isinstance(obj,list):
   i=0
   while i<len(obj):
    token=marker(obj[i]);key=token[:-5] if token.endswith('BEGIN') else ''
    if key in groups:
     if i==0:raise ConversionError('Panel figure has no primary image')
     end=next((j for j in range(i+1,len(obj)) if marker(obj[j])==key+'END'),None)
     primary=obj[i-1]
     if end is None or primary.get('t')!='Para':raise ConversionError('Panel figure boundaries were lost')
     first=[x for x in primary['c'] if x.get('t')=='Image']
     tail=[x for block in obj[i+1:end] if block.get('t')=='Para' for x in block['c']]
     panels=[x for x in tail if x.get('t')=='Image']
     if len(first)!=1 or len(panels)+1!=groups[key] or any(x.get('t') not in ('Image','Space','SoftBreak') for x in tail):raise ConversionError('Panel figure contents were not preserved')
     for panel in panels:
      panel['c'][0][0]=first[0]['c'][0][0];panel['c'][1]=copy.deepcopy(first[0]['c'][1]);panel['c'][2][1]='fig:'
      primary['c'] += [{'t':'Space'},panel]
     del obj[i:end+1];remaining.remove(key);continue
    walk(obj[i]);i+=1
  elif isinstance(obj,dict):
   for value in obj.values():walk(value)
 walk(ast)
 if remaining:raise ConversionError('Panel figure markers were not parsed')
 return ast
def pandoc(args,*,cwd=None,input=None):
 crops={};panel_groups={}
 if input and '-f' in args and args[args.index('-f')+1]=='latex' and '-t' in args and args[args.index('-t')+1]=='json':
  from latex_compat import unwrap_layout
  input=unwrap_layout(input)
  from review_syntax import _without_comments
  lexical=_without_comments(input)
  pattern=r'(?:\\twimagecrop\{(-?\d+),(-?\d+),(-?\d+),(-?\d+)\}\s*)?\\includegraphics(?:\[[^]]*\])?\{([^}]+)\}'
  for match in re.finditer(pattern,lexical):crops.setdefault(match[5],[]).append(list(map(int,match.groups()[:4])) if match[1] is not None else None)
  clean=re.sub(r'\\twimagecrop\{-?\d+,-?\d+,-?\d+,-?\d+\}\s*(?=\\includegraphics)', '', lexical)
  if r'\twimagecrop' in clean:raise ConversionError('Picture crop must immediately precede an image')
  if any(any(c is not None for c in values) for values in crops.values()):input=clean
  input,panel_groups=_prepare_panel_figures(input)
 p=subprocess.run([pandoc_path(),*map(str,args)],cwd=cwd,input=input,stdout=subprocess.PIPE,stderr=subprocess.PIPE,encoding='utf8',errors='strict')
 if p.returncode or re.search(r'not found|not supported|could not convert|skipping',p.stderr,re.I):
  raise ConversionError(p.stderr.strip() or f'Pandoc exited {p.returncode}')
 if panel_groups or any(any(c is not None for c in values) for values in crops.values()):
  ast=json.loads(p.stdout)
  if panel_groups:_restore_panel_figures(ast,panel_groups)
  def attach(obj):
   if isinstance(obj,list):
    for child in obj:attach(child)
   elif isinstance(obj,dict):
    if obj.get('t')=='Image':
     target=obj['c'][2];queue=crops.get(target[0],[])
     crop=queue.pop(0) if queue else None
     if crop is not None:
      if crop[0]+crop[2]>=100000 or crop[1]+crop[3]>=100000:raise ConversionError('Picture crop leaves no visible image')
      if target[1] not in ('','fig:'):raise ConversionError('Cropped image has an unsupported title')
      target[1]+='TWIMAGECROP:'+','.join(map(str,crop))
    for value in obj.values():attach(value)
  attach(ast)
  if any(c is not None for values in crops.values() for c in values):raise ConversionError('Cropped image was not parsed by Pandoc')
  return json.dumps(ast,ensure_ascii=False),p.stderr
 return p.stdout,p.stderr

def inspect_template(path):
 """Inventory paragraphs and effective formatting, without interpreting template prose as instructions."""
 with zipfile.ZipFile(path) as z:
  doc=E.fromstring(z.read('word/document.xml'));styles=E.fromstring(z.read('word/styles.xml'))
 body=doc.find('w:body',NS)
 st={e.get(W+'styleId'):e for e in styles.findall('w:style',NS)}
 default=next((k for k,v in st.items() if v.get(W+'default')=='1' and v.get(W+'type')=='paragraph'),None)
 def effective(p,kind):
  out=tag(kind)
  dd=styles.find(f'w:docDefaults/w:{kind}Default/w:{kind}',NS)
  out=merge_props(out,dd)
  sid=p.find('w:pPr/w:pStyle',NS);sid=sid.get(W+'val') if sid is not None else default
  seen=set();chain=[]
  while sid in st and sid not in seen:
   seen.add(sid);s=st[sid];chain.append(s)
   parent=s.find('w:basedOn',NS);sid=parent.get(W+'val') if parent is not None else None
  for s in reversed(chain):out=merge_props(out,s.find('w:'+kind,NS))
  if kind=='pPr':out=merge_props(out,p.find('w:pPr',NS))
  else:
   run=next((r for r in p.findall('w:r',NS) if r.find('w:t',NS) is not None),None)
   if run is not None:out=merge_props(out,run.find('w:rPr',NS))
  return strip_props(out)
 rows=[]
 for i,p in enumerate(body):
  if p.tag!=W+'p':continue
  sid=p.find('w:pPr/w:pStyle',NS);sid=sid.get(W+'val') if sid is not None else default
  sn=st[sid].find('w:name',NS) if sid in st else None
  rows.append({'index':i,'text':text(p),'style_id':sid,'style_name':sn.get(W+'val') if sn is not None else '',
               'pPr':E.tostring(effective(p,'pPr'),encoding='unicode'),'rPr':E.tostring(effective(p,'rPr'),encoding='unicode')})
 return doc,styles,rows,effective

def adapt(template,out,*,body_start=None,body_end=None,samples=None):
 template=Path(template).resolve();out=Path(out).resolve();out.mkdir(parents=True,exist_ok=True)
 if template==out/'template.docx':raise ConversionError('Use a distinct adapter directory, not the original template path.')
 doc,styles,rows,effective=inspect_template(template);b=doc.find('w:body',NS)
 samples=samples or {};warnings=[]
 def marker(value):
  if value is None:return None
  if isinstance(value,int):
   if not 0<=value<len(b):raise ConversionError('Template body index is out of range')
   return value
  hits=[i for i,p in enumerate(b) if text(p).startswith(value)]
  if len(hits)!=1:raise ConversionError(f'Template boundary must identify one node: {value!r}; got {len(hits)}')
  return hits[0]
 start=marker(body_start)
 if start is None:
  candidates=[r for r in rows if re.search(r'heading\s*1|一级标题|标题\s*1',r['style_name'],re.I) and r['text'].strip()]
  start=candidates[0]['index'] if candidates else None
 end=marker(body_end)
 if end is None:
  endings=[i for i,p in enumerate(b) if text(p).startswith(('本人承诺','作者声明','Author declaration'))]
  end=endings[-1] if endings else len(b)-1
 if start is None:raise ConversionError('Cannot infer replacement body. Specify --body-start or --body-start-index; template prompts are not instructions.')
 if start>=end:raise ConversionError('Template body start must precede end')
 active=[r for r in rows if start<=r['index']<end]
 roles={}
 for role in ROLES:
  candidate=None
  if role in samples:
   candidate=next((r for r in rows if r['index']==samples[role]),None)
   if candidate is None:raise ConversionError(f'Role sample {role} is not a top-level paragraph')
  elif role.startswith('heading'):
   level=role[-1];cn={'1':'一','2':'二','3':'三'}[level]
   candidate=next((r for r in active if re.search(r'heading\s*'+level+r'\b|标题\s*'+level+r'\b|'+cn+'级标题',r['style_name'],re.I)),None)
  elif role=='caption':candidate=next((r for r in active if re.match(r'^[表图]\s*\d+',r['text'])),None)
  elif role=='body':
   candidate=next((r for r in active if re.search(r'正文|body text|^normal$',r['style_name'],re.I) and len(r['text'])>15),None)
  elif role=='bibliography':candidate=next((r for r in active if re.match(r'^\[\d+\]',r['text'])),None)
  if candidate is None:
   if role=='table_text':
    table=next((p for p in b[start:end] if p.tag==W+'tbl'),None)
    p=table.find('.//w:tc/w:p',NS) if table is not None else None
    if p is not None:candidate={'index':None,'style_id':None,'style_name':'table cell sample','pPr':E.tostring(effective(p,'pPr'),encoding='unicode'),'rPr':E.tostring(effective(p,'rPr'),encoding='unicode')}
   if candidate is None:
    fallback=roles.get('body')
    if fallback is None:
     candidate=next((r for r in active if r['text'].strip() and not re.search('heading|标题',r['style_name'],re.I)),active[0])
    else:candidate={'index':fallback['sample_index'],'style_id':fallback['source_style_id'],'style_name':fallback['source_style_name'],'pPr':fallback['pPr'],'rPr':fallback['rPr']}
    warnings.append(f'{role}: no distinct exemplar; inherited body/sample formatting. Review before production.')
  roles[role]={'sample_index':candidate['index'],'source_style_id':candidate['style_id'],'source_style_name':candidate['style_name'],'pPr':candidate['pPr'],'rPr':candidate['rPr']}
 # Encode exemplar properties as conventional writer styles. No per-template parsing code is duplicated.
 for role,(sid,name) in ROLES.items():
  old=next((s for s in styles if s.get(W+'styleId')==sid),None)
  if old is not None:styles.remove(old)
  s=tag('style',type='paragraph',styleId=sid);s.append(tag('name',val=name))
  s.append(E.fromstring(roles[role]['pPr'].encode()));s.append(E.fromstring(roles[role]['rPr'].encode()))
  if role.startswith('heading'):
   p=s.find('w:pPr',NS)
   if p.find('w:outlineLvl',NS) is None:p.append(tag('outlineLvl',val=int(role[-1])-1))
   if p.find('w:keepNext',NS) is None:p.append(tag('keepNext'))
  styles.append(s)
 # Pandoc's FirstParagraph otherwise bypasses the template body indent.
 for sid,name in [('FirstParagraph','First Paragraph'),('Compact','Compact'),('ImageCaption','Image Caption'),('TableCaption','Table Caption')]:
  old=next((s for s in styles if s.get(W+'styleId')==sid),None)
  if old is not None:styles.remove(old)
  s=tag('style',type='paragraph',styleId=sid);s.append(tag('name',val=name));s.append(tag('basedOn',val='Caption' if 'Caption' in sid else 'BodyText'));styles.append(s)
 shutil.copyfile(template,out/'template.docx')
 write_zip(template,out/'reference.docx',{'word/styles.xml':xml_bytes(styles)})
 sect=b.find('w:sectPr',NS)
 pg=sect.find('w:pgSz',NS) if sect is not None else None;mar=sect.find('w:pgMar',NS) if sect is not None else None
 width=int(pg.get(W+'w','11906')) if pg is not None else 11906
 width-=sum(int(mar.get(W+k,'1440')) for k in ['left','right']) if mar is not None else 2880
 table=next((p for p in b[start:end] if p.tag==W+'tbl'),None)
 tablepr=table.find('w:tblPr',NS) if table is not None else None
 profile={'schema':'texword-template-v1','template_origin':str(template),'template_sha256':digest(out/'template.docx'),
          'body_start':start,'body_end':end,'roles':roles,'content_width_twips':width,
          'table_properties':E.tostring(tablepr,encoding='unicode') if tablepr is not None else None,
          'heading1_chinese':bool(re.match(r'^[一二三四五六七八九十]+、',text(b[start]))),'warnings':warnings,
          'paragraph_inventory':[{'index':r['index'],'text':r['text'][:160],'style_name':r['style_name']} for r in rows]}
 dump(out/'profile.json',profile)
 write_adapter(out,profile)
 return profile

def write_adapter(out,profile):
 """Generate actual template-specific Word-building code, not a converter wrapper."""
 config={k:v for k,v in profile.items() if k!='paragraph_inventory'}
 source='''# Generated Word-template adapter. Formatting can be calibrated here.\nfrom pathlib import Path\nfrom word_runtime import TemplateDocument\n\nPROFILE = '''+pprint.pformat(config,width=105,sort_dicts=False)+'''\n\nclass Report(TemplateDocument):\n    def __init__(self):\n        here = Path(__file__).resolve().parent\n        super().__init__(here / "template.docx", PROFILE, here)\n'''
 (out/'template_adapter.py').write_text(source,encoding='utf8')
 write_runtime_link(out)

def write_runtime_link(out):
 """Source adapters share the maintained runtime; compiled bundles copy it."""
 runtime=str((ROOT/'word_runtime.py').resolve())
 source='''# Source adapter runtime link; edit tools/texword/word_runtime.py instead.
import importlib.util as _util
from pathlib import Path as _Path
import sys as _sys
_source = _Path('''+repr(runtime)+''')
_sys.path.insert(0, str(_source.parent))
_spec = _util.spec_from_file_location(__name__, _source)
_runtime = _util.module_from_spec(_spec)
_sys.modules[__name__] = _runtime
_spec.loader.exec_module(_runtime)
globals().update(vars(_runtime))
'''
 (Path(out)/'word_runtime.py').write_text(source,encoding='utf8')

def raw(xml):return {'t':'RawInline','c':['openxml',xml]}
def run(s):
 e=tag('r');t=E.SubElement(e,W+'t');t.set('{http://www.w3.org/XML/1998/namespace}space','preserve');t.text=s
 return E.tostring(e,encoding='unicode')
def bookmark_name(key):return 'TW_'+hashlib.sha1(key.encode()).hexdigest()[:24]
def field(code,value):
 r1=tag('r');r1.append(tag('fldChar',fldCharType='begin'))
 r2=tag('r');it=E.SubElement(r2,W+'instrText');it.set('{http://www.w3.org/XML/1998/namespace}space','preserve');it.text=' '+code+' '
 r3=tag('r');r3.append(tag('fldChar',fldCharType='separate'))
 r4=tag('r');r4.append(tag('fldChar',fldCharType='end'))
 return ''.join(E.tostring(x,encoding='unicode') for x in [r1,r2,r3])+run(str(value))+E.tostring(r4,encoding='unicode')

class AstTransform:
 def __init__(self,profile):
  self.profile=profile;self.targets={};self.citations={};self.next_bookmark=10000;self.count={'table':0,'figure':0,'equation':0};self.heading=[0]*6
  self.active_span=None;self.target_spans={}
 def anchor(self,key,content):
  bid=self.next_bookmark;self.next_bookmark+=1
  return E.tostring(tag('bookmarkStart',id=bid,name=bookmark_name(key)),encoding='unicode')+content+E.tostring(tag('bookmarkEnd',id=bid),encoding='unicode')
 def register(self,key,value,kind):
  if key in self.targets:raise ConversionError(f'Duplicate label: {key}')
  self.targets[key]=(str(value),kind)
  if self.active_span:self.target_spans[key]=self.active_span
 def prepare(self,obj,parent_label=None):
  if isinstance(obj,list):
   for x in obj:self.prepare(x,parent_label)
   return
  if not isinstance(obj,dict):return
  t=obj.get('t');c=obj.get('c')
  if t=='Para':
   images=[x for x in c if x.get('t')=='Image']
   figures=[x for x in images if x['c'][2][1].startswith('fig:')]
   if len(figures)>1 and len(figures)==len(images) and all(x.get('t') in ('Image','Space','SoftBreak') for x in c):
    # Pandoc repeats a LaTeX figure's label/caption on every panel. Keep
    # the panels inline and lower their shared caption exactly once.
    first=figures[0]['c'];key=first[0][0] or parent_label
    if any(x['c'][0][0]!=first[0][0] or x['c'][1]!=first[1] for x in figures[1:]):
     raise ConversionError('Side-by-side figure panels must share one label and caption')
    self.count['figure']+=1;n=self.count['figure'];key=key or 'figure:'+str(n)
    self.register(key,self.profile.get('number_prefix','')+str(n),'figure')
    caption=copy.deepcopy(first[1])
    for image in images:
     image['c'][0][0]='';image['c'][1]=[];image['c'][2][1]=''
    prefix=raw(run(self.profile.get('figure_caption_prefix','图 '))+self.anchor(key,run(self.profile.get('number_prefix',''))+field('SEQ Figure \\* ARABIC',n))+run('　'))
    blocks=[{'t':'Div','c':[['',[],[['custom-style','CaptionedFigure']]],[{'t':'Para','c':c}]]},{'t':'Div','c':[['',[],[['custom-style','ImageCaption']]],[{'t':'Para','c':[prefix]+caption}]]}]
    obj.clear();obj.update({'t':'Div','c':[['',[],[]],blocks]})
    self.prepare(blocks,parent_label);return
  if t in ('RawBlock','Para'):
   marker=c[1] if t=='RawBlock' else ''.join(x.get('c','') for x in c if x.get('t')=='Str')
   m=re.search(r'TWRT_(BEGIN|END)_(s\d+)',marker)
   if m:self.active_span=m.group(2) if m.group(1)=='BEGIN' else None
  if t=='Cite':
   limit=self.profile.get('max_citation_group',3)
   if len(c[0])>limit:raise ConversionError('Each citation group supports at most '+('three' if limit==3 else str(limit))+' references.')
   for ci in c[0]:self.citations.setdefault(ci['citationId'],len(self.citations)+1)
   return
  if t=='Header':
   if 'unnumbered' in c[1][1]:return
   level,attr,inlines=c;self.heading[level-1]+=1
   for i in range(level,len(self.heading)):self.heading[i]=0
   number='.'.join(map(str,self.heading[:level]))
   display=number
   if level==1 and self.profile['heading1_chinese']:
    if self.heading[0]>10:raise ConversionError('Chinese top-level headings above ten need a numbering policy.')
    display='一二三四五六七八九十'[self.heading[0]-1]
   key=attr[0] or 'section:'+number
   self.register(key,display,'section')
   c[2]=[raw(self.anchor(key,run(display))+run('、' if level==1 and self.profile['heading1_chinese'] else ' '))]+inlines
   return
  if t=='Div':
   label=c[0][0] or parent_label
   self.prepare(c[1],label);return
  if t=='Table':
   self.count['table']+=1;n=self.count['table'];key=c[0][0] or parent_label or 'table:'+str(n)
   self.register(key,self.profile.get('number_prefix','')+str(n),'table')
   prefix=raw(run(self.profile.get('table_caption_prefix','表 '))+self.anchor(key,run(self.profile.get('number_prefix',''))+field('SEQ Table \\r '+str(n)+' \\* ARABIC',n))+run('　'))
   if not c[1][1]:c[1][1]=[{'t':'Plain','c':[]}]
   c[1][1][0]['c'].insert(0,prefix)
  if t=='Image' and c[2][1].startswith('fig:'):
   self.count['figure']+=1;n=self.count['figure'];key=c[0][0] or parent_label or 'figure:'+str(n)
   self.register(key,self.profile.get('number_prefix','')+str(n),'figure')
   c[1].insert(0,raw(run(self.profile.get('figure_caption_prefix','图 '))+self.anchor(key,run(self.profile.get('number_prefix',''))+field('SEQ Figure \\* ARABIC',n))+run('　')))
  if t=='Math' and c[0]['t']=='DisplayMath':
   labels=re.findall(r'\\label\{([^}]+)\}',c[1])
   if len(labels)>1:raise ConversionError('Unsupported: multiple labels in one display equation; split the equations.')
   if labels:
    self.count['equation']+=1;n=self.count['equation'];key=labels[0]
    self.register(key,n,'equation')
    obj['_equation']=(key,n)
  if t=='RawInline' or t=='RawBlock':
   if c[0] in ['latex','tex'] and c[1].strip():raise ConversionError('Unsupported raw LaTeX: '+c[1][:160])
  for key,value in obj.items():
   if key!='_equation':self.prepare(value,parent_label)
 def finish(self,obj):
  if isinstance(obj,list):return [self.finish(x) for x in obj]
  if not isinstance(obj,dict):return obj
  t=obj.get('t');c=obj.get('c')
  if t=='Cite':
   s=''
   for ci in c[0]:
    key=ci['citationId'];n=self.citations[key]
    if ci['citationPrefix'] or ci['citationSuffix']:raise ConversionError('Unsupported citation prefix/suffix; write explanatory text outside cite.')
    if self.profile.get('citation_style')=='superscript':
     if s:s+=run(',')
     s+=field('REF '+bookmark_name('bib:'+key)+' '+chr(92)+'h',n)
    else:s+=run('[')+field('REF '+bookmark_name('bib:'+key)+' '+chr(92)+'h',n)+run(']')
   if self.profile.get('citation_style')=='superscript':
    tree=E.fromstring(('<w:p xmlns:w="'+NS['w']+'">'+s+'</w:p>').encode())
    for r in tree:
     pr=r.find(W+'rPr')
     if pr is None:pr=tag('rPr');r.insert(0,pr)
     pr.append(tag('vertAlign',val='superscript'))
    s=''.join(E.tostring(r,encoding='unicode') for r in tree)
   return raw(s)
  if t=='Link' and c[2][0].startswith('#'):
   key=c[2][0][1:]
   if key not in self.targets:raise ConversionError('Unresolved reference: '+key)
   value,kind=self.targets[key];xml=field('REF '+bookmark_name(key)+' \\h',value)
   if dict(c[0][2]).get('reference-type')=='eqref':xml=run('(')+xml+run(')')
   return raw(xml)
  if t=='Div' and c[0][0].startswith('ref-'):
   key=c[0][0][4:]
   if key in self.citations:
    first=next((x for x in c[1] if x['t'] in ['Para','Plain']),None)
    if first is None:raise ConversionError('Bibliography entry has no paragraph: '+key)
    first['c'].insert(0,raw(run('[')+self.anchor('bib:'+key,run(str(self.citations[key])) )+run('] ')))
  if t=='Para':
   # A display equation needs its own paragraph, independent of blank lines in the TeX source.
   blocks=[];current=[]
   for inline in c:
    if inline.get('t')=='Math' and inline['c'][0]['t']=='DisplayMath':
     if any(x['t'] not in ['Space','SoftBreak'] for x in current):blocks.append({'t':'Para','c':current})
     current=[];eq=inline.pop('_equation',None)
     if eq is None:
      labels=re.findall(r'\\label\{([^}]+)\}',inline['c'][1])
      if labels:eq=(labels[0],self.targets[labels[0]][0])
     inline['c'][1]=re.sub(r'\\label\{[^}]+\}','',inline['c'][1]).strip();items=[inline]
     if eq:
      key,n=eq;items+=[raw(run('　(')+self.anchor(key,field('SEQ Equation \\* ARABIC',n))+run(')'))]
     blocks.append({'t':'Para','c':items})
    else:current.append(inline)
   if blocks:
    if any(x['t'] not in ['Space','SoftBreak'] for x in current):blocks.append({'t':'Para','c':current})
    return {'t':'Div','c':[['',[],[]],[{'t':'Para','c':[self.finish(x) for x in p['c']]} for p in blocks]]}
  return {k:self.finish(v) for k,v in obj.items() if k!='_equation'}

def transform_ast(ast,profile,*,cwd,bibliography=None,proposal_citations=()):
 tr=AstTransform(profile);tr.prepare(ast['blocks'])
 # Revision proposals are not ordinary document text until accepted. Add only
 # their missing citation IDs to citeproc's input so its normal bibliography
 # renderer creates the same native targets used by ordinary \cite commands.
 synthetic_id=None
 missing=[key for key in dict.fromkeys(proposal_citations) if key not in tr.citations]
 if missing:
  synthetic_id='TWCITEPROPOSAL'+hashlib.sha256('\0'.join(missing).encode('utf8')).hexdigest()[:24]
  cite_inlines=[]
  for key in missing:
   citation={'citationId':key,'citationPrefix':[],'citationSuffix':[],
             'citationMode':{'t':'NormalCitation'},'citationNoteNum':0,'citationHash':0}
   cite_inlines.append({'t':'Cite','c':[[citation],[{'t':'Str','c':synthetic_id}] ]})
  ast['blocks'].append({'t':'Div','c':[[synthetic_id,[],[]],[{'t':'Para','c':cite_inlines}]]})
  tr.prepare(ast['blocks'][-1])
 if tr.citations:
  args=['-f','json','-t','json','--citeproc','--csl',profile.get('csl',ROOT/'numeric.csl')]
  if bibliography:args+=['--bibliography',bibliography]
  out,_=pandoc(args,cwd=cwd,input=json.dumps(ast,ensure_ascii=False));ast=json.loads(out)
  if synthetic_id:
   ast['blocks']=[b for b in ast['blocks']
                  if not (b.get('t')=='Div' and b.get('c',[[]])[0][0]==synthetic_id)]
  # Equation labels remain in math source until the final transformation, so citeproc may discard private metadata safely.
 transformed=tr.finish(ast)
 return transformed,tr

def compile_code(*args,**kwargs):
 from codegen import compile_code as implementation
 return implementation(*args,**kwargs)


def main():
 p=argparse.ArgumentParser(description=__doc__);subs=p.add_subparsers(dest='command',required=True)
 a=subs.add_parser('adapt');a.add_argument('template',type=Path);a.add_argument('out',type=Path)
 a.add_argument('--body-start');a.add_argument('--body-end');a.add_argument('--body-start-index',type=int);a.add_argument('--body-end-index',type=int)
 a.add_argument('--sample',action='append',default=[],help='role=body-child-index; e.g. body=22')
 i=subs.add_parser('inspect');i.add_argument('template',type=Path);i.add_argument('--out',type=Path)
 c=subs.add_parser('compile');c.add_argument('adapter',type=Path);c.add_argument('source',type=Path);c.add_argument('output',type=Path);c.add_argument('--bibliography',type=Path);c.add_argument('--overwrite',action='store_true')
 c.add_argument('--roundtrip',action='store_true',help='Embed stable source blocks and baselines for Word-to-LaTeX import')
 s=subs.add_parser('sync',help='Map saved Word edits back to LaTeX; default is a preview')
 s.add_argument('word',type=Path);s.add_argument('source',type=Path);s.add_argument('--output',type=Path);s.add_argument('--apply',action='store_true');s.add_argument('--report',type=Path)
 a=subs.add_parser('adopt',help='Import an existing Word into an editable plain-LaTeX project')
 a.add_argument('word',type=Path);a.add_argument('project',type=Path);a.add_argument('--output-word',type=Path,required=True);a.add_argument('--body-start',type=int);a.add_argument('--body-end',type=int)
 args=p.parse_args()
 try:
  if args.command=='adapt':
   samples={s.split('=',1)[0]:int(s.split('=',1)[1]) for s in args.sample}
   result=adapt(args.template,args.out,body_start=args.body_start_index if args.body_start_index is not None else args.body_start,body_end=args.body_end_index if args.body_end_index is not None else args.body_end,samples=samples)
   result={k:v for k,v in result.items() if k not in ['roles','paragraph_inventory','table_properties']}
  elif args.command=='inspect':
   _,_,rows,_=inspect_template(args.template);result=[{k:v for k,v in r.items() if k not in ['pPr','rPr']} for r in rows]
   if args.out:dump(args.out,result);result={'inventory':str(args.out),'paragraphs':len(rows)}
  elif args.command=='adopt':
   from adoption import create_project,write_workflow
   result=create_project(args.word,args.project,body_start=args.body_start,body_end=args.body_end);write_workflow(args.project,args.output_word)
  elif args.command=='sync':
   from roundtrip import sync_word
   result=sync_word(args.word,args.source,output=args.output,apply=args.apply,report=args.report)
  else:result=compile_code(args.adapter,args.source,args.output,bibliography=args.bibliography,overwrite=args.overwrite,roundtrip=args.roundtrip)
  print(json.dumps(result,ensure_ascii=False,indent=2))
 except (ConversionError,ValueError) as e:p.exit(2,str(e)+'\n')

if __name__=='__main__':main()
