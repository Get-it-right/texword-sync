"""Runtime for generated Word-building Python programs. No LaTeX reader or Pandoc."""
from __future__ import annotations
import argparse,copy,hashlib,json,re
from pathlib import Path
from lxml import etree as E
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import parse_xml,OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import Emu

NS={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main','m':'http://schemas.openxmlformats.org/officeDocument/2006/math'}
W='{'+NS['w']+'}'
def text(e):return ''.join(E._Element.xpath(e,'.//w:t/text() | .//m:t/text()',namespaces=NS))
def node(local,**attrs):
 e=OxmlElement('w:'+local)
 for k,v in attrs.items():e.set(W+k,str(v))
 return e
def tokens_text(tokens):return ''.join(v.get('text','') for v in tokens)

def normalize_hidden_math_limits(root):
 """Remove Pandoc's zero-width payload only from explicitly hidden n-ary limits."""
 m='{'+NS['m']+'}';removed=0
 for operator in root.iter(m+'nary'):
  for name in ('sub','sup'):
   hidden=operator.find('m:naryPr/m:'+name+'Hide',NS)
   slot=operator.find('m:'+name,NS)
   if hidden is None or hidden.get(m+'val','1') not in ('1','true') or slot is None:continue
   payload=''.join(t.text or '' for t in slot.iter(m+'t'))
   if not payload or set(payload)!={'\u200b'}:continue
   if any(run.tag!=m+'r' or any(c.tag not in (m+'rPr',W+'rPr',m+'t') for c in run) for run in slot):continue
   for run in list(slot):slot.remove(run)
   removed+=1
 return removed

def layout_blank(element):
 """An empty paragraph, excluding manual breaks, section boundaries and objects."""
 if element.tag!=W+'p' or element.find('w:pPr/w:sectPr',NS) is not None:return False
 return all(child.tag==W+'pPr' or child.tag in (W+'bookmarkStart',W+'bookmarkEnd',W+'proofErr') or
            (child.tag==W+'r' and all(c.tag==W+'rPr' or (c.tag==W+'t' and not (c.text or '').strip()) for c in child)) for child in element)

def layout_flow(body):
 """Read logical body blocks through source controls, leaving tables atomic."""
 for element in body:
  if element.tag==W+'sdt':
   content=element.find('w:sdtContent',NS)
   if content is not None:yield from layout_flow(content)
  else:yield element

class TemplateDocument:
 def __init__(self,template,profile,assets):
  self.template=Path(template).resolve();self.profile=copy.deepcopy(profile);self.assets=Path(assets)
  if hashlib.sha256(self.template.read_bytes()).hexdigest()!=profile['template_sha256']:raise ValueError('Template changed; regenerate template adapter')
  self.doc=Document(self.template);self.body=self.doc.element.body
  self._originals=[copy.deepcopy(e) for e in self.body]
  self.prefix=[text(x) for x in list(self.body)[:profile['body_start']]]
  self.suffix=[copy.deepcopy(x) for x in list(self.body)[profile['body_end']:] if x.tag!=W+'sectPr']
  for x in list(self.body)[profile['body_start']:]:
   if x.tag!=W+'sectPr':self.body.remove(x)
  self._install_styles()
  self.source_map=None;self._source_open=None;self.reviews=[];self._review_ids={};self._next_original_bookmark=700000
  self._native_table_widths={}
  self._native_spacing=None
 def _capture_spacing(self,doc,manifest):
  from roundtrip_model import signature,doc_relationships,math_model
  from tracked_reviews import project_revisions
  project_revisions(doc,manifest,project_accepted=True)
  controls=[i for i,e in enumerate(doc.element.body) if e.tag==W+'sdt']
  if not controls:return
  flow=list(layout_flow(list(doc.element.body)[controls[0]:controls[-1]+1]))
  rels=doc_relationships(doc);solid=[];gaps=[];blank=[]
  for element in flow:
   if layout_blank(element):blank.append(copy.deepcopy(element));continue
   if solid:gaps.append(blank)
   solid.append((signature([element],rels),element.tag,
                 [math_model(m) for m in element.findall('.//m:oMath',NS)]));blank=[]
  self._native_spacing=(solid,gaps)
 def _restore_spacing(self):
  if self._native_spacing is None:return 0
  import difflib
  from roundtrip_model import signature,doc_relationships
  controls=[i for i,e in enumerate(self.body) if e.tag==W+'sdt']
  if not controls:return 0
  flow=list(layout_flow(list(self.body)[controls[0]:controls[-1]+1]));solid=[];gaps=[];blank=[]
  rels=doc_relationships(self.doc)
  for element in flow:
   if layout_blank(element):blank.append(element);continue
   if solid:gaps.append(blank)
   solid.append((signature([element],rels),element));blank=[]
  old,old_gaps=self._native_spacing;mapping={}
  for op,a,b,c,d in difflib.SequenceMatcher(None,[x[0] for x in old],[x[0] for x in solid],autojunk=False).get_opcodes():
   if op=='equal':mapping.update({i:c+i-a for i in range(a,b)})
   elif op=='replace' and b-a==d-c==1 and old[a][1]==solid[c][1].tag:mapping[a]=c
   elif op=='replace' and b-a==d-c:
    # Word may add equivalent defaults to several adjacent equations. Match
    # those against their saved math baselines, without guessing actual edits
    # or paragraph splits/merges from position alone.
    for i,j in zip(range(a,b),range(c,d)):
     if (old[i][1]==solid[j][1].tag and old[i][2] and
         signature([solid[j][1]],rels,math_baselines=old[i][2])==old[i][0]):
      mapping[i]=j
  restored=0
  for i,saved in enumerate(old_gaps):
   left=mapping.get(i);right=mapping.get(i+1)
   if left is None or right!=left+1:continue
   for element in gaps[left]:element.getparent().remove(element)
   anchor=solid[left][1]
   for element in saved:
    copied=copy.deepcopy(element);anchor.addnext(copied);anchor=copied
   restored+=1
  return restored
 def _numbered_tables(self,doc,references):
  """Associate tables with stable caption targets, never their current ordinal."""
  def blocks(parent):
   for element in parent:
    if element.tag==W+'sdt':
     content=element.find('w:sdtContent',NS)
     if content is not None:yield from blocks(content)
    else:yield element
  pending=None;seen=set()
  for element in blocks(doc.element.body):
   if element.tag==W+'p':
    names=[b.get(W+'name') for b in element.iter(W+'bookmarkStart')
           if references.get(b.get(W+'name'),{}).get('kind')=='table']
    if len(names)>1:raise ValueError('Table caption has multiple semantic targets')
    if names:pending=names[0]
    elif text(element).strip():pending=None
   elif element.tag==W+'tbl':
    if pending:
     if pending in seen:raise ValueError('Duplicate native table layout target')
     seen.add(pending);yield pending,element
    pending=None
 def _restore_table_widths(self):
  restored=[];skipped=[]
  if not self._native_table_widths or self.source_map is None:return restored,skipped
  for key,table in self._numbered_tables(self.doc,self.source_map.get('references',{})):
   layout=self._native_table_widths.get(key)
   if layout is None:continue
   grid=table.find('w:tblGrid',NS);saved=layout['grid']
   if grid is None or len(grid)!=len(saved):
    skipped.append({'target':key,'reason':'column_count_changed'});continue
   widths=[int(column.get(W+'w')) for column in saved]
   if not widths or any(w<=0 for w in widths):raise ValueError('Native table has invalid column widths')
   cells=[]
   for row in table.findall('w:tr',NS):
    before=row.find('w:trPr/w:gridBefore',NS);offset=int(before.get(W+'val')) if before is not None else 0
    for cell in row.findall('w:tc',NS):
     span=cell.find('w:tcPr/w:gridSpan',NS);span=int(span.get(W+'val')) if span is not None else 1
     if span<1 or offset<0 or offset+span>len(widths):raise ValueError('Table cell exceeds native column grid')
     cells.append((cell,sum(widths[offset:offset+span])));offset+=span
   table.replace(grid,copy.deepcopy(saved))
   properties=table.find('w:tblPr',NS)
   if properties is None:properties=node('tblPr');table.insert(0,properties)
   for name in ('tblW','tblLayout'):
    prior=properties.find('w:'+name,NS)
    if prior is not None:properties.remove(prior)
    if layout[name] is not None:properties.append(copy.deepcopy(layout[name]))
   for cell,width in cells:
    properties=cell.find('w:tcPr',NS)
    if properties is None:properties=node('tcPr');cell.insert(0,properties)
    prior=properties.find('w:tcW',NS)
    if prior is not None:properties.remove(prior)
    properties.append(node('tcW',w=width,type='dxa'))
   restored.append(key)
  return restored,skipped
 def preserve_template_regions(self,path):
  """Carry native cover/suffix and numbered table widths through a rebuild."""
  current=Document(path)
  from roundtrip_model import read_manifest
  from native_comparison import project_reader
  project_reader(current,read_manifest(current))
  children=list(current.element.body);mapped=[]
  for i,e in enumerate(children):
   tag=e.find('w:sdtPr/w:tag',NS) if e.tag==W+'sdt' else None
   if tag is not None and tag.get(W+'val','').startswith('texword:'):mapped.append(i)
  if not mapped:raise ValueError('Template passthrough requires mapped body sections')
  ignored={W+'bookmarkStart',W+'bookmarkEnd',W+'proofErr'}
  prefix=[e for e in children[:mapped[0]] if e.tag not in ignored]
  suffix=[e for e in children[mapped[-1]+1:] if e.tag not in ignored]
  from roundtrip_model import read_manifest,signature,doc_relationships
  manifest=read_manifest(current)
  # Citeproc appends an external .bib bibliography after the source controls.
  # It is generated content, not a template footer to copy a second time.
  # Accept only the complete, unchanged mapped citation registry; bibliography
  # edits still require the normal import path and cannot disappear here.
  tail=[e for e in suffix if e.tag!=W+'sectPr']
  if tail and re.search(r'\\bibliography\s*\{',manifest['source']):
   expected=manifest.get('outside_nodes',[])[manifest.get('outside_prefix_count',0):]
   if suffix[-1].tag==W+'sectPr' and len(expected)==len(tail)+1:expected=expected[:-1]
   if len(expected)==len(tail) and len(tail)>1:
    from section_mapping import _restore_saved_emphasis
    normalized=copy.deepcopy(tail)
    wrapper=E.Element(W+'body')
    for element in normalized:wrapper.append(element)
    _restore_saved_emphasis(current,normalized,{'units':[{'node_signatures':expected}]},doc_relationships(current))
    same=[signature([e],doc_relationships(current)) for e in normalized]==expected
    names={name for e in tail[1:] for name in E._Element.xpath(e,'.//w:bookmarkStart/@w:name',namespaces=NS)}
    citations={name for name,info in manifest.get('references',{}).items() if info['kind']=='citation'}
    heading=text(tail[0])==self.profile.get('bibliography_heading','参考文献')
    styles={s.get(W+'styleId'):s.find(W+'name') for s in current.styles.element}
    def bibliography_entry(element):
     style=element.find('w:pPr/w:pStyle',NS)
     if element.tag!=W+'p' or style is None:return False
     sid=style.get(W+'val');name=styles.get(sid)
     return sid=='Bibliography' or (name is not None and name.get(W+'val','').lower()=='bibliography')
    entries=all(bibliography_entry(e) for e in tail[1:])
    if same and heading and entries and names==citations:
     suffix=[e for e in suffix if e.tag==W+'sectPr']
  if sum(not layout_blank(e) for e in prefix)!=sum(not layout_blank(e) for e in self._originals[:self.profile['body_start']]):raise ValueError('Template prefix structure changed')
  if sum(not layout_blank(e) for e in suffix if e.tag!=W+'sectPr')!=sum(not layout_blank(e) for e in self.suffix):raise ValueError('Template suffix structure changed')
  references=manifest.get('references',{})
  for key,table in self._numbered_tables(current,references):
   grid=table.find('w:tblGrid',NS)
   if grid is None:raise ValueError('Native table has no column grid')
   self._native_table_widths[key]={'grid':copy.deepcopy(grid),**{
    name:copy.deepcopy(table.find('w:tblPr/w:'+name,NS)) for name in ('tblW','tblLayout')}}
  rns='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
  from docx.opc.packuri import PackURI
  from io import BytesIO
  package=self.doc.part.package;allocated={str(p.partname) for p in package.parts};copied={}
  def import_part(part):
   if part in copied:return copied[part]
   if part.content_type.startswith('image/'):
    result=package.get_or_add_image_part(BytesIO(part.blob));copied[part]=result;return result
   if not part.content_type.endswith('+xml'):raise ValueError('Unsupported native template related part: '+part.content_type)
   index=1;digest=hashlib.sha256(part.blob).hexdigest()[:16]
   name='/word/tw_native_'+digest+'_'+str(index)+'.xml'
   while name in allocated:
    index+=1;name='/word/tw_native_'+digest+'_'+str(index)+'.xml'
   allocated.add(name);result=type(part).load(PackURI(name),part.content_type,part.blob,package);copied[part]=result
   for rel in part.rels.values():
    target=rel.target_ref if rel.is_external else import_part(rel.target_part)
    result.load_rel(rel.reltype,target,rel.rId,rel.is_external)
   return result
  prefix=[copy.deepcopy(e) for e in prefix];suffix=[copy.deepcopy(e) for e in suffix]
  for e in prefix+suffix:
   for n in e.iter():
    for key,rid in n.attrib.items():
     if not key.startswith(rns):continue
     original=current.part.rels.get(rid)
     if original is None:raise ValueError('Native template relationship is missing: '+rid)
     target=original.target_ref if original.is_external else import_part(original.target_part)
     n.set(key,self.doc.part.relate_to(target,original.reltype,is_external=original.is_external))
  for e in list(self.body)[:self.profile['body_start']]:self.body.remove(e)
  for i,e in enumerate(prefix):self.body.insert(i,copy.deepcopy(e))
  section=next((e for e in suffix if e.tag==W+'sectPr'),None)
  if section is not None:
   old=self.body.find('w:sectPr',NS)
   if old is not None:self.body.replace(old,copy.deepcopy(section))
  self.prefix=[text(e) for e in prefix]
  self.suffix=[copy.deepcopy(e) for e in suffix if e.tag!=W+'sectPr']
  self.profile['body_start']=len(prefix)
  self._capture_spacing(current,read_manifest(current))
 def set_reviews(self,definitions):
  if not hasattr(self.doc,'comments'):raise ValueError('Native review comments require python-docx >= 1.2')
  for spec in definitions:
   d=copy.deepcopy(spec)
   comment=self.doc.comments.add_comment(text=d['rendered_text'],author=d.get('author','Codex'),initials='')
   d['native_id']=str(comment.comment_id);self._review_ids[d['compiled_id']]=d['native_id'];self.reviews.append(d)
 def enable_roundtrip(self):
  self.source_map=json.loads((self.assets/'source_map.json').read_text(encoding='utf8'))
 def original_blocks(self,indices,*,anchors=None):
  elements=[copy.deepcopy(self._originals[i]) for i in indices]
  for a in anchors or []:
   self._original_anchor(elements,a)
  for e in elements:self.body.insert(len(self.body)-1,e)
 def _original_anchor(self,elements,spec):
  # Bind to the existing printed number, keeping all original run properties.
  name=spec['name'];value=spec['value'];bid=self._next_original_bookmark;self._next_original_bookmark+=1
  if spec['kind'] in ('table','figure'):
   prefix='表' if spec['kind']=='table' else '图'
   for element in elements:
    if element.tag!=W+'p':continue
    texts=element.findall('.//w:t',NS);caption=''.join(t.text or '' for t in texts)
    number=re.match(r'^\s*'+prefix+r'\s*(\d+)',caption)
    if number:
     a,b=number.span(1);offset=0;inserted=False
     for t in texts:
      old=t.text or '';stop=offset+len(old)
      if offset<b and stop>a:
       t.text=old[:max(0,a-offset)]+(value if not inserted else '')+old[max(0,b-offset):]
       inserted=True
      offset=stop
     break
  for element in elements:
   for t in element.findall('.//w:t',NS):
    txt=t.text or '';match=re.search(re.escape(value),txt)
    if not match:continue
    r=t.getparent()
    if r.tag!=W+'r':continue
    parent=r.getparent();at=parent.index(r);pieces=[]
    for part in [txt[:match.start()],txt[match.start():match.end()],txt[match.end():]]:
     if part:
      rr=copy.deepcopy(r)
      for child in list(rr):
       if child.tag!=W+'rPr':rr.remove(child)
      tt=node('t');tt.set('{http://www.w3.org/XML/1998/namespace}space','preserve');tt.text=part;rr.append(tt);pieces.append((part,rr))
    parent.remove(r);done=False
    for part,rr in pieces:
     if part==value and not done:
      parent.insert(at,node('bookmarkStart',id=bid,name=name));at+=1;parent.insert(at,rr);at+=1;parent.insert(at,node('bookmarkEnd',id=bid));at+=1;done=True
     else:parent.insert(at,rr);at+=1
    return
  # An unnumbered source has no target. Refuse a fabricated reference value.
  raise ValueError('Cannot bind original reference target '+name+' to printed value '+value)
 def begin_source(self,span_id):
  if self.source_map is None or self._source_open is not None:raise ValueError('Invalid source-map boundary')
  self._source_open=(span_id,list(self.body))
 def end_source(self):
  if self._source_open is None:raise ValueError('Missing source-map begin')
  span_id,before=self._source_open
  children=[x for x in self.body if x not in before and x.tag!=W+'sectPr']
  if not children:raise ValueError('Empty mapped source block: '+span_id)
  sdt=node('sdt');pr=node('sdtPr');pr.append(node('tag',val='texword:'+span_id))
  # Keep the source key internal; ordinary editing should show only the report.
  w15='http://schemas.microsoft.com/office/word/2012/wordml'
  appearance=E.SubElement(pr,'{'+w15+'}appearance',nsmap={'w15':w15});appearance.set('{'+w15+'}val','hidden')
  pr.append(node('id',val=900000+int(span_id[1:])));sdt.append(pr);content=node('sdtContent');sdt.append(content)
  index=self.body.index(children[0])
  for x in children:content.append(x)
  self.body.insert(index,sdt);self._source_open=None
 def _install_styles(self):
  mapping={'body':('BodyText','Body Text'),'heading1':('Heading1','heading 1'),'heading2':('Heading2','heading 2'),'heading3':('Heading3','heading 3'),'caption':('Caption','Caption'),'table_text':('TWTableText','TW Table Text'),'bibliography':('Bibliography','Bibliography')}
  for role,(sid,name) in mapping.items():
   old=next((s for s in self.doc.styles.element if s.get(W+'styleId')==sid),None)
   if old is not None:self.doc.styles.element.remove(old)
   st=node('style',type='paragraph',styleId=sid);st.append(node('name',val=name))
   for prop in ['pPr','rPr']:st.append(parse_xml(self.profile['roles'][role][prop]))
   if role.startswith('heading'):
    pp=st.find('w:pPr',NS)
    outline=pp.find('w:outlineLvl',NS)
    if outline is None:pp.append(node('outlineLvl',val=int(role[-1])-1))
    else:outline.set(W+'val',str(int(role[-1])-1))
    if pp.find('w:keepNext',NS) is None:pp.append(node('keepNext'))
   self.doc.styles.element.append(st)
  for sid,name,base in [('FirstParagraph','First Paragraph','BodyText'),('Compact','Compact','BodyText'),('TableCaption','Table Caption','Caption'),('ImageCaption','Image Caption','Caption')]:
   old=next((s for s in self.doc.styles.element if s.get(W+'styleId')==sid),None)
   if old is not None:self.doc.styles.element.remove(old)
   s=node('style',type='paragraph',styleId=sid);s.append(node('name',val=name));s.append(node('basedOn',val=base));self.doc.styles.element.append(s)
 def _append_tokens(self,p,tokens):
  for item in tokens:
   if 'text' in item:
    r=node('r')
    if item.get('rPr'):r.append(parse_xml(item['rPr']))
    t=E.SubElement(r,W+'t');t.set('{http://www.w3.org/XML/1998/namespace}space','preserve');t.text=item['text'];p.append(r)
   elif 'field' in item:
    value=item['field'];a=node('r');a.append(node('fldChar',fldCharType='begin'));p.append(a)
    a=node('r');i=E.SubElement(a,W+'instrText');i.set('{http://www.w3.org/XML/1998/namespace}space','preserve');i.text=' '+value['code']+' ';p.append(a)
    a=node('r');a.append(node('fldChar',fldCharType='separate'));p.append(a)
    self._append_tokens(p,[{'text':value['value']}])
    a=node('r');a.append(node('fldChar',fldCharType='end'));p.append(a)
   elif 'math' in item:p.append(parse_xml(item['math']))
   elif 'bookmark_start' in item:p.append(node('bookmarkStart',**item['bookmark_start']))
   elif 'bookmark_end' in item:p.append(node('bookmarkEnd',id=item['bookmark_end']))
   elif 'review_start' in item:p.append(node('commentRangeStart',id=self._review_ids[item['review_start']]))
   elif 'review_end' in item:
    cid=self._review_ids[item['review_end']];p.append(node('commentRangeEnd',id=cid))
    r=node('r');pr=node('rPr');pr.append(node('rStyle',val='CommentReference'));r.append(pr);r.append(node('commentReference',id=cid));p.append(r)
   elif 'break' in item:
    r=node('r');r.append(node('br',**({'type':item['break']} if item['break'] else {})));p.append(r)
   elif 'tab' in item:
    r=node('r');r.append(node('tab'));p.append(r)
   elif 'hyperlink' in item:
    spec=item['hyperlink'];link=node('hyperlink')
    if spec.get('anchor'):link.set(W+'anchor',spec['anchor'])
    else:
     rid=self.doc.part.relate_to(spec['url'],RT.HYPERLINK,is_external=True)
     link.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id',rid)
    self._append_tokens(link,spec['tokens']);p.append(link)
   elif 'image' in item:
    from docx.text.paragraph import Paragraph
    im=item['image'];r=Paragraph(p,self.doc._body).add_run()
    picture=r.add_picture(str(self.assets/im['path']),width=Emu(im['width']),height=Emu(im['height']))
    if im.get('crop'):
     fill=picture._inline.graphic.graphicData.pic.blipFill
     rect=fill.find('{http://schemas.openxmlformats.org/drawingml/2006/main}srcRect')
     if rect is None:
      rect=OxmlElement('a:srcRect');fill.insert(1,rect)
     for key,value in zip(('l','t','r','b'),im['crop']):rect.set(key,str(value))
   else:raise ValueError('Unsupported generated token: '+str(item))
 def _paragraph(self,tokens,style='BodyText',pPr=None):
  p=node('p');pp=parse_xml(pPr) if pPr else node('pPr')
  ps=pp.find('w:pStyle',NS)
  if ps is None:ps=node('pStyle');pp.insert(0,ps)
  ps.set(W+'val',style);p.append(pp);self._append_tokens(p,tokens)
  return p
 def paragraph(self,tokens,*,style='BodyText',pPr=None):
  p=self._paragraph(tokens,style,pPr);self.body.insert(len(self.body)-1,p);return p
 def heading(self,level,tokens):return self.paragraph(tokens,style='Heading'+str(level))
 def equation(self,tokens):
  normalized=[{'tab':True}]
  for token in tokens:
   if 'math' in token:
    root=parse_xml(token['math'])
    if E.QName(root).localname=='oMathPara':
     for math in root:
      if E.QName(math).localname=='oMath':normalized.append({'math':E.tostring(math,encoding='unicode')})
    else:normalized.append(token)
    normalized.append({'tab':True})
   elif 'text' in token and token['text'].startswith('\u3000('):normalized.append({'text':token['text'].lstrip()})
   else:normalized.append(token)
  pp=node('pPr');pp.append(node('ind',left=0,firstLine=0));tabs=E.SubElement(pp,W+'tabs')
  tabs.append(node('tab',val='center',pos=self.profile['content_width_twips']//2))
  tabs.append(node('tab',val='right',pos=self.profile['content_width_twips']))
  return self.paragraph(normalized,style='BodyText',pPr=E.tostring(pp,encoding='unicode'))
 def caption(self,tokens,*,style='Caption'):return self.paragraph(tokens,style=style)
 def bibliography_entry(self,tokens):return self.paragraph(tokens,style='Bibliography')
 def table(self,rows,*,widths=None,properties=None):
  width=self.profile['content_width_twips']
  n=max(sum(c.get('span',1) for c in row) for row in rows)
  ratios=widths if widths and len(widths)==n and sum(widths)>0 else [1]*n
  widths=[int(width*x/sum(ratios)) for x in ratios];widths[-1]+=width-sum(widths)
  table=node('tbl');pr=parse_xml(self.profile['table_properties']) if self.profile.get('table_properties') else node('tblPr')
  for key in ['tblW','tblLayout','tblInd']:
   old=pr.find('w:'+key,NS)
   if old is not None:pr.remove(old)
  pr.append(node('tblW',w=width,type='dxa'));pr.append(node('tblLayout',type='fixed'));table.append(pr)
  if pr.find('w:tblBorders',NS) is None:
   border=node('tblBorders')
   for pos in ['top','left','bottom','right','insideH','insideV']:border.append(node(pos,val='single',sz=4,color='808080'))
   pr.append(border)
  grid=E.SubElement(table,W+'tblGrid')
  for w in widths:grid.append(node('gridCol',w=w))
  for index,row in enumerate(rows):
   tr=E.SubElement(table,W+'tr');tp=E.SubElement(tr,W+'trPr');tp.append(node('cantSplit'))
   if index==0:tp.append(node('tblHeader'))
   offset=0
   for cell in row:
    tc=E.SubElement(tr,W+'tc');cp=E.SubElement(tc,W+'tcPr');span=cell.get('span',1)
    cp.append(node('tcW',w=sum(widths[offset:offset+span]),type='dxa'));offset+=span
    if span>1:cp.append(node('gridSpan',val=span))
    if cell.get('vmerge') is not None:cp.append(node('vMerge',**({'val':cell['vmerge']} if cell['vmerge'] else {})))
    if index==0:cp.append(node('shd',fill='E9EEF4',val='clear'))
    for p in cell['paragraphs']:tc.append(self._paragraph(p['tokens'],'TWTableText'))
    if not cell['paragraphs']:tc.append(self._paragraph([],'TWTableText'))
  self.body.insert(len(self.body)-1,table)
 def finish(self,output,*,overwrite=False):
  output=Path(output).resolve()
  if output.suffix.lower()!='.docx':raise ValueError('Output must use .docx; refusing to overwrite source/code files')
  if output in [self.template,Path(self.profile['template_origin']).resolve()]:raise ValueError('Refusing to overwrite original template')
  if output.exists() and not overwrite:raise ValueError('Output exists; pass --overwrite for a generated file.')
  # Apply after template-specific formatting and before mapping/signature capture.
  preserved_widths,skipped_widths=self._restore_table_widths()
  preserved_spacing=self._restore_spacing()
  # Normalize generated and reused math before capturing source-map baselines.
  normalized_limits=sum(normalize_hidden_math_limits(block)
   for block in list(self.body)[self.profile['body_start']:])
  for p in self.suffix:self.body.insert(len(self.body)-1,copy.deepcopy(p))
  names=self.doc.element.xpath('.//w:bookmarkStart/@w:name');instructions=self.doc.element.xpath('.//w:instrText/text()')
  missing=sorted(set(re.findall(r'REF\s+((?:TW_|BIB_)\w+)',' '.join(instructions)))-set(names))
  if missing:raise ValueError('Unresolved generated cross-references: '+str(missing))
  if len(names)!=len(set(names)):raise ValueError('Duplicate bookmark names')
  actual=[text(x) for x in self.body if x.tag!=W+'sectPr'];suffix=[text(x) for x in self.suffix]
  if actual[:len(self.prefix)]!=self.prefix or (suffix and actual[-len(suffix):]!=suffix):raise ValueError('Template prefix/suffix changed')
  if self.source_map is not None:
   if self._source_open is not None:raise ValueError('Unclosed source-map block')
   from roundtrip_model import attach_manifest,signature,doc_relationships,math_model
   from review_comments import read_comments,review_snapshot
   manifest=copy.deepcopy(self.source_map);manifest['reviews']=copy.deepcopy(self.reviews);manifest['semantic_version']=2
   comments=read_comments(self.doc);rels=doc_relationships(self.doc);outside=[];ids=[]
   for e in self.body:
    tag=e.find('w:sdtPr/w:tag',NS) if e.tag==W+'sdt' else None
    if tag is not None and tag.get(W+'val','').startswith('texword:'):
     sid=tag.get(W+'val').split(':',1)[1];ids.append(sid)
     record=next(s for s in manifest['spans'] if s['id']==sid)
     record['baseline']=signature(list(e.find('w:sdtContent',NS)),rels)
     maths=e.findall('.//m:oMath',NS)
     if len(maths)!=len(record.get('math_sources',[])):raise ValueError('Source math mapping differs from generated Word: '+sid)
     record['math_baselines']=[math_model(m) for m in maths]
     record['review_baseline']=review_snapshot(list(e.find('w:sdtContent',NS)),comments)
    else:outside.append(e)
   if ids!=[s['id'] for s in manifest['spans']]:raise ValueError('Source-map coverage mismatch')
   manifest['outside_baseline']=signature(outside,rels);manifest['outside_reviews']=review_snapshot(outside,comments)
   if manifest.get('granularity')=='section':
    from section_mapping import group_sections
    manifest=group_sections(self.doc,manifest,rels,comments)
   attach_manifest(self.doc,manifest)
  from tracked_reviews import install_revisions
  install_revisions(self.doc,self.reviews,render_tokens=self._append_tokens)
  output.parent.mkdir(parents=True,exist_ok=True);self.doc.save(output)
  report={'status':'passed','output':str(output),'missing_bookmarks':missing,'prefix_preserved':True,'suffix_preserved':True,
          'math_objects':len(self.doc.element.xpath('.//m:oMath')),'cross_reference_fields':len(re.findall(r'REF\s+(?:TW_|BIB_)',' '.join(instructions))),
          'sequence_fields':sum('SEQ ' in s for s in instructions),'tables':len(self.doc.element.xpath('.//w:tbl')),'generation_requires_latex':False,'generation_requires_pandoc':False}
  report['review_comments']=len(self.doc.comments)
  report['normalized_hidden_math_limits']=normalized_limits
  report['preserved_table_widths']=preserved_widths;report['skipped_table_widths']=skipped_widths
  report['preserved_native_spacing_gaps']=preserved_spacing
  report['tracked_replacements']=sum(d['kind']=='revision' for d in self.reviews)
  output.with_suffix('.validation.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf8')
  return report

def output_arguments(default_name):
 parser=argparse.ArgumentParser(description='Run generated Word-building code. No LaTeX or Pandoc needed.')
 parser.add_argument('output',nargs='?',type=Path,default=Path(default_name));parser.add_argument('--overwrite',action='store_true')
 parser.add_argument('--template-regions-from',type=Path,help='Preserve compatible native cover and suffix from this mapped Word document')
 return parser.parse_args()
