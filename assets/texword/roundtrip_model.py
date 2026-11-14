"""Shared, Pandoc-free provenance and content comparison for generated DOCX files."""
from __future__ import annotations
import copy, hashlib, json, re
from lxml import etree as E

NS={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'm':'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
W='{'+NS['w']+'}'; M='{'+NS['m']+'}'; RT_NS='urn:texword:roundtrip:1'

def math_model(e):
 """Word rewrites math run boundaries and adds font/default property nodes on save."""
 local=E.QName(e).localname
 if local=='oMathPara':
  maths=e.findall('m:oMath',NS)
  if len(maths)==1:return math_model(maths[0])
 if not e.tag.startswith(M) or local in ('rPr','ctrlPr'):return None
 if local=='type' and e.get(M+'val')=='bar':return None  # Default fraction bar, omitted by Word.
 if local=='r':
  text=''.join(e.itertext());styles=[]
  if any(c.isalpha() for c in text):
   for name,default in (('sty','i'),('scr','roman'),('nor','0')):
    node=e.find('m:rPr/m:'+name,NS)
    if node is not None:
     value=node.get(M+'val','1' if name=='nor' else default)
     if name=='nor':value='0' if value in ('0','false','off') else '1'
     if value!=default:styles.append([name,value])
  return ['text',text,styles] if styles else ['text',text]
 children=[]
 for child in e:
  value=math_model(child)
  if value is None:continue
  values=[value]
  if value[0]=='text' and value[2:]:
   # Font semantics affect letters. Word can combine letters and punctuation
   # in one run or split them; both serializations must retain one identity.
   values=[]
   for char in value[1]:
    suffix=value[2:] if char.isalpha() else []
    if values and values[-1][2:]==suffix:values[-1][1]+=char
    else:values.append(['text',char]+suffix)
  for value in values:
   if children and value[0]=='text' and children[-1][0]=='text' and value[2:]==children[-1][2:]:children[-1][1]+=value[1]
   else:children.append(value)
 attrs=sorted([E.QName(k).localname,v] for k,v in e.attrib.items())
 if local.endswith('Pr') and not children and not attrs:return None
 return [local,attrs,e.text or '',children]

def sha(value):
 return hashlib.sha256(value.encode('utf8') if isinstance(value,str) else value).hexdigest()

def field_parts(p,*,include_comments=False):
 """Return logical runs; ignore cached results of REF/SEQ, including Word simple fields."""
 result=[]; active=None; depth=0
 for x in p:
  local=E.QName(x).localname
  if local in ('pPr','bookmarkStart','bookmarkEnd','proofErr','permStart','permEnd'):continue
  if local in ('commentRangeStart','commentRangeEnd'):
   if include_comments:result.append(('review_start' if local=='commentRangeStart' else 'review_end',x.get(W+'id'),None))
   continue
  if x.tag==W+'fldSimple':
   result.append(('field',x.get(W+'instr','').strip(),None));continue
  if local in ('hyperlink','sdt'):
   result.extend(field_parts(x.find('w:sdtContent',NS) if local=='sdt' else x,include_comments=include_comments));continue
  if x.tag in (M+'oMath',M+'oMathPara'):
   result.append(('math',copy.deepcopy(x),None));continue
  if x.tag!=W+'r':
   result.append(('unknown',E.tostring(x,encoding='unicode'),None));continue
  fmt=tuple(x.find('w:rPr/w:'+tag,NS) is not None and x.find('w:rPr/w:'+tag,NS).get(W+'val','1') not in ('0','false') for tag in ('b','i'))
  for c in x:
   name=E.QName(c).localname
   if name in ('rPr','lastRenderedPageBreak','commentReference'):continue
   if name=='fldChar':
    kind=c.get(W+'fldCharType')
    if kind=='begin':
     if active is not None:raise ValueError('Nested Word fields cannot be mapped safely')
     active='';depth=1
    elif kind=='end':
     if active is None:raise ValueError('Malformed Word field')
     result.append(('field',active.strip(),None));active=None;depth=0
   elif name=='instrText':
    if active is None:raise ValueError('Word field instruction without begin')
    active+=c.text or ''
   elif active is not None:continue
   elif name=='t':result.append(('text',c.text or '',fmt))
   elif name=='tab':result.append(('tab','',None))
   elif name=='br':result.append(('break',c.get(W+'type',''),None))
   elif name=='drawing':result.append(('drawing',copy.deepcopy(c),None))
   else:result.append(('unknown',E.tostring(c,encoding='unicode'),None))
 if depth:raise ValueError('Unclosed Word field')
 return result

def content_model(elements,relationships):
 """Semantic content, independent of run splits, cached field numbers and pagination."""
 def paragraph(p):
  # Word can serialize the same external link as w:hyperlink or HYPERLINK
  # field runs. Keep the visible caption and URL in the existing native-link
  # model, so old manifests still compare correctly after a Word save.
  local_relationships=relationships
  if any(re.match(r'\s*HYPERLINK\b',x.text or '') for x in p.findall('.//w:instrText',NS)):
   p=copy.deepcopy(p);local_relationships=dict(relationships);children=list(p);i=0
   while i<len(children):
    first=children[i]
    if not any(x.get(W+'fldCharType')=='begin' for x in first.iter(W+'fldChar')):i+=1;continue
    j=i
    while j<len(children) and not any(x.get(W+'fldCharType')=='end' for x in children[j].iter(W+'fldChar')):j+=1
    if j==len(children):break  # field_parts reports malformed fields.
    group=children[i:j+1];instruction=''.join(x.text or '' for c in group for x in c.iter(W+'instrText'))
    match=re.fullmatch(r'\s*HYPERLINK\s+"([^"]+)"\s*(?:\\h\s*)?',instruction)
    if match:
     link=E.Element(W+'hyperlink');rid='texword-normalized-'+str(i)
     link.set('{'+NS['r']+'}id',rid);local_relationships[rid]=match.group(1);cached=False
     for child in group:
      run=copy.deepcopy(child)
      for item in list(run):
       if item.tag==W+'fldChar':
        if item.get(W+'fldCharType')=='separate':cached=True
        elif item.get(W+'fldCharType')=='end':cached=False
        run.remove(item)
       elif item.tag==W+'instrText' or (not cached and item.tag!=W+'rPr'):run.remove(item)
      if any(x.tag!=W+'rPr' for x in run):link.append(run)
     p.insert(p.index(first),link)
     for child in group:p.remove(child)
    i=j+1
  parts=[]
  for kind,value,fmt in field_parts(p):
   if kind=='math':
    value=math_model(value)
   if kind=='drawing':
    ids=E._Element.xpath(value,'.//@r:embed',namespaces=NS)
    geometry=[]
    for item in value.iter():
     local=E.QName(item).localname
     if local=='extent':geometry.append(['extent',item.get('cx'),item.get('cy')])
     elif local=='srcRect':
      crop=sorted((k,int(v)) for k,v in item.attrib.items() if int(v))
      if crop:geometry.append(['crop',crop])
     elif local=='xfrm':
      attrs=sorted((k,int(v)) for k,v in item.attrib.items() if int(v))
      if attrs:geometry.append(['transform',attrs])
    value={'assets':[relationships.get(i,'MISSING_IMAGE') for i in ids],'geometry':geometry}
   if kind=='field':value=' '.join(value.split())
   if kind=='text' and parts and parts[-1][0]=='text' and parts[-1][2]==fmt:parts[-1][1]+=value
   else:parts.append([kind,value,fmt])
  links=[local_relationships.get(x.get('{'+NS['r']+'}id'),'#'+x.get(W+'anchor','')) for x in p.findall('.//w:hyperlink',NS)]
  if links:parts.append(['hyperlinks',links,None])
  return parts
 def block(e):
  if e.tag==W+'p':return ['p',paragraph(e)]
  if e.tag==W+'tbl':return ['table',[[[c.find('w:tcPr/w:gridSpan',NS).get(W+'val') if c.find('w:tcPr/w:gridSpan',NS) is not None else '1', c.find('w:tcPr/w:vMerge',NS).get(W+'val','continue') if c.find('w:tcPr/w:vMerge',NS) is not None else None, [block(p) for p in c if p.tag not in (W+'tcPr',W+'bookmarkStart',W+'bookmarkEnd',W+'proofErr')]] for c in r.findall('w:tc',NS)] for r in e.findall('w:tr',NS)]]
  if e.tag==W+'sdt':return ['sdt',[block(c) for c in e.find('w:sdtContent',NS)]]
  if e.tag==W+'sectPr':return None
  return ['other',E.QName(e).localname,''.join(e.itertext())]
 return [block(e) for e in elements if e.tag not in (W+'sectPr',W+'bookmarkStart',W+'bookmarkEnd',W+'proofErr')]

def equivalent_math(a,b):
 """Ignore only observed Word defaults; retain operators and explicit nondefaults."""
 def canonical(value,parent=None):
  if not isinstance(value,list) or not value:return value
  if value[0]=='text':
   # Word may move a commentReference into a math run without any math text.
   return None if value[1]=='' else ['text',value[1].replace('\u2212','-')]+value[2:]
  local,attrs,text,children=value;settings=dict(attrs)
  if local=='d' and sum(child[0]=='e' for child in children)==1:
   # Word omits an unused empty separator for a single delimiter argument.
   # With multiple arguments omission means '|', so retain that distinction.
   children=[
    [child[0],child[1],child[2],[prop for prop in child[3]
      if not (prop[0]=='sepChr' and dict(prop[1]).get('val')=='')]]
    if child[0]=='dPr' else child for child in children]
  if local=='chr' and parent=='naryPr' and settings.get('val')=='∫':return None
  if local=='baseJc' and parent=='mPr' and settings.get('val')=='center':return None
  if local=='mcJc' and parent=='mcPr' and settings.get('val')=='center':return None
  if local in ('subHide','supHide') and settings.get('val','1') in ('0','false'):return None
  if local=='begChr' and settings.get('val')=='(':return None
  if local=='endChr' and settings.get('val')==')':return None
  if local=='grow' and settings.get('val','1') in ('1','true'):return None
  if local=='mcs':
   # Word coalesces adjacent matrix columns with identical properties. Expand
   # counts for comparison, retaining column count and nondefault alignment.
   columns=[]
   for column in children:
    props=next((c for c in column[3] if c[0]=='mcPr'),None)
    count=next((int(dict(c[1]).get('val','1')) for c in props[3] if c[0]=='count'),1) if props else 1
    rest=[c for c in props[3] if c[0]!='count'] if props else []
    properties=canonical(['mcPr',[], '',rest])
    columns.extend([properties]*count)
   return ['matrix_columns',columns]
  children=[c for child in children if (c:=canonical(child,local)) is not None]
  if local.endswith('Pr') and not attrs and not children:return None
  return [local,attrs,text,children]
 return canonical(a)==canonical(b)

def signature(elements,relationships,math_baselines=None):
 model=content_model(elements,relationships)
 if math_baselines is not None:
  entries=[]
  def collect(value):
   if not isinstance(value,list):return
   if len(value)==3 and value[0]=='math':entries.append(value)
   else:
    for child in value:collect(child)
  collect(model)
  # Preserve old hash identities, including manifests from before this fix.
  if len(entries)==len(math_baselines):
   for item,old in zip(entries,math_baselines):
    if equivalent_math(item[1],old):item[1]=old
 return sha(json.dumps(model,ensure_ascii=False,sort_keys=True))

def doc_relationships(doc):
 return {rid:sha(rel.target_part.blob) if not rel.is_external else rel.target_ref for rid,rel in doc.part.rels.items()}


def upgrade_semantic_manifest(baseline,manifest):
 """Recompute legacy identities from a pristine generated baseline, never edits."""
 from tracked_reviews import project_revisions
 original=read_manifest(baseline)
 if original['source']!=manifest['source'] or [r['id'] for r in original['spans']]!=[r['id'] for r in manifest['spans']]:
  raise ValueError('Generated baseline does not match legacy source mapping')
 result=copy.deepcopy(manifest);project_revisions(baseline,original);rels=doc_relationships(baseline)
 controls={};outside=[]
 for e in baseline.element.body:
  tag=e.find('w:sdtPr/w:tag',NS) if e.tag==W+'sdt' else None
  if tag is not None and tag.get(W+'val','').startswith('texword:'):
   controls[tag.get(W+'val')[8:]]=list(e.find('w:sdtContent',NS))
  elif e.tag not in (W+'bookmarkStart',W+'bookmarkEnd',W+'proofErr'):outside.append(e)
 for record in result['spans']:
  elements=controls[record['id']];record['baseline']=signature(elements,rels)
  record['math_baselines']=[math_model(m) for e in elements for m in e.findall('.//m:oMath',NS)]
  offset=0
  for unit in record.get('units',[]):
   n=len(unit['node_signatures']);part=elements[offset:offset+n]
   if len(part)!=n:raise ValueError('Generated baseline node coverage differs from legacy mapping')
   unit['node_signatures']=[signature([e],rels) for e in part]
   unit['math_baselines']=[math_model(m) for e in part for m in e.findall('.//m:oMath',NS)];offset+=n
  if 'units' in record and offset!=len(elements):raise ValueError('Legacy node coverage is incomplete')
 result['outside_baseline']=signature(outside,rels)
 if result.get('granularity')=='section':result['outside_nodes']=[signature([e],rels) for e in outside]
 result['semantic_version']=2
 return result

def attach_manifest(doc,manifest):
 from docx.opc.part import Part
 from docx.opc.packuri import PackURI
 from docx.opc.constants import RELATIONSHIP_TYPE as RT
 root=E.Element('{'+RT_NS+'}roundtrip',nsmap={'tw':RT_NS})
 root.text=json.dumps(manifest,ensure_ascii=False)
 owned=[]
 for rel in list(doc.part.rels.values()):
  if not rel.is_external and rel.reltype.endswith('/customXml'):
   existing=E.fromstring(rel.target_part.blob)
   if existing.tag==root.tag:owned.append(rel)
 if owned:
  owned[0].target_part._blob=E.tostring(root,encoding='UTF-8',xml_declaration=True)
  # Word CompareDocuments copies both inputs' custom XML. Only our own map
  # relationships are retired; unrelated document/custom XML stays untouched.
  for duplicate in owned[1:]:doc.part.drop_rel(duplicate.rId)
  return
 part=Part(PackURI('/customXml/texword-roundtrip.xml'),'application/xml',E.tostring(root,encoding='UTF-8',xml_declaration=True),doc.part.package)
 doc.part.relate_to(part,RT.CUSTOM_XML)

def read_manifest(doc):
 candidates=[]
 for rel in doc.part.rels.values():
  if not rel.is_external and rel.reltype.endswith('/customXml'):
   root=E.fromstring(rel.target_part.blob)
   if root.tag=='{'+RT_NS+'}roundtrip':candidates.append(json.loads(root.text))
 if len(candidates)==1:return candidates[0]
 if candidates:
  # Conservative migration for comparison packages saved before duplicate
  # collapse existed. Word can reorder the relationships during Save().
  ids=doc.element.xpath('./w:body/w:sdt/w:sdtPr/w:tag/@w:val')
  ids=[x[8:] for x in ids if x.startswith('texword:')]
  exact=[m for m in candidates if [r['id'] for r in m['spans']]==ids]
  compatible=exact or [m for m in candidates if set(ids).issubset({r['id'] for r in m['spans']})]
  if not compatible:raise ValueError('Ambiguous comparison source maps: no compatible controls')
  cores=[{k:v for k,v in m.items() if k!='native_comparison'} for m in compatible]
  if any(core!=cores[0] for core in cores[1:]):raise ValueError('Ambiguous comparison source maps: incompatible content baselines')
  result=copy.deepcopy(cores[0]);native=[m['native_comparison'] for m in compatible if m.get('native_comparison')]
  if native:
   merged={};provenance=[]
   for item in native:
    for key,value in item.items():
     if key in ('base_sha256','proposal_sha256','legacy_provenance_candidates'):continue
     if key in merged and merged[key]!=value:raise ValueError('Ambiguous native comparison metadata: '+key)
     merged[key]=copy.deepcopy(value)
    pair={k:item[k] for k in ('base_sha256','proposal_sha256') if k in item}
    for pair in ([pair] if pair else [])+item.get('legacy_provenance_candidates',[]):
     if pair not in provenance:provenance.append(pair)
   if len(provenance)==1:merged.update(provenance[0])
   elif provenance:merged['legacy_provenance_candidates']=provenance
   result['native_comparison']=merged
  return result
 raise ValueError('This Word document has no TexWord source map; regenerate with --roundtrip first')
