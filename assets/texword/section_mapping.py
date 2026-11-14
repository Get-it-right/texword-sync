"""Section-sized Word controls, with fine source records kept only as metadata."""
from __future__ import annotations
import copy,difflib,json,os,re,tempfile
from pathlib import Path
from lxml import etree as E
from roundtrip_model import W,NS,sha,signature,doc_relationships,field_parts


def merge_source_edits(base,latex,word,*,accepted_reviews=None):
 """Merge nonoverlapping source edits; the user's Word wins true overlaps."""
 decisions=[]
 if accepted_reviews:
  from review_syntax import prepare_reviews
  def settle(value,local=False):
   _,reviews=prepare_reviews(value)
   for review in reversed(reviews):
    resolution=accepted_reviews.get(review['key'])
    if resolution is None:continue
    if local and review['raw']!=resolution['raw']:
     decisions.append({'original':resolution['raw'],'proposed':review['raw'],
                       'retained_word':[resolution['source']]})
    value=value[:review['start']]+resolution['source']+value[review['end']:]
   return value
  base,latex,word=settle(base),settle(latex,True),settle(word)
 if latex==base:return word,decisions
 if word==base or word==latex:return latex,decisions
 pattern=r'\\\[.*?\\\]|\\\(.*?\\\)|(?<!\\)\$(?:\\.|[^$])*?(?<!\\)\$|\\(?:cite|ref|eqref|label)\s*\{[^{}]*\}|\\[A-Za-z@]+\*?|\\.|[A-Za-z0-9_]+|\r\n|[ \t]+|.'
 def tokens(value):
  from review_syntax import prepare_reviews
  _,reviews=prepare_reviews(value);result=[];cursor=0
  for review in reviews:
   result.extend(re.findall(pattern,value[cursor:review['start']],re.S))
   result.append(review['raw']);cursor=review['end']
  result.extend(re.findall(pattern,value[cursor:],re.S))
  if ''.join(result)!=value:raise ValueError('Source tokenization did not preserve input')
  return result
 original=tokens(base)
 def changes(value):
  target=tokens(value)
  return [(a,b,target[c:d]) for op,a,b,c,d in difflib.SequenceMatcher(None,original,target,autojunk=False).get_opcodes() if op!='equal']
 user=changes(word);local=changes(latex);accepted=list(user);skipped=decisions
 def overlaps(left,right):
  a,b,_=left;c,d,_=right
  if a==b and c==d:return a==c
  if a==b:return c<a<d
  if c==d:return a<c<b
  return a<d and c<b
 for edit in local:
  if edit in user:continue
  collisions=[u for u in user if overlaps(edit,u)]
  if collisions:
   a,b,value=edit
   skipped.append({'original':''.join(original[a:b]),'proposed':''.join(value),
                   'retained_word':[''.join(u[2]) for u in collisions]})
  else:accepted.append(edit)
 merged=list(original)
 for a,b,value in sorted(accepted,key=lambda e:(e[0],e[1]),reverse=True):merged[a:b]=value
 return ''.join(merged),skipped


def _key(e):
 tag=e.find('w:sdtPr/w:tag',NS) if e.tag==W+'sdt' else None
 return tag.get(W+'val','')[8:] if tag is not None and tag.get(W+'val','').startswith('texword:') else None


def group_sections(doc,manifest,relationships,comments):
 """Flatten temporary fine controls into one invisible control per heading."""
 from review_comments import review_snapshot
 fine={s['id']:s for s in manifest['spans']};body=doc.element.body;nodes=list(body)
 positions=[i for i,e in enumerate(nodes) if _key(e) in fine]
 starts=[positions[0]]
 for i in positions[1:]:
  if re.match(r'\\(?:sub)*section\*?\s*\{',fine[_key(nodes[i])]['latex']):starts.append(i)
 bounds=starts+[positions[-1]+1];sections=[];counts={}
 for a,b in zip(bounds,bounds[1:]):
  units=[];children=[];records=[]
  for e in nodes[a:b]:
   sid=_key(e)
   if sid in fine:
    r=copy.deepcopy(fine[sid]);elements=list(e.find('w:sdtContent',NS));records.append(r)
   else:r={'kind':'layout','latex':'','id':'layout'};elements=[e]
   r['node_signatures']=[signature([x],relationships) for x in elements]
   r['review_baseline']=review_snapshot(elements,comments)
   units.append(r);children.extend(elements)
  first,last=records[0],records[-1];heading=first['latex']
  counts[heading]=counts.get(heading,0)+1
  sid='s'+str(int(sha(heading+'#'+str(counts[heading]))[:12],16))
  sdt=E.Element(W+'sdt');pr=E.SubElement(sdt,W+'sdtPr')
  E.SubElement(pr,W+'tag').set(W+'val','texword:'+sid)
  E.SubElement(pr,W+'id').set(W+'val',str(100000+len(sections)))
  w15='http://schemas.microsoft.com/office/word/2012/wordml'
  E.SubElement(pr,'{'+w15+'}appearance',nsmap={'w15':w15}).set('{'+w15+'}val','hidden')
  content=E.SubElement(sdt,W+'sdtContent')
  for child in children:content.append(child)
  body.insert(body.index(nodes[a]),sdt)
  for e in nodes[a:b]:
   if e.getparent() is body:body.remove(e)
  sections.append({'id':sid,'kind':'section','start':first['start'],'end':last['end'],
    'latex':manifest['source'][first['start']:last['end']], 'units':units,
    'baseline':signature(children,relationships),'review_baseline':review_snapshot(children,comments),
    'math_sources':[m for r in records for m in r.get('math_sources',[])],
    'math_baselines':[m for r in records for m in r.get('math_baselines',[])]})
 manifest=copy.deepcopy(manifest);manifest['spans']=sections
 ignored=(W+'bookmarkStart',W+'bookmarkEnd',W+'proofErr')
 outside=[e for e in body if _key(e) is None and e.tag not in ignored]
 manifest['outside_baseline']=signature(outside,relationships)
 manifest['outside_reviews']=review_snapshot(outside,comments)
 manifest['outside_nodes']=[signature([e],relationships) for e in outside]
 ordered=list(body);first=next(i for i,e in enumerate(ordered) if _key(e))
 last=max(i for i,e in enumerate(ordered) if _key(e))
 manifest['outside_prefix_count']=sum(_key(e) is None and e.tag not in ignored for e in ordered[:first])
 manifest['outside_suffix_count']=sum(_key(e) is None and e.tag not in ignored for e in ordered[last+1:])
 return manifest


def _heading_level(doc,e):
 if e.tag!=W+'p':return None
 direct=e.find('w:pPr/w:outlineLvl',NS)
 if direct is not None:return int(direct.get(W+'val'))+1
 style=e.find('w:pPr/w:pStyle',NS)
 if style is None:return None
 sid=style.get(W+'val');seen=set()
 while sid and sid not in seen:
  seen.add(sid);s=next((s for s in doc.styles.element if s.get(W+'styleId')==sid),None)
  if s is None:break
  level=s.find('w:pPr/w:outlineLvl',NS)
  if level is not None and int(level.get(W+'val'))<3:return int(level.get(W+'val'))+1
  match=re.fullmatch(r'Heading([123])',sid)
  if match:return int(match.group(1))
  parent=s.find('w:basedOn',NS);sid=parent.get(W+'val') if parent is not None else None
 return None


def _convert(doc,elements,old,manifest,reviews,images=None):
 """Convert only changed gaps; unchanged complex tables/math never pass Pandoc."""
 from roundtrip import block_to_latex
 from adoption import simplify_latex
 if len(old)==1 and old[0].get('kind')=='environment':
  match=re.match(r'\s*\\begin\{(abstract|center)\}',old[0].get('latex',''))
  if match:
   basis=dict(old[0],kind='paragraph')
   value=_convert(doc,elements,[basis],manifest,reviews,images)
   return '\\begin{'+match[1]+'}\n'+value+'\n\\end{'+match[1]+'}'
 if len(old)==1 and old[0]['kind']=='bibliography':
  return block_to_latex(doc,elements,old[0],manifest,reviews,images)
 # Never silently filter a user block: unknown containers carry content and
 # metadata that the supported LaTeX representation cannot preserve.
 for element in elements:
  if element.tag not in (W+'p',W+'tbl',W+'bookmarkStart',W+'bookmarkEnd',W+'proofErr'):
   raise ValueError('Unsupported changed block container: '+E.QName(element).localname)
 elements=[e for e in elements if e.tag in (W+'p',W+'tbl')]
 if images is not None:
  from image_roundtrip import expand,figure_at
  elements=expand(elements)
 heading_bases=[]
 for unit in old:
  if '\\begin{figure}' in unit.get('latex',''):
   # Adopted block maps can attach a newly inserted figure to the following
   # heading. Keep its original starred/numbered semantics when splitting it.
   from source_spans import split_source
   heading_bases.extend(dict(unit,**part) for part in split_source('\\begin{document}\n'+unit['latex']+'\n\\end{document}') if part['kind']=='header')
 # A changed paragraph and its adjacent table can share one conversion gap.
 # Retain the sole table's semantic labels even when the gap has other units.
 table_bases=[unit for unit in old if unit.get('kind')=='table']
 table_basis=table_bases[0] if len(table_bases)==1 and sum(e.tag==W+'tbl' for e in elements)==1 else None
 out=[];i=0
 while i<len(elements):
  figure=figure_at(elements,i) if images is not None else None
  if figure:
   part,n=figure;out.append(images.figure(part,old,reviews));i+=n;continue
  e=elements[i];level=_heading_level(doc,e);kind='paragraph';n=1
  basis=old[0] if len(old)==1 else {}
  if level and heading_bases:
   basis=heading_bases.pop(0)
  if level is None and basis.get('kind') in ('heading','header'):
   match=re.match(r'\\((?:sub)*section)',basis['latex'])
   if match:level=1+match.group(1).count('sub')
  if level:
   kind='header';e=copy.deepcopy(e)
   pr=e.find('w:pPr',NS)
   if pr is None:pr=E.SubElement(e,W+'pPr')
   style=pr.find('w:pStyle',NS)
   if style is None:style=E.SubElement(pr,W+'pStyle')
   style.set(W+'val','Heading'+str(level));part=[e]
  elif e.tag==W+'tbl':kind='table';part=[e]
  elif i+1<len(elements) and elements[i+1].tag==W+'tbl' and _table_caption(doc,e,table_basis):kind='table';n=2;part=elements[i:i+2]
  elif basis.get('kind')=='equation':kind='equation';part=[e]
  elif e.tag==W+'p' and len(e.findall('.//m:oMath',NS))==1 and any(k=='field' and re.match(r'^SEQ\s+Equation\b',v) for k,v,_ in field_parts(e)):
   # Several edited equations may share one changed section gap. Recover each
   # display's own semantic label from its native anchor, not the gap ordinal.
   labels=[manifest['references'][name]['key'] for name in e.xpath('.//w:bookmarkStart/@w:name') if name in manifest['references'] and manifest['references'][name]['kind']=='equation']
   matches=[unit for unit in old if unit.get('kind')=='equation' and any('\\label{'+label+'}' in unit.get('latex','') for label in labels)]
   if len(matches)>1:raise ValueError('Ambiguous equation labels in changed section gap')
   basis=matches[0] if matches else {'kind':'equation','latex':'\\begin{equation}'+''.join('\\label{'+label+'}' for label in labels)+'\\end{equation}'}
   kind='equation';part=[e]
  else:part=[e]
  if kind=='table' and table_basis is not None:basis=table_basis
  same_kind=basis.get('kind')==kind or (kind=='header' and basis.get('kind')=='heading')
  record=dict(basis if same_kind and '\\begin{figure}' not in basis.get('latex','') else {},id='section-edit',kind=kind)
  if level:record['heading_level']=level
  record.setdefault('latex','');record.setdefault('math_sources',[]);record.setdefault('math_baselines',[])
  # A gap can contain several edited paragraphs or a split/merged paragraph.
  # Give every converter the gap's intact formula sources; _sanitize matches
  # each actual formula by semantic content including its explicit style.
  if len(old)>1:
   record['math_sources']=[m for unit in old for m in unit.get('math_sources',[])]
   record['math_baselines']=[m for unit in old for m in unit.get('math_baselines',[])]
  latex=block_to_latex(doc,part,record,manifest,reviews,images)
  if kind=='header':latex=simplify_latex(latex)
  out.append(latex);i+=n
 return '\n\n'.join(x for x in out if x.strip())


def _table_caption(doc,element,basis=None):
 if element.tag!=W+'p':return False
 if any(kind=='field' and re.match(r'^SEQ\s+Table\b',value) for kind,value,_ in field_parts(element)):return True
 style=element.find('w:pPr/w:pStyle',NS)
 styles={s.get(W+'styleId'):s for s in doc.styles.element.findall(W+'style')}
 seen=set();sid=style.get(W+'val') if style is not None else None
 while sid and sid not in seen:
  seen.add(sid);definition=styles.get(sid)
  name=definition.find(W+'name') if definition is not None else None
  if sid in ('TableCaption','Caption') or (name is not None and name.get(W+'val','').lower() in ('caption','table caption','题注','表题')):return True
  parent=definition.find(W+'basedOn') if definition is not None else None
  sid=parent.get(W+'val') if parent is not None else None
 if (basis and basis.get('kind')=='table' and '\\caption' in basis.get('latex','')
     and len(basis.get('node_signatures',[]))>1
     and signature([element],doc_relationships(doc))==basis['node_signatures'][0]):return True
 return False


def section_to_latex(doc,elements,record,manifest,reviews,comments,images=None):
 from review_comments import review_snapshot,reconcile_review_ids
 rels=doc_relationships(doc);old_nodes=[h for r in record['units'] for h in r['node_signatures']]
 baselines=record.get('math_baselines',[]);offset=0;new_nodes=[]
 count=sum(len(e.findall('.//m:oMath',NS)) for e in elements)
 for e in elements:
  n=len(e.findall('.//m:oMath',NS))
  new_nodes.append(signature([e],rels,baselines[offset:offset+n] if count==len(baselines) else None));offset+=n
 equal={}
 for op,a,b,c,d in difflib.SequenceMatcher(None,old_nodes,new_nodes,autojunk=False).get_opcodes():
  if op=='equal':equal.update({i:c+i-a for i in range(a,b)})
 anchors=[];offset=0
 for k,r in enumerate(record['units']):
  n=len(r['node_signatures']);mapped=[equal.get(i) for i in range(offset,offset+n)]
  if n and all(x is not None for x in mapped) and mapped==list(range(mapped[0],mapped[0]+n)):
   current=elements[mapped[0]:mapped[0]+n]
   normalized,_=reconcile_review_ids(review_snapshot(current,comments),r.get('review_baseline',[]))
   header=re.match(r'\\((?:sub)*section)',r['latex'])
   expected_level=1+header.group(1).count('sub') if header else None
   current_level=_heading_level(doc,current[0]) if len(current)==1 else None
   if normalized==r.get('review_baseline',[]) and current_level==expected_level:anchors.append((k,mapped[0],mapped[0]+n))
  offset+=n
 out=[];old_cursor=new_cursor=0
 for k,a,b in anchors+[(len(record['units']),len(elements),len(elements))]:
  if a>new_cursor:
   out.append(_convert(doc,elements[new_cursor:a],record['units'][old_cursor:k],manifest,reviews,images))
  if k<len(record['units']):out.append(record['units'][k]['latex'])
  old_cursor=k+1;new_cursor=b
 return '\n\n'.join(x for x in out if x.strip())


def _restore_saved_emphasis(doc,elements,record,relationships):
 """Read legacy Word saves with inherited bold, only on an exact baseline match.

 Word removes redundant direct bold when a style supplies it. Reconstruct that
 reader representation without writing the document or hiding actual edits.
 """
 expected=[h for unit in record['units'] for h in unit['node_signatures']]
 if len(expected)!=len(elements):return
 styles={s.get(W+'styleId'):s for s in doc.styles.element.findall(W+'style')}
 def inherited(style_id,seen=None):
  seen=set() if seen is None else seen
  if not style_id or style_id in seen or style_id not in styles:return None
  seen.add(style_id);style=styles[style_id];bold=style.find('w:rPr/w:b',NS)
  if bold is not None:return bold.get(W+'val','1') not in ('0','false','off')
  parent=style.find('w:basedOn',NS)
  return inherited(parent.get(W+'val'),seen) if parent is not None else None
 offset=0;baseline=record.get('math_baselines',[])
 for i,(element,wanted) in enumerate(zip(elements,expected)):
  count=len(element.findall('.//m:oMath',NS));maths=baseline[offset:offset+count];offset+=count
  if signature([element],relationships,maths)==wanted:continue
  modes=('header','all') if element.tag==W+'tbl' else ('all',)
  for mode in modes:
   candidate=copy.deepcopy(element)
   scope=candidate.find('w:tr',NS) if mode=='header' else candidate
   if scope is None:continue
   paragraphs=[scope] if scope.tag==W+'p' else scope.findall('.//w:p',NS)
   for paragraph in paragraphs:
    style=paragraph.find('w:pPr/w:pStyle',NS)
    paragraph_bold=inherited(style.get(W+'val') if style is not None else 'Normal')
    for run in paragraph.findall('.//w:r',NS):
     if run.find('w:rPr/w:b',NS) is not None:continue
     character=run.find('w:rPr/w:rStyle',NS)
     bold=inherited(character.get(W+'val')) if character is not None else None
     if (paragraph_bold if bold is None else bold) is not True:continue
     properties=run.find(W+'rPr')
     if properties is None:properties=E.Element(W+'rPr');run.insert(0,properties)
     E.SubElement(properties,W+'b').set(W+'val','1')
   if signature([candidate],relationships,maths)==wanted:
    element.getparent().replace(element,candidate);elements[i]=candidate;break


def sync_sections(word,source,doc,manifest,destination,report,images=None,*,merge_bases=None,capture_baselines=False,allow_structure=True,allow_template_updates=False):
 from texword import ConversionError,dump,pandoc
 from adoption import source_opcodes,_position
 from roundtrip import _mapped_interval
 from review_comments import read_comments,review_snapshot,prepare_import_reviews,reconcile_review_ids
 current=source.read_bytes().decode('utf8');base=manifest['source'];word_hash=sha(word.read_bytes())
 rels=doc_relationships(doc);comments=read_comments(doc);records=manifest['spans'];byid={r['id']:r for r in records}
 nodes=list(doc.element.body);mapped={};outside=[];outside_positions=[];conflicts=[];patches=[]
 for i,e in enumerate(nodes):
  sid=_key(e)
  if sid:
   if sid not in byid or sid in mapped:conflicts.append({'id':sid,'reason':'Unknown or duplicate section mapping'})
   mapped[sid]=(i,list(e.find('w:sdtContent',NS)))
  elif e.tag not in (W+'bookmarkStart',W+'bookmarkEnd',W+'proofErr'):outside.append(e);outside_positions.append(i)
 # Word can also fold cover-table bold runs into their inherited style.
 # Reuse exact-baseline restoration before checking protected template nodes.
 _restore_saved_emphasis(doc,outside,{'units':[{'node_signatures':manifest['outside_nodes']}]},rels)
 additions={};template_changed=False
 blank_signature=signature([E.Element(W+'p')],{})
 for op,a,b,c,d in difflib.SequenceMatcher(None,manifest['outside_nodes'],[signature([e],rels) for e in outside],autojunk=False).get_opcodes():
  if op=='equal':continue
  protected=(b<=manifest['outside_prefix_count'] or a>=len(manifest['outside_nodes'])-manifest['outside_suffix_count'])
  old_count=sum(value!=blank_signature for value in manifest['outside_nodes'][a:b])
  from word_runtime import layout_blank
  new_count=sum(not layout_blank(e) for e in outside[c:d])
  if allow_template_updates and protected and old_count==new_count:
   template_changed=True;continue
  if op!='insert' or a<manifest['outside_prefix_count'] or a>len(manifest['outside_nodes'])-manifest['outside_suffix_count']:
   conflicts.append({'reason':'Template content outside sections changed'});continue
  for j in range(c,d):
   e=outside[j]
   # Word inserts empty carrier paragraphs between adjacent block controls.
   # They are layout, not newly written source text; retain them in Word.
   if e.tag==W+'p' and all(kind=='text' and not value.strip() for kind,value,_ in field_parts(e)) and e.find('w:pPr/w:sectPr',NS) is None:continue
   additions[outside_positions[j]]=e
 # Existing template comments remain outside the editable section bodies.
 retained=[e for i,e in zip(outside_positions,outside) if i not in additions]
 normalized,_=reconcile_review_ids(review_snapshot(retained,comments),manifest.get('outside_reviews',[]))
 if normalized!=manifest.get('outside_reviews',[]):conflicts.append({'reason':'Template comments changed'})
 ops=source_opcodes(base,current);segments={};changes=0;skipped_latex=[];word_bases={}
 for k,r in enumerate(records):
  stop=records[k+1]['start'] if k+1<len(records) else r['end']
  a=_position(ops,r['start']);b=_position(ops,stop)
  if a is None or b is None:conflicts.append({'id':r['id'],'reason':'Source changes crossed section boundaries'});continue
  if r['id'] not in mapped:
   word_bases[r['id']]=''
   if current[a:b]!=base[r['start']:stop]:
    skipped_latex.append({'id':r['id'],'original':base[r['start']:stop],
                         'proposed':current[a:b],'retained_word':['']})
   patches.append({'id':r['id'],'operation':'delete'});changes+=1;continue
  elements=mapped[r['id']][1]
  try:
   _restore_saved_emphasis(doc,elements,r,rels)
   snapshot=review_snapshot(elements,comments);normalized,origins=reconcile_review_ids(snapshot,r.get('review_baseline',[]))
   changed=signature(elements,rels,r.get('math_baselines'))!=r['baseline'] or normalized!=r.get('review_baseline',[])
   expected_outline=[];offset=0
   for unit in r['units']:
    heading=re.match(r'\\((?:sub)*section)',unit['latex'])
    if heading:expected_outline.append((offset,1+heading.group(1).count('sub')))
    offset+=len(unit['node_signatures'])
   actual_outline=[(i,level) for i,e in enumerate(elements) if (level:=_heading_level(doc,e)) is not None]
   changed=changed or actual_outline!=expected_outline
   original_segment=base[r['start']:stop]
   word_segment=original_segment
   if changed:
    reviews=prepare_import_reviews(snapshot,manifest,origins)
    replacement=section_to_latex(doc,elements,r,manifest,reviews,comments,images)
    tail=base[r['end']:stop]
    word_segment=replacement+tail
    from roundtrip import same_inline_math_notation
    if same_inline_math_notation(word_segment,original_segment):word_segment=original_segment
   word_bases[r['id']]=word_segment
   if changed or merge_bases is not None or manifest.get('accepted_review_sources'):
    common=(merge_bases or {}).get(r['id'],original_segment)
    segments[r['id']],skipped=merge_source_edits(common,current[a:b],word_segment,
        accepted_reviews=manifest.get('accepted_review_sources'))
    skipped_latex.extend(dict(item,id=r['id']) for item in skipped)
    # Word may normalize math XML without changing the resulting LaTeX.
    if segments[r['id']]!=current[a:b]:
     patches.append({'id':r['id'],'operation':'edit','before':current[a:b],'after':segments[r['id']]});changes+=1
   else:segments[r['id']]=current[a:b]
  except (ConversionError,ValueError) as exc:conflicts.append({'id':r['id'],'reason':str(exc)})
 pieces=[];pending=[]
 def flush():
  if pending:
   pieces.append(_convert(doc,pending,[],manifest,prepare_import_reviews(review_snapshot(pending,comments),manifest,{}),images)+'\n\n');pending.clear()
 for i,e in enumerate(nodes):
  if i in additions:pending.append(e)
  elif _key(e) in byid:
   try:flush()
   except (ConversionError,ValueError) as exc:conflicts.append({'reason':str(exc)})
   pieces.append(segments.get(_key(e),''))
 try:flush()
 except (ConversionError,ValueError) as exc:conflicts.append({'reason':str(exc)})
 start=_position(ops,records[0]['start']);end=_position(ops,records[-1]['end'])
 if start is None or end is None:conflicts.append({'reason':'Source document bounds changed'})
 order=[_key(e) for e in nodes if _key(e) in byid]
 reordered=order!=[r['id'] for r in records if r['id'] in mapped]
 result={'status':'conflict' if conflicts else 'preview','word':str(word),'source':str(source),
  'output':str(destination) if destination else None,'changed_blocks':changes+bool(additions)+reordered+template_changed,
  'template_changed':template_changed,
  'conflicts':conflicts,'patches':patches,'applied':False,'granularity':'section',
  'added_nodes':len(additions),'deleted_sections':len(byid.keys()-mapped.keys()),'reordered_sections':reordered,
  'skipped_latex_edits':skipped_latex,'merge_policy':'word_wins_overlaps'}
 if capture_baselines:result['word_bases']=word_bases
 if not allow_structure and (result['added_nodes'] or result['deleted_sections'] or result['reordered_sections']):
  result['status']='conflict';conflicts.append({'reason':'Structural changes require a full pull to refresh Word mappings; source-only import was not applied'})
 if conflicts:dump(report,result);raise ConversionError('Section synchronization stopped without changing LaTeX; see '+str(report))
 updated=current[:start]+''.join(pieces)+current[end:]
 # Preserve source bytes exactly for a no-change preview.
 if not result['changed_blocks']:updated=current
 else:
  from roundtrip import ensure_import_preamble
  updated,result['preamble_additions']=ensure_import_preamble(updated)
 if images is not None:
  try:images.validate(updated)
  except ConversionError as exc:
   result['status']='conflict';result['conflicts'].append({'reason':str(exc)});dump(report,result);raise
  result['image_assets']=images.inventory()
 if destination:
  from review_syntax import prepare_reviews
  validated,_=prepare_reviews(updated);pandoc(['-f','latex','-t','json'],cwd=source.parent,input=validated)
  if source.read_bytes().decode('utf8')!=current or sha(word.read_bytes())!=word_hash:raise ConversionError('Inputs changed during synchronization; retry')
  destination.parent.mkdir(parents=True,exist_ok=True)
  with tempfile.NamedTemporaryFile(dir=destination.parent,suffix='.tex',delete=False) as f:temporary=Path(f.name);f.write(updated.encode('utf8'))
  try:
   if images is not None:images.publish(destination,temporary)
   else:os.replace(temporary,destination)
  finally:
   if temporary.exists():temporary.unlink()
  result['status']='applied';result['applied']=True
 dump(report,result)
 return {k:v for k,v in result.items() if k!='patches'}
