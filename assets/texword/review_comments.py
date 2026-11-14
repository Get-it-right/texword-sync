"""Native Word review comments; suggestions never substitute their anchor text."""
from __future__ import annotations
import json,re
from lxml import etree as E

NS={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14':'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15':'http://schemas.microsoft.com/office/word/2012/wordml'}
W='{'+NS['w']+'}'

def inject_review_markers(ast,definitions):
 markers={}
 for d in definitions:
  cid=str(d['compiled_id'])
  markers[d['start_marker']]='<w:commentRangeStart xmlns:w="'+NS['w']+'" w:id="'+cid+'"/>'
  markers[d['end_marker']]='<w:commentRangeEnd xmlns:w="'+NS['w']+'" w:id="'+cid+'"/><w:r xmlns:w="'+NS['w']+'"><w:commentReference w:id="'+cid+'"/></w:r>'
 if not markers:return ast
 pattern=re.compile('('+'|'.join(map(re.escape,markers))+')')
 found={k:0 for k in markers}
 def walk(o):
  if isinstance(o,list):
   result=[]
   for x in o:
    v=walk(x)
    if isinstance(x,dict) and x.get('t')=='Str' and isinstance(v,list):result+=v
    else:result.append(v)
   return result
  if not isinstance(o,dict):return o
  if o.get('t')=='Str':
   pieces=pattern.split(o['c']);out=[]
   for part in pieces:
    if part in markers:found[part]+=1;out.append({'t':'RawInline','c':['openxml',markers[part]]})
    elif part:out.append({'t':'Str','c':part})
   return out
  return {k:walk(v) for k,v in o.items()}
 result=walk(ast)
 if any(count!=1 for count in found.values()):raise ValueError('A review anchor was not parsed exactly once; use a nonempty ordinary inline text anchor')
 return result

def read_comments(doc):
 comments={};resolved={}
 for rel in doc.part.rels.values():
  if rel.is_external:continue
  if rel.reltype.endswith('/comments'):
   root=E.fromstring(rel.target_part.blob)
   for c in root.findall('w:comment',NS):
    # Retain paragraph breaks, but never include Word's annotation reference glyph.
    def paragraph_text(p):
     nodes=E._Element.xpath(p,'.//w:t | .//w:tab | .//w:br | .//w:cr',namespaces=NS)
     return ''.join((x.text or '') if x.tag==W+'t' else '\t' if x.tag==W+'tab' else '\n' for x in nodes)
    text='\n'.join(paragraph_text(p) for p in c.findall('w:p',NS))
    unsupported=bool(E._Element.xpath(c,'.//w:drawing | .//w:pict | .//w:object | .//w:sym | .//w:tbl | .//w:hyperlink | .//w:fldChar | .//w:footnoteReference | .//w:ins | .//w:del | .//w:b | .//w:i | .//w:u | .//m:oMath',namespaces=dict(NS,m='http://schemas.openxmlformats.org/officeDocument/2006/math')))
    ids=E._Element.xpath(c,'.//w:p/@w14:paraId',namespaces=NS)
    comments[c.get(W+'id')]={'text':text,'author':c.get(W+'author',''),'unsupported':unsupported,'paragraph_ids':ids}
  elif rel.reltype.endswith('/commentsExtended'):
   root=E.fromstring(rel.target_part.blob)
   for c in root:
    if c.get('{'+NS['w15']+'}paraIdParent'):raise ValueError('Threaded comment replies require manual reconciliation; no replies were discarded')
    resolved[c.get('{'+NS['w15']+'}paraId')]=c.get('{'+NS['w15']+'}done','0')=='1'
 for c in comments.values():c['resolved']=any(resolved.get(p,False) for p in c.pop('paragraph_ids'))
 return comments

def review_snapshot(elements,comments):
 result=[]
 for e in elements:
  paragraphs=([e] if e.tag==W+'p' else [])+list(e.findall('.//w:p',NS))
  for p in paragraphs:
   active=None;anchor=[];managed_origin=None
   children=list(p);start_index=0
   for index,x in enumerate(children):
    if x.tag==W+'commentRangeStart':
     if active is not None:raise ValueError('Overlapping comments require manual reconciliation')
     active=x.get(W+'id');anchor=[]
     start_index=index
     managed_origin=x.get('{urn:texword:roundtrip:1}managedOrigin')
    elif x.tag==W+'commentRangeEnd':
     if active!=x.get(W+'id'):raise ValueError('Comment range crosses paragraphs or is malformed')
     if active not in comments:raise ValueError('Word comment content is missing: '+str(active))
     item=dict(comments[active],id=active,anchor=''.join(anchor))
     def prose(nodes):return ''.join(t.text or '' for node in nodes for t in node.iter(W+'t'))
     item['context']=[prose(children[:start_index])[-24:],prose(children[index+1:])[:24]]
     if managed_origin is not None:item['managed_origin']=managed_origin
     result.append(item);active=None
    elif active is not None:anchor+=E._Element.xpath(x,'.//w:t/text() | .//m:t/text()',namespaces=dict(NS,m='http://schemas.openxmlformats.org/officeDocument/2006/math'))
   if active is not None:raise ValueError('Comment range crosses paragraphs; reconcile it manually')
 return result

def reconcile_review_ids(snapshot,baseline):
 """Word may renumber comment IDs on save. Match only unambiguous local evidence."""
 fields=('anchor','text','author')
 # Snapshot order follows document positions, not numeric IDs. An unchanged
 # ordered sequence remains unambiguous even when Word renumbers duplicates.
 if len(snapshot)==len(baseline):
  identical=all(all(a.get(k)==b.get(k) for k in fields) for a,b in zip(snapshot,baseline))
  stable_ids=all(a['id']==b['id'] and a.get('author')==b.get('author')
                 and (a.get('anchor')==b.get('anchor') or a.get('text')==b.get('text'))
                 for a,b in zip(snapshot,baseline))
  if identical or stable_ids:
   normalized=[];origins={}
   for item,old in zip(snapshot,baseline):
    value=dict(item);value.pop('managed_origin',None);value['id']=old['id']
    if 'context' not in old:value.pop('context',None)
    normalized.append(value);origins[item['id']]=old['id']
   return normalized,origins
 remaining=list(baseline);normalized=[];origins={}
 for item in snapshot:
  if 'managed_origin' in item:
   candidates=[b for b in remaining if b['id']==item['managed_origin']]
  else:candidates=[b for b in remaining if all(b.get(k)==item.get(k) for k in ('anchor','text','author'))]
  if 'managed_origin' not in item:
   if candidates and item.get('context') is not None:
    unchanged_ids=[b for b in candidates if b['id']==item['id']]
    contextual=[b for b in candidates if b.get('context')==item['context']]
    if len(unchanged_ids)==1:candidates=unchanged_ids
    elif any('context' in b for b in candidates):candidates=contextual
   if not candidates:
    candidates=[b for b in remaining if b.get('anchor')==item.get('anchor') and b.get('author')==item.get('author')
                and (b.get('context')==item.get('context') if 'context' in item and 'context' in b else b['id']==item['id'])]
   if not candidates and 'context' in item:
    candidates=[b for b in remaining if b.get('text')==item.get('text') and b.get('author')==item.get('author') and b.get('context')==item['context']]
  if len(candidates)>1:raise ValueError('Ambiguous comment identity after Word save; reconcile duplicate anchors/notes manually')
  value=dict(item);value.pop('managed_origin',None)
  if candidates:
   original=candidates[0];remaining.remove(original);origins[item['id']]=original['id'];value['id']=original['id']
   if 'context' not in original:value.pop('context',None)
  else:value['id']='new:'+item['id']
  normalized.append(value)
 return normalized,origins

def prepare_import_reviews(snapshot,manifest,origins=None):
 """A resolved or removed comment is never approval to apply the suggested replacement."""
 definitions={d['native_id']:d for d in manifest.get('reviews',[])};result={}
 for item in snapshot:
  if item['unsupported']:raise ValueError('Rich comment content requires manual reconciliation')
  if item['resolved']:raise ValueError('A comment was marked resolved; explicitly confirm its disposition before rebuilding (the proposal was not applied)')
  origin=origins.get(item['id']) if origins is not None else item['id']
  old=definitions.get(origin);key=old['key'] if old else 'word-comment-'+item['id']
  if old and item['text']==old['rendered_text']:spec=dict(old)
  else:spec={'kind':'comment','key':key,'note':item['text'],'author':item['author'] or 'Codex'}
  result[item['id']]=spec
 return result
