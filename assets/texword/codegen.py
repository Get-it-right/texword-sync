"""Lower parsed Word nodes to inspectable Python construction statements."""
from __future__ import annotations
import copy,json,os,posixpath,pprint,re,shutil,subprocess,sys,tempfile,time,zipfile
from pathlib import Path
from lxml import etree as E
from texword import NS,W,ConversionError,load,dump,digest,pandoc,transform_ast,write_adapter,ROOT,bookmark_name
R='{'+NS['r']+'}'

class WordIR:
 def __init__(self,docx,assets):
  self.z=zipfile.ZipFile(docx);self.assets=assets
  rel=E.fromstring(self.z.read('word/_rels/document.xml.rels'))
  self.rels={x.get('Id'):{'target':x.get('Target'),'external':x.get('TargetMode')=='External'} for x in rel}
 def tokens(self,p):
  out=[];field=None;separate=False
  for x in p:
   local=E.QName(x).localname
   if local=='pPr':continue
   if local=='r':
    rp=x.find('w:rPr',NS);fmt=E.tostring(rp,encoding='unicode') if rp is not None else None
    for c in x:
     name=E.QName(c).localname
     if name=='rPr':continue
     if name=='fldChar':
      typ=c.get(W+'fldCharType')
      if typ=='begin':field={'code':'','value':''};separate=False
      elif typ=='separate':separate=True
      elif typ=='end':
       if field is None:raise ConversionError('Malformed generated field')
       field['code']=field['code'].strip();out.append({'field':field});field=None
     elif name=='instrText':
      if field is None:raise ConversionError('Field instruction without begin')
      field['code']+=c.text or ''
     elif name=='t':
      if field is not None:
       if separate:field['value']+=c.text or ''
      else:
       token={'text':c.text or ''}
       if fmt:token['rPr']=fmt
       out.append(token)
     elif name=='tab':out.append({'tab':True})
     elif name=='br':out.append({'break':c.get(W+'type','')})
     elif name=='commentReference':continue
     elif name in ['drawing','pict']:
      embeds=c.xpath('.//@r:embed',namespaces=NS)
      extents=c.xpath('.//*[local-name()="extent"]')
      if len(embeds)!=1 or not extents:raise ConversionError('Unsupported drawing: expected one embedded image')
      target=posixpath.normpath(posixpath.join('word',self.rels[embeds[0]]['target']))
      raw=self.z.read(target);name='image_'+__import__('hashlib').sha256(raw).hexdigest()[:16]+Path(target).suffix
      (self.assets/name).write_bytes(raw)
      image={'path':name,'width':int(extents[0].get('cx')),'height':int(extents[0].get('cy'))}
      crop=c.xpath('.//*[local-name()="docPr"]/@title')
      if crop and re.fullmatch(r'TWIMAGECROP:-?\d+,-?\d+,-?\d+,-?\d+',crop[0]):
       image['crop']=list(map(int,crop[0].split(':',1)[1].split(',')))
      out.append({'image':image})
     elif name not in ['lastRenderedPageBreak','proofErr']:raise ConversionError('Unsupported Word inline: '+name)
   elif local in ['oMath','oMathPara']:out.append({'math':E.tostring(x,encoding='unicode')})
   elif local=='bookmarkStart':out.append({'bookmark_start':{'name':x.get(W+'name'),'id':x.get(W+'id')}})
   elif local=='bookmarkEnd':out.append({'bookmark_end':x.get(W+'id')})
   elif local=='commentRangeStart':out.append({'review_start':x.get(W+'id')})
   elif local=='commentRangeEnd':out.append({'review_end':x.get(W+'id')})
   elif local=='hyperlink':
    spec={'tokens':self.tokens(x)}
    if x.get(W+'anchor'):spec['anchor']=x.get(W+'anchor')
    else:spec['url']=self.rels[x.get(R+'id')]['target']
    out.append({'hyperlink':spec})
   elif local in ['proofErr','permStart','permEnd']:continue
   else:raise ConversionError('Unsupported paragraph node: '+local)
  if field is not None:raise ConversionError('Unclosed generated field')
  merged=[]
  for token in out:
   if merged and 'text' in token and 'text' in merged[-1] and token.get('rPr')==merged[-1].get('rPr'):
    merged[-1]['text']+=token['text']
   else:merged.append(token)
  return merged
 def paragraph(self,p):
  style=p.find('w:pPr/w:pStyle',NS);style=style.get(W+'val') if style is not None else 'BodyText'
  pp=copy.deepcopy(p.find('w:pPr',NS));ppr=None
  if pp is not None:
   if pp.find('w:numPr',NS) is not None:raise ConversionError('Unsupported: list numbering needs an explicit runtime numbering adapter')
   ps=pp.find('w:pStyle',NS)
   if ps is not None:pp.remove(ps)
   if len(pp):ppr=E.tostring(pp,encoding='unicode')
  return {'style':style,'tokens':self.tokens(p),'pPr':ppr}
 def table(self,t):
  rows=[]
  for tr in t.findall('w:tr',NS):
   row=[]
   for c in tr.findall('w:tc',NS):
    cell={'paragraphs':[self.paragraph(p) for p in c.findall('w:p',NS)]}
    if c.find('w:tbl',NS) is not None:raise ConversionError('Unsupported: nested tables')
    span=c.find('w:tcPr/w:gridSpan',NS)
    if span is not None:cell['span']=int(span.get(W+'val'))
    vm=c.find('w:tcPr/w:vMerge',NS)
    if vm is not None:cell['vmerge']=vm.get(W+'val','')
    row.append(cell)
   rows.append(row)
  grid=t.find('w:tblGrid',NS)
  return rows,[int(x.get(W+'w','1')) for x in grid] if grid is not None else []
 def close(self):self.z.close()

def compile_code(adapter,source,code,*,bibliography=None,overwrite=False,roundtrip=False):
 """Compile TeX to substantive Python source; do not generate the final Word document."""
 started=time.perf_counter();adapter=Path(adapter).resolve();source=Path(source).resolve();code=Path(code).resolve()
 profile=load(adapter/'profile.json');template=adapter/'template.docx'
 if code.suffix.lower()!='.py':raise ConversionError('Generated code must have a .py suffix')
 if code in [source,template,Path(profile['template_origin']).resolve()]:raise ConversionError('Refusing to overwrite original source/template')
 if code.exists() and not overwrite:raise ConversionError('Generated code exists; use --overwrite for derived code')
 if digest(template)!=profile['template_sha256']:raise ConversionError('Template changed; regenerate the template adapter')
 if re.search(r'\\begin\{(?:tikzpicture|tikzcd|pspicture)\}',source.read_text(encoding='utf8')):raise ConversionError('Unsupported: diagram environments require explicitly exported figures')
 from review_syntax import prepare_reviews,format_review
 from review_comments import inject_review_markers
 imported=(adapter/'imported.json').exists();reuse={};layout={}
 if imported:
  if not roundtrip:raise ConversionError('Adopted projects require --roundtrip')
  from adoption import prepare_adopted
  parsed,source_map,reviews,reuse,layout=prepare_adopted(adapter,source)
 else:
  prepared,reviews=prepare_reviews(source.read_bytes().decode('utf8'))
  for index,d in enumerate(reviews):d['compiled_id']=str(index);d['rendered_text']=format_review(d)
  ast_raw,_=pandoc(['-f','latex','-t','json'],cwd=source.parent,input=prepared)
  parsed=json.loads(ast_raw);source_map=None
  if roundtrip:
   from roundtrip import map_source
   parsed,source_map=map_source(source,parsed)
 parsed=inject_review_markers(parsed,reviews)
 proposal_citations=[key for d in reviews if d['kind']=='revision' for key in d.get('proposal_citations',[])]
 ast,tr=transform_ast(
  parsed,profile,cwd=source.parent,
  bibliography=Path(bibliography).resolve() if bibliography else None,
  proposal_citations=() if imported else proposal_citations,
 )
 if imported:
  bibliography=load(adapter/'imported.json').get('bibliography',{})
  if isinstance(bibliography,dict):
   bibliography_numbers={key:entry['number'] for key,entry in bibliography.items()}
  elif isinstance(bibliography,list):
   bibliography_numbers={entry['key']:entry['number'] for entry in bibliography}
  else:
   raise ConversionError('Adopted bibliography registry must be a key-to-entry mapping or an entry list')
  citation_numbers={key:entry.get('number',bibliography_numbers.get(key))
                    for key,entry in source_map.get('native_references',{}).items()
                    if entry.get('kind')=='citation' and (entry.get('number') is not None or key in bibliography_numbers)}
  citation_targets={key for key,entry in source_map.get('native_references',{}).items()
                    if entry.get('kind')=='citation'}
 else:
  bibliography_numbers={key:number for key,number in tr.citations.items()}
  citation_numbers=bibliography_numbers
  citation_targets={bookmark_name('bib:'+key) for key in tr.citations}
 for review in reviews:
  if review['kind']!='revision':
   continue
  tokens=[]
  for part in review.get('proposal_parts',[]):
   if 'text' in part:
    if part['text']:
     tokens.append({'text':part['text']})
    continue
   for key in part.get('citations',[]):
    if imported:
     target=key
    else:
     target=bookmark_name('bib:'+key)
    if key not in citation_numbers or target not in citation_targets:
     raise ConversionError('Tracked revision citation target is unavailable: '+key)
    tokens.extend([
     {'text':'['},
     {'field':{'code':'REF '+target+' \\h','value':str(citation_numbers[key])}},
     {'text':']'},
    ])
  review['proposal_tokens']=tokens
  review['proposal']=''.join(
   part['text'] if 'text' in part else ''.join(f"[{citation_numbers[key]}]" for key in part.get('citations',[]))
   for part in review.get('proposal_parts',[])
  )
  review['proposal_empty']=not review['proposal'].strip()
  review['citation_targets']=sorted(citation_targets)
  review['rendered_text']=format_review(review)
 code.parent.mkdir(parents=True,exist_ok=True);assets=code.parent/(code.stem+'_assets');assets.mkdir(exist_ok=True)
 shutil.copyfile(template,assets/'template.docx')
 # Preserve user calibration in the generated template-specific Python code.
 shutil.copyfile(adapter/'template_adapter.py',assets/'template_adapter.py')
 shutil.copyfile(ROOT/'word_runtime.py',assets/'word_runtime.py')
 shutil.copyfile(ROOT/'review_comments.py',assets/'review_comments.py')
 shutil.copyfile(ROOT/'tracked_reviews.py',assets/'tracked_reviews.py')
 shutil.copyfile(ROOT/'native_comparison.py',assets/'native_comparison.py')
 if roundtrip:
  source_map['granularity']=profile.get('source_mapping_granularity','block')
  source_map['allow_merged_table_edits']=profile.get('allow_merged_table_edits',False)
  source_map['references']={bookmark_name(k):{'key':k,'kind':kind} for k,(value,kind) in tr.targets.items()}
  source_map['references'].update({bookmark_name('bib:'+k):{'key':k,'kind':'citation'} for k in tr.citations})
  source_map['references'].update(source_map.get('native_references',{}))
  dump(assets/'source_map.json',source_map)
  shutil.copyfile(ROOT/'roundtrip_model.py',assets/'roundtrip_model.py')
  shutil.copyfile(ROOT/'section_mapping.py',assets/'section_mapping.py')
 statements=[];bibliography_heading=False;counts={'paragraphs':0,'tables':0,'equations':0,'headings':0}
 with tempfile.TemporaryDirectory(prefix='texword_compile_') as temp:
  body=Path(temp)/'intermediate.docx'
  # Pandoc is a compile-time frontend: it parses TeX and supplies native OMML. The emitted program never calls it.
  pandoc(['-f','json','-t','docx','--reference-doc',adapter/'reference.docx','--resource-path',source.parent,'-o',body],cwd=source.parent,input=json.dumps(ast,ensure_ascii=False))
  ir=WordIR(body,assets)
  try:
   document=E.fromstring(ir.z.read('word/document.xml')).find('w:body',NS)
   skip=set();reused_until=None
   for index,item in enumerate(document):
    if index in skip:continue
    if item.tag==W+'sectPr':continue
    marker=''.join(item.xpath('.//w:t/text()',namespaces=NS)) if item.tag==W+'p' else ''
    if reused_until:
     if marker=='TWRT_END_'+reused_until:statements.append('    document.end_source()\n');reused_until=None
     continue
    if marker in layout:
     statements.append('    document.original_blocks('+repr(layout[marker])+')\n');continue
    m=re.fullmatch(r'TWRT_BEGIN_(s\d+)',marker)
    if m and m.group(1) in reuse:
     sid=m.group(1);anchors=[{'name':bookmark_name(k),'value':value,'kind':kind} for k,(value,kind) in tr.targets.items() if tr.target_spans.get(k)==sid]
     statements.append('    document.begin_source('+repr(sid)+')\n    document.original_blocks('+repr(reuse[sid])+', anchors='+repr(anchors)+')\n')
     reused_until=sid;continue
    if item.tag in [W+'bookmarkStart',W+'bookmarkEnd']:
     # Pandoc's Div anchors are redundant: explicit TW anchors live in numbered paragraphs/captions.
     if item.get(W+'name','').startswith('TW_'):raise ConversionError('Unexpected top-level semantic bookmark')
     continue
    if item.tag==W+'tbl':
     rows,widths=ir.table(item);counts['tables']+=1
     statements.append('    document.table(\n        rows='+pprint.pformat(rows,width=105,sort_dicts=False)+',\n        widths='+repr(widths)+')\n')
     continue
    if item.tag!=W+'p':raise ConversionError('Unsupported body node: '+E.QName(item).localname)
    p=ir.paragraph(item);tokens=p['tokens'];style=p['style'];counts['paragraphs']+=1
    marker=''.join(v.get('text','') for v in tokens)
    if roundtrip and re.fullmatch(r'TWRT_(?:BEGIN_s\d+|END_s\d+)',marker):
     operation,span_id=marker[5:].split('_',1)
     statements.append('    document.'+('begin_source('+repr(span_id)+')' if operation=='BEGIN' else 'end_source()')+'\n')
     continue
    if any('math' in v for v in tokens) and index+1<len(document) and document[index+1].tag==W+'p':
     following=ir.paragraph(document[index+1])
     if any('SEQ Equation' in v.get('field',{}).get('code','') for v in following['tokens']):
      tokens+=following['tokens'];skip.add(index+1)
    if not tokens:statements.append('    document.paragraph([])\n');continue
    token_code=pprint.pformat(tokens,width=110,sort_dicts=False)
    if re.match(r'^Heading[123]$',style):
     counts['headings']+=1;statements.append('    document.heading('+style[-1]+',\n        '+token_code+')\n')
    elif any('math' in v for v in tokens) and any('Equation' in v.get('field',{}).get('code','') for v in tokens):
     counts['equations']+=1;statements.append('    document.equation(\n        '+token_code+')\n')
    elif style=='Bibliography':
     if not bibliography_heading and not imported:statements.append("    document.heading(1, [{'text': "+repr(profile.get('bibliography_heading','参考文献'))+"}])\n");bibliography_heading=True
     statements.append('    document.bibliography_entry(\n        '+token_code+')\n')
    elif 'Caption' in style:statements.append('    document.caption(\n        '+token_code+', style='+repr(style)+')\n')
    else:statements.append('    document.paragraph(\n        '+token_code+',\n        style='+repr(style)+', pPr='+repr(p['pPr'])+')\n')
  finally:ir.close()
 setup='    document.enable_roundtrip()\n' if roundtrip else ''
 if reviews:setup+='    document.set_reviews('+pprint.pformat(reviews,width=105,sort_dicts=False)+')\n'
 source_code='''# Generated Word-building program. This file is editable Python source.\n# Execution reads neither LaTeX nor BibTeX and never invokes Pandoc.\nfrom pathlib import Path\nimport sys\nsys.path.insert(0, str(Path(__file__).resolve().parent / '''+repr(assets.name)+'''))\nfrom template_adapter import Report\nfrom word_runtime import output_arguments\n\ndef build(output, *, overwrite=False, template_regions_from=None):\n    document = Report()\n    if template_regions_from is not None: document.preserve_template_regions(template_regions_from)\n'''+setup+ '\n'.join(statements)+'''\n    return document.finish(output, overwrite=overwrite)\n\nif __name__ == "__main__":\n    import json\n    args = output_arguments(str(Path(__file__).with_suffix(".docx")))\n    print(json.dumps(build(args.output, overwrite=args.overwrite, template_regions_from=args.template_regions_from), ensure_ascii=False, indent=2))\n'''
 compile(source_code,str(code),'exec')
 code.write_text(source_code,encoding='utf8')
 report={'status':'compiled','source':str(source),'generated_code':str(code),'assets':str(assets),'counts':counts,
         'citations':len(tr.citations),'labels':len(tr.targets),'elapsed_seconds':time.perf_counter()-started,'template_warnings':profile['warnings'],
         'review_comments':len(reviews),
         'final_docx_generated_at_compile_time':False}
 dump(code.with_suffix('.compile.json'),report)
 return report
