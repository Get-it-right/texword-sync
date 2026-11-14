"""One document check entrypoint; no unit-suite reruns or ad-hoc reconversion."""
from __future__ import annotations
import argparse,json,re,subprocess,time,zipfile
from pathlib import Path
from lxml import etree as E
from docx import Document
from texword import ROOT,digest,load,dump,ConversionError
from roundtrip_model import field_parts,read_manifest,sha
from roundtrip import sync_word

def check_project(project,*,baseline=False,word=False,candidate=None):
 started=time.perf_counter();project=Path(project).resolve();config=load(project/'workflow.json');path=Path(config['word']);source=Path(config['source']);checks={};problems=[]
 if candidate is not None:
  if word or baseline:raise ConversionError('Candidate checking is read-only and cannot operate Word or adoption baseline')
  path=Path(candidate).resolve()
 sync_report=path.with_suffix('.sync.json') if candidate is not None else project/'sync_check.json'
 result_report=path.with_suffix('.check.json') if candidate is not None else project/'check_result.json'
 if word:
  command=['C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe','-NoProfile','-ExecutionPolicy','Bypass','-File',str(ROOT/'check_word.ps1'),'-ProjectPath',str(project)]
  r=subprocess.run(command,capture_output=True,text=True,encoding='utf8',errors='replace')
  if r.returncode:raise ConversionError(r.stderr.strip() or r.stdout.strip())
  checks['microsoft_word']=load(project/'word_check.json')
  if checks['microsoft_word'].get('field_update_result',0):problems.append('Microsoft Word reported a field update error')
 verified_word_hash=digest(path);verified_source_hash=digest(source)
 document=Document(path)
 manifest=read_manifest(document)
 if manifest.get('native_comparison',{}).get('schema')==1:
  from native_comparison import project_reader
  checks['native_review']={'revision_nodes':len(document.element.xpath('.//w:ins|.//w:del')),'reader_view':'proposed_with_settled_author_decisions'}
  project_reader(document,manifest)
 with zipfile.ZipFile(path) as z:
  if z.testzip():raise ConversionError('DOCX ZIP integrity failed')
  checks['package_integrity']=True
 names=document.element.xpath('.//w:bookmarkStart/@w:name');fields=[v for p in document.element.xpath('.//w:p') for k,v,f in field_parts(p) if k=='field']
 targets=[m.group(1) for v in fields if (m:=re.match(r'REF\s+(\S+)',v))]
 missing=sorted(set(targets)-set(names));duplicates=sorted(n for n in set(names) if names.count(n)>1)
 checks['references']={'REF':len(targets),'SEQ':sum(v.startswith('SEQ ') for v in fields),'missing':missing,'duplicate_bookmarks':duplicates}
 if missing or duplicates:problems.append('Unresolved or duplicate reference targets')
 text=''.join(document.element.xpath('.//w:t/text() | .//m:t/text()'))
 errors=re.findall(r'Error!\s*Reference source not found|错误[!！]\s*未找到引用源',text,re.I)
 if errors:problems.append('Word reference error text present')
 checks['structure']={'math':len(document.element.xpath('.//m:oMath')),'tables':len(document.element.xpath('.//w:tbl')),'sections':len(document.element.xpath('.//w:sectPr')),'source_controls':sum(x.startswith('texword:') for x in document.element.xpath('.//w:sdtPr/w:tag/@w:val'))}
 manifest=read_manifest(document);checks['source_map_integrity']=sha(manifest['source'])==manifest['source_sha256']
 if not checks['source_map_integrity']:problems.append('Corrupt source-map baseline')
 try:
  if candidate is None and manifest.get('semantic_version',1)<2:
   from workflow_runtime import import_word_edits
   status=import_word_edits(project,config,path,source,apply=False)
  else:status=sync_word(path,source,report=sync_report)
 except ConversionError as exc:
  status=load(sync_report)
  problems.append('Word-to-LaTeX synchronization failed: '+str(exc))
 checks['sync']={'pending_word_blocks':status['changed_blocks'],'source_matches_generation':source.read_bytes().decode('utf8')==manifest['source'],'conflicts':status['conflicts']}
 if status['changed_blocks']:problems.append('Saved Word changes await sync')
 if not checks['sync']['source_matches_generation']:problems.append('LaTeX changes await build')
 if baseline:
  original=load(Path(config['adapter'])/'imported.json');oldpath=Path(original['original'])
  if digest(oldpath)!=original['original_sha256']:problems.append('Original Word changed since adoption')
  old=Document(oldpath);checks['baseline']={'original_unchanged':digest(oldpath)==original['original_sha256'],'visible_text_identical':text==''.join(old.element.xpath('.//w:t/text() | .//m:t/text()'))}
  for key,selector in [('math','.//m:oMath'),('tables','.//w:tbl'),('section_properties','.//w:sectPr')]:
   # Baseline mode checks semantic structure; Word's default math serialization
   # and XML declaration spelling are not content differences.
   if key=='math':
    from roundtrip_model import math_model
    same=[math_model(e) for e in old.element.xpath(selector)]==[math_model(e) for e in document.element.xpath(selector)]
   elif key=='tables':
    from roundtrip_model import signature,doc_relationships
    same=[signature([e],doc_relationships(old)) for e in old.element.xpath(selector)]==[signature([e],doc_relationships(document)) for e in document.element.xpath(selector)]
   else:
    def section(e):return [[c.tag,dict(c.attrib)] for c in e if E.QName(c).localname in ('pgSz','pgMar','cols','type','pgNumType')]
    same=[section(e) for e in old.element.xpath(selector)]==[section(e) for e in document.element.xpath(selector)]
   checks['baseline'][key+'_identical']=same
  if not all(checks['baseline'].values()):problems.append('Initial adoption preservation check failed')
  oldfields=[v for p in old.element.xpath('.//w:p') for k,v,f in field_parts(p) if k=='field']
  checks['baseline']['native_field_codes_identical']=oldfields==fields
  if oldfields!=fields:problems.append('Native field instructions changed')
  with zipfile.ZipFile(oldpath) as a,zipfile.ZipFile(path) as b:
   media=[n for n in a.namelist() if n.startswith('word/media/')]
   checks['baseline']['media_bytes_identical']=all(n in b.namelist() and a.read(n)==b.read(n) for n in media)
   if not checks['baseline']['media_bytes_identical']:problems.append('Template image bytes changed')
 if digest(path)!=verified_word_hash or digest(source)!=verified_source_hash:
  problems.append('Word or LaTeX changed while checking; the new version was not acknowledged')
 result={'status':'passed' if not problems else 'needs_attention','document':str(path),'checks':checks,'problems':problems,'elapsed_seconds':time.perf_counter()-started,
         'verified_word_sha256':verified_word_hash,'verified_source_sha256':verified_source_hash}
 dump(result_report,result)
 if baseline and result['status']=='passed':dump(project/'adoption_acceptance.json',result)
 return result

if __name__=='__main__':
 p=argparse.ArgumentParser(description=__doc__);p.add_argument('project',type=Path);p.add_argument('--baseline',action='store_true');p.add_argument('--word',action='store_true');a=p.parse_args()
 try:r=check_project(a.project,baseline=a.baseline,word=a.word)
 except (ConversionError,ValueError) as exc:p.exit(2,str(exc)+'\n')
 print(json.dumps(r,ensure_ascii=False,indent=2));raise SystemExit(0 if r['status']=='passed' else 1)
