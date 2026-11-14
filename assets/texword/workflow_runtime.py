"""Fixed daily commands for the mapped Word/LaTeX editing cycle."""
from __future__ import annotations
import argparse,json,os,subprocess,sys,tempfile,time
from pathlib import Path
from texword import ROOT,ConversionError,compile_code,digest,dump,load
from roundtrip import sync_word
from roundtrip_model import read_manifest
from docx import Document


def _run_generator(code,output,*,template_regions_from=None):
 """Run an existing self-contained generator for a temporary candidate."""
 command=[sys.executable,str(code),str(output),'--overwrite']
 if template_regions_from is not None:command.extend(['--template-regions-from',str(template_regions_from)])
 result=subprocess.run(command,
                       capture_output=True,text=True,encoding='utf8',errors='replace')
 if result.returncode:raise ConversionError(result.stderr.strip() or result.stdout.strip())


def _acknowledge_check(project,config,checked):
 """Acknowledge only the exact saved pair checked by the readback owner."""
 if checked.get('status')!='passed':raise ConversionError('Cannot acknowledge an unsuccessful check')
 word=Path(config['word']);source=Path(config['source'])
 if digest(word)!=checked['verified_word_sha256'] or digest(source)!=checked['verified_source_sha256']:
  raise ConversionError('Word or LaTeX changed after checking; the new version has not been acknowledged')
 config['synchronized_word_sha256']=checked['verified_word_sha256']
 config.pop('import_receipt',None)
 dump(project/'workflow.json',config)


def _publish_completed_result(project,config,result,*,delivered):
 """Persist skipped suggestions before clearing their pending delivery state."""
 result.setdefault('skipped_latex_edits',list(config.get('pending_merge_skips',[])))
 result['writeback_completed']=delivered
 result['next_action']=('Word writeback is complete; resume editing.' if delivered else
                        'Finish the LaTeX edits, then run preflight. After a requested pause, run push --write-window once.')
 dump(project/'workflow_result.json',result)
 if delivered and 'pending_merge_skips' in config:
  config.pop('pending_merge_skips');dump(project/'workflow.json',config)
 return result


def word_session(project,action,expected_hash=None):
 command=['C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe',
          '-NoProfile','-ExecutionPolicy','Bypass','-File',str(ROOT/'word_session.ps1'),
          '-ProjectPath',str(project),'-Action',action]
 if expected_hash:command+=['-ExpectedHash',expected_hash]
 r=subprocess.run(command,capture_output=True,text=True,encoding='utf8',errors='replace')
 if r.returncode:raise ConversionError(r.stderr.strip() or r.stdout.strip())
 if action=='Inspect':
  try:
   result=json.loads(r.stdout)
   if not isinstance(result.get('open'),bool):raise ValueError('Missing Word open state')
   if result['open'] and not isinstance(result.get('saved'),bool):raise ValueError('Missing Word saved state')
   return result
  except (ValueError,TypeError,AttributeError) as exc:
   raise ConversionError('Could not inspect the current Word editing state: '+r.stdout.strip()) from exc


def write_preflight(project,*,write_window=False,expected_hash=None,settle_seconds=0.4):
 """Inspect twice without saving; request a write window before expensive work."""
 project=Path(project).resolve();config=load(project/'workflow.json');word=Path(config['word'])
 def snapshot():
  state=word_session(project,'Inspect')
  if not isinstance(state,dict) or not isinstance(state.get('open'),bool):
   raise ConversionError('Word inspection did not return a valid editing state.')
  return dict(state,word_sha256=digest(word))
 first=snapshot()
 if settle_seconds:time.sleep(settle_seconds)
 current=snapshot();changing=first!=current
 unsaved=current['open'] and not current.get('saved',False)
 changed=current['word_sha256']!=config.get('synchronized_word_sha256')
 after_import=expected_hash is not None
 reasons=[]
 if changing:reasons.append('Word content or saved file changed during the preflight.')
 if after_import:
  if unsaved:reasons.append('Word has new unsaved edits after import.')
  if current['word_sha256']!=expected_hash:reasons.append('The saved Word file changed after import.')
 elif current['open'] and not write_window:
  if unsaved:reasons.append('Word has unsaved edits.')
  if changed:reasons.append('Word has newly saved changes that need to be imported in a write window.')
 result={'status':'needs_write_window' if reasons else 'ready','phase':'before_generation',
         'word':str(word),'word_open':current['open'],'unsaved_edits':unsaved,
         'saved_word_changed':changed,'changing':changing,'write_window_confirmed':write_window,
         'word_sha256':current['word_sha256'],'reasons':reasons,
         'next_action':'Save Word, pause editing, then run push --write-window.' if reasons else 'Continue synchronization.'}
 return result


def _wait_for_write_window(project,result,action):
 result=dict(result,action=action,rebuilt_word=False,word_replaced=False)
 dump(Path(project)/'workflow_result.json',result)
 return result


def upgrade_sections(project,config,word_hash):
 """Retag the edited old Word without first importing deleted paragraph tags."""
 from roundtrip_model import signature,doc_relationships,attach_manifest
 from section_mapping import group_sections,_key
 from review_comments import read_comments
 word=Path(config['word']);edited=Document(word);old=read_manifest(edited)
 outside=[e for e in edited.element.body if _key(e) is None]
 if signature(outside,doc_relationships(edited))!=old['outside_baseline']:
  raise ConversionError('Template changed before mapping upgrade; no content was replaced.')
 with tempfile.TemporaryDirectory(prefix='.mapping_upgrade_',dir=project) as temporary:
  baseline=Path(temporary)/'baseline.docx'
  _run_generator(config['code'],baseline)
  original=Document(baseline);prior=read_manifest(original)
  if prior['source']!=old['source'] or [r['id'] for r in prior['spans']]!=[r['id'] for r in old['spans']]:
   raise ConversionError('Generated baseline does not match the old Word mapping; upgrade stopped.')
  base=group_sections(original,prior,doc_relationships(original),read_comments(original))
  actual=group_sections(edited,old,doc_relationships(edited),read_comments(edited))
  if set(r['id'] for r in actual['spans'])-set(r['id'] for r in base['spans']):
   raise ConversionError('A heading changed structurally during mapping upgrade; preserve the current Word for reconciliation.')
  base['granularity']='section'
  for key in ('outside_baseline','outside_reviews','outside_nodes','outside_prefix_count','outside_suffix_count'):base[key]=actual[key]
  attach_manifest(edited,base);staged=Path(temporary)/word.name;edited.save(staged)
  word_session(project,'Close',word_hash)
  if digest(word)!=word_hash:raise ConversionError('Word changed during mapping upgrade; it has not been replaced.')
  os.replace(staged,word)
 return digest(word)


def synchronize(project,action,*,write_window=False,source_only=False):
 """A successful atomic import survives a later build failure for safe retry."""
 if source_only and action!='pull':raise ConversionError('source_only is supported only for pull')
 project=Path(project).resolve();config=load(project/'workflow.json');state={'phase':'preflight','word_replaced':False}
 try:return _synchronize(project,action,write_window=write_window,source_only=source_only,state=state)
 except Exception as exc:
  failed=dict(state,status='failed',action=action,error=str(exc),word=config['word'],source=config['source'])
  for name in ('word','source'):
   path=Path(config[name]);failed[name+'_sha256']=digest(path) if path.exists() else None
  dump(project/'workflow_result.json',failed)
  raise


def import_word_edits(project,config,word,source,*,source_only=False,apply=True):
 """Preview or import saved edits using the last common Word baseline.

 Legacy signatures are compared against a pristine generated copy. This
 operation never opens, closes or rewrites Word; apply controls LaTeX writes.
 """
 manifest=read_manifest(Document(word));receipt=config.get('import_receipt',{})
 common=receipt.get('word_bases') if receipt.get('manifest_source_sha256')==manifest['source_sha256'] else None
 options=dict(apply=apply,report=project/'sync_check.json',merge_bases=common,capture_baselines=apply,allow_structure=not source_only,allow_template_updates=True)
 if manifest.get('semantic_version',1)>=2:
  return sync_word(word,source,**options)
 with tempfile.TemporaryDirectory(prefix='.semantic_baseline_',dir=project) as temporary:
  baseline=Path(temporary)/'baseline.docx'
  _run_generator(config['code'],baseline)
  return sync_word(word,source,baseline_word=baseline,**options)


def _synchronize(project,action,*,write_window=False,source_only=False,state=None):
 """Merge both sides, rebase mapping, and show the checked result in Word."""
 from check_project import check_project
 project=Path(project).resolve();config=load(project/'workflow.json')
 state={} if state is None else state
 word=Path(config['word']);source=Path(config['source'])
 requested=load(Path(config['adapter'])/'profile.json').get('source_mapping_granularity','block')
 # Preparation reads only the saved package. The user can keep typing in Word;
 # sync_word guards both disk identities, and the final push imports newer edits.
 if source_only:
  ready={'unsaved_edits':False}
 else:
  ready=write_preflight(project,write_window=write_window)
  if ready['status']!='ready':return _wait_for_write_window(project,ready,action)
 if action!='build' and not config.get('pending_template_refresh') and not ready['unsaved_edits'] and digest(word)==config.get('synchronized_word_sha256'):
  manifest=read_manifest(Document(word))
  if (manifest.get('semantic_version',1)>=2 and manifest.get('granularity','block')==requested
      and source.read_bytes().decode('utf8')==manifest['source']):
   receipt=load(project/'check_result.json') if (project/'check_result.json').exists() else {}
   reuse=(receipt.get('status')=='passed' and receipt.get('verified_word_sha256')==digest(word)
          and receipt.get('verified_source_sha256')==digest(source))
   if not reuse:
    receipt=check_project(project,word=False)
   if receipt['status']!='passed':raise ConversionError('Read-only synchronization check needs attention')
   result={'status':'passed','action':action,'word':str(word),'source':str(source),
           'imported_word_blocks':0,'rebuilt_word':False,'word_opened':False,
           'checks':'reused' if reuse else 'passed'}
   if source_only:result.update(source_only=True,word_scope='saved_file_only')
   return _publish_completed_result(project,config,result,delivered=not source_only)
 if source_only and read_manifest(Document(word)).get('granularity','block')!=requested:
  raise ConversionError('Mapping granularity changes require a full pull; source-only does not rewrite Word')
 state['phase']='import_saved_word' if source_only else 'save_and_import'
 if not source_only:word_session(project,'Save')
 word_hash=digest(word)
 if requested=='section' and read_manifest(Document(word)).get('granularity','block')!='section':
  word_hash=upgrade_sections(project,config,word_hash)
 # A completed import can precede a failed compilation. Do not import the same
 # Word against its old source baseline twice; retain the acknowledged hash.
 if word_hash==config.get('synchronized_word_sha256'):
  imported={'changed_blocks':0,'already_synchronized':True,
            'skipped_latex_edits':config.get('pending_merge_skips',[])}
 else:
  imported=import_word_edits(project,config,word,source,source_only=source_only)
  word_hash=imported['word_sha256']
  if imported.get('template_changed'):config['pending_template_refresh']=True
  config['pending_merge_skips']=config.get('pending_merge_skips',[])+imported.get('skipped_latex_edits',[])
  imported['skipped_latex_edits']=config['pending_merge_skips']
  config['synchronized_word_sha256']=word_hash
  config['import_receipt']={'word_sha256':word_hash,'source_sha256':digest(source),
                          'manifest_source_sha256':read_manifest(Document(word))['source_sha256'],
                          'word_bases':imported['word_bases']}
  dump(project/'workflow.json',config)
 source_hash=digest(source)
 if source_only:
  result={'status':'passed','action':action,'source_only':True,'word':str(word),'source':str(source),
          'imported_word_blocks':imported['changed_blocks'],'rebuilt_word':False,'word_opened':False,
          'skipped_latex_edits':imported.get('skipped_latex_edits',[]),'word_scope':'saved_file_only'}
  return _publish_completed_result(project,config,result,delivered=False)
 manifest=read_manifest(Document(word))
 rebuild=(action=='build' or config.get('pending_template_refresh',False) or source.read_bytes().decode('utf8')!=manifest['source'] or manifest.get('granularity','block')!=requested or manifest.get('semantic_version',1)<2)
 if rebuild:
  state['phase']='generate'
  ready=write_preflight(project,write_window=write_window,expected_hash=word_hash)
  if ready['status']!='ready':return _wait_for_write_window(project,ready,action)
  compile_code(config['adapter'],source,config['code'],bibliography=config.get('bibliography'),roundtrip=True,overwrite=True)
  # Complete generation before closing or replacing the user's current file.
  with tempfile.TemporaryDirectory(prefix='.word_build_',dir=project) as temporary:
   staged=Path(temporary)/word.name
   _run_generator(config['code'],staged,template_regions_from=word)
   if digest(source)!=source_hash:raise ConversionError('LaTeX changed during generation; Word has not been replaced. Run again.')
   state['phase']='validate_candidate';candidate_check=check_project(project,candidate=staged)
   if candidate_check['status']!='passed':
    dump(project/'check_result.json',dict(candidate_check,phase='before_replace'))
    raise ConversionError('Candidate readback failed before replacing Word; see '+str(project/'check_result.json'))
   if config.get('native_review'):
    from native_comparison import prepare_candidate
    state['phase']='native_review'
    state['native_review']=prepare_candidate(config,staged,word)
    candidate_check=check_project(project,candidate=staged)
    if candidate_check['status']!='passed':
     dump(project/'check_result.json',dict(candidate_check,phase='native_review_before_replace'))
     raise ConversionError('Native-revision candidate readback failed; current Word was preserved')
   state['phase']='replace';word_session(project,'Close',word_hash)
   if digest(word)!=word_hash:raise ConversionError('Word changed before replacement; run again.')
   # Word may release its file handle shortly after Close returns. Keep the
   # original race guard during a bounded retry rather than regenerate.
   deadline=time.monotonic()+5
   while True:
    if digest(word)!=word_hash:raise ConversionError('Word changed before replacement; run again.')
    try:os.replace(staged,word);break
    except PermissionError:
     if time.monotonic()>=deadline:raise
     time.sleep(0.2)
   state['word_replaced']=True
   receipt=load(staged.with_suffix('.validation.json'));receipt['output']=str(word)
   dump(word.with_suffix('.validation.json'),receipt)
  config['synchronized_word_sha256']=digest(word);config.pop('pending_template_refresh',None)
  if config.get('native_review'):config['native_review'].pop('initial_refresh_word_sha256',None)
  dump(project/'workflow.json',config)
 # Opens the same configured document, refreshes native fields, and checks the
 # source map. Rebuilding after pull gives the next Word edit a fresh baseline.
 state['phase']='refresh_and_check';checked=check_project(project,word=rebuild)
 if checked['status']!='passed':raise ConversionError('Synchronization needs attention; see '+str(project/'check_result.json'))
 _acknowledge_check(project,config,checked)
 result={'status':'passed','action':action,'word':str(word),'source':str(source),
         'imported_word_blocks':imported['changed_blocks'],'rebuilt_word':rebuild,
         'word_opened':rebuild,'checks':'passed',
         'skipped_latex_edits':imported.get('skipped_latex_edits',[])}
 if 'native_review' in checked['checks']:result['native_review']=checked['checks']['native_review']
 return _publish_completed_result(project,config,result,delivered=True)


def main(project,argv=None):
 project=Path(project).resolve()
 # A project may select a compatible frontend without replacing system Pandoc.
 frontend=load(project/'workflow.json').get('pandoc')
 if frontend:os.environ.setdefault('PANDOC',frontend)
 parser=argparse.ArgumentParser(description=__doc__)
 parser.add_argument('action',choices=['pull','push','status','sync','build','update','check','preflight'],
                     help='pull: Word edits to LaTeX; push: LaTeX edits to Word. Both merge saved edits and refresh the mapping.')
 parser.add_argument('--baseline',action='store_true',help='Initial adoption comparison only')
 parser.add_argument('--word',action='store_true',help='Open Word and refresh native fields when checking')
 parser.add_argument('--write-window',action='store_true',help='Use only after the user has saved and paused editing for this writeback.')
 parser.add_argument('--source-only',action='store_true',help='pull only: import the saved Word file; no desktop Word calls or editing pause')
 args=parser.parse_args(argv)
 if args.source_only and args.action!='pull':parser.error('--source-only requires pull')
 try:
  if args.action=='preflight':
   result=write_preflight(project,write_window=args.write_window)
   dump(project/'workflow_result.json',result)
  elif args.action=='check':
   from check_project import check_project
   result=check_project(project,baseline=args.baseline,word=args.word)
   if args.word and result['status']=='passed':
    config=load(project/'workflow.json');_acknowledge_check(project,config,result)
    _publish_completed_result(project,config,dict(result,action='check',word=config['word'],source=config['source'],word_opened=True,rebuilt_word=False),delivered=True)
  elif args.action=='status':
   config=load(project/'workflow.json')
   result=import_word_edits(project,config,Path(config['word']),Path(config['source']),apply=False)
  else:
   # Exclusive creation prevents two fixed commands from replacing the same
   # current document concurrently. A stale lock is reported, never ignored.
   lock=project/'.workflow.lock'
   try:handle=lock.open('x',encoding='utf8')
   except FileExistsError:raise ConversionError('A synchronization is already active, or its lock needs recovery: '+str(lock))
   try:
    with handle:handle.write(str(os.getpid()))
    result=synchronize(project,args.action,write_window=args.write_window,source_only=args.source_only)
   finally:lock.unlink()
  print(json.dumps(result,ensure_ascii=False,indent=2))
  return 0 if result['status'] in ('passed','preview','ready','needs_write_window') else 1
 except (ConversionError,PermissionError) as exc:
  parser.exit(2,str(exc)+'\n')
