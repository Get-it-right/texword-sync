"""Exercise repeated editing cycles and conflict preservation on a tiny document."""
import contextlib,io,sys,subprocess
from pathlib import Path
from unittest.mock import patch
import unittest
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import test_adoption as adoption_fixture
from docx import Document
import texword as tw
from check_project import check_project
from workflow_runtime import synchronize
from roundtrip_model import NS,W,signature,read_manifest
from lxml import etree as E


class HyperlinkNormalizationTests(unittest.TestCase):
 def test_native_and_word_field_links_preserve_url_and_caption(self):
  ns='xmlns:w="'+NS['w']+'" xmlns:r="'+NS['r']+'"'
  native=E.fromstring(('<w:p '+ns+'>'+tw.run('前文')+'<w:hyperlink r:id="rId1">'+tw.run('链接')+'</w:hyperlink>'+tw.run('后文')+'</w:p>').encode())
  field=E.fromstring(('<w:p '+ns+'>'+tw.run('前文')+tw.field('HYPERLINK "https://example.org"','链接')+tw.run('后文')+'</w:p>').encode())
  rels={'rId1':'https://example.org'}
  self.assertEqual(signature([native],rels),signature([field],rels))
  changed=E.fromstring(E.tostring(field));next(t for t in changed.iter(W+'t') if t.text=='链接').text='不同说明'
  self.assertNotEqual(signature([native],rels),signature([changed],rels))
  self.assertNotEqual(signature([native],{'rId1':'https://different.org'}),signature([field],rels))
  compared=E.fromstring(('<w:p '+ns+'>'+tw.run('前文')+tw.field('HYPERLINK "https://example.org" \\h','链接')+tw.run('后文')+'</w:p>').encode())
  self.assertEqual(signature([native],rels),signature([compared],rels))
  self.assertNotEqual(signature([native],{'rId1':'https://different.org'}),signature([compared],rels))


class WorkflowTests(unittest.TestCase):
 def setUp(self):
  self.fixture=adoption_fixture.AdoptionTests();self.fixture.setUp();self.fixture.build()
  f=self.fixture
  tw.dump(f.root/'workflow.json',{'source':str(f.source),'adapter':str(f.adapter),
          'code':str(f.code),'word':str(f.out),'synchronized_word_sha256':tw.digest(f.out)})
  self.bridge=patch('workflow_runtime.word_session',side_effect=self.word_session);self.bridge.start()
  self.check=patch('check_project.check_project',side_effect=lambda project,**kw:check_project(project,**dict(kw,word=False)));self.check.start()
 def tearDown(self):
  self.check.stop();self.bridge.stop();self.fixture.tearDown()
 def word_session(self,project,action,expected_hash=None):
  if action=='Inspect':return {'open':False,'saved':None}
  if action=='Close':self.assertEqual(expected_hash,tw.digest(self.fixture.out))
 def edit_word(self,before,after):
  d=Document(self.fixture.out)
  t=next(t for t in d.element.xpath('.//w:sdt//w:t') if before in (t.text or ''))
  t.text=t.text.replace(before,after);d.save(self.fixture.out)
 def test_repeated_pull_refreshes_baseline_and_keeps_citations(self):
  f=self.fixture
  for before,after in [('原来的正文','第一次修改'),('第一次修改','第二次修改')]:
   self.edit_word(before,after);result=synchronize(f.root,'pull')
   self.assertEqual(result['imported_word_blocks'],1)
   self.assertIn(after,f.source.read_text(encoding='utf8'))
   self.assertIn(r'\cite{BIB_one}',f.source.read_text(encoding='utf8'))
   self.assertIn(after,tw.document_text(f.out))
  self.assertFalse(synchronize(f.root,'pull')['rebuilt_word'])
 def test_push_merges_independent_word_and_source_edits(self):
  f=self.fixture;self.edit_word('原来的正文','Word改正文')
  f.source.write_text(f.source.read_text(encoding='utf8').replace('第一章','源码改标题'),encoding='utf8')
  result=synchronize(f.root,'push');self.assertEqual(result['status'],'passed')
  for text in ('Word改正文','源码改标题'):
   self.assertIn(text,tw.document_text(f.out));self.assertIn(text,f.source.read_text(encoding='utf8'))
 def test_build_applies_adapter_changes_without_source_edits(self):
  f=self.fixture;source=f.source.read_bytes()
  adapter=f.adapter/'template_adapter.py'
  adapter.write_text(adapter.read_text(encoding='utf8')+'\n_original_init = Report.__init__\ndef _format_refresh(self):\n _original_init(self)\n self.doc.core_properties.comments = "adapter format refresh"\nReport.__init__ = _format_refresh\n',encoding='utf8')
  self.assertTrue(synchronize(f.root,'build')['rebuilt_word'])
  self.assertEqual(Document(f.out).core_properties.comments,'adapter format refresh')
  self.assertEqual(f.source.read_bytes(),source)
 def test_overlap_keeps_word_and_reports_skipped_latex(self):
  f=self.fixture;self.edit_word('原来的正文','Word改正文')
  f.source.write_text(f.source.read_text(encoding='utf8').replace('原来的正文','源码改正文'),encoding='utf8')
  result=synchronize(f.root,'push')
  self.assertEqual(result['status'],'passed');self.assertTrue(result['skipped_latex_edits'])
  self.assertIn('Word改正文',tw.document_text(f.out));self.assertNotIn('源码改正文',tw.document_text(f.out))
 def test_failed_generation_preserves_current_word_and_can_retry(self):
  f=self.fixture;self.edit_word('原来的正文','Word改正文');before=f.out.read_bytes();source=f.source.read_bytes()
  real_run=subprocess.run
  def run(command,**kw):
   if len(command)>1 and command[1]==str(f.code):return subprocess.CompletedProcess(command,1,'','test generation failure')
   return real_run(command,**kw)
  with patch('workflow_runtime.subprocess.run',side_effect=run):
   with self.assertRaises(tw.ConversionError):synchronize(f.root,'pull')
  self.assertEqual(before,f.out.read_bytes())
  self.assertIn('Word改正文',f.source.read_text(encoding='utf8'))
  self.assertEqual(synchronize(f.root,'pull')['status'],'passed')
  self.assertIn('Word改正文',tw.document_text(f.out))
 def test_unsaved_word_requests_window_before_save_import_or_generation(self):
  f=self.fixture;word=f.out.read_bytes();source=f.source.read_bytes()
  with patch('workflow_runtime.word_session',return_value={'open':True,'saved':False,'content_sha256':'draft'}), patch('workflow_runtime.compile_code') as compile_, patch('workflow_runtime.sync_word') as sync_, patch('workflow_runtime.time.sleep'):
   result=synchronize(f.root,'push')
  self.assertEqual(result['status'],'needs_write_window')
  compile_.assert_not_called();sync_.assert_not_called()
  self.assertEqual(word,f.out.read_bytes());self.assertEqual(source,f.source.read_bytes())
 def test_saved_changes_in_open_word_are_detected_before_import(self):
  self.edit_word('原来的正文','Word改正文');f=self.fixture
  with patch('workflow_runtime.word_session',return_value={'open':True,'saved':True}), patch('workflow_runtime.sync_word') as sync_, patch('workflow_runtime.time.sleep'):
   result=synchronize(f.root,'push')
  self.assertEqual(result['status'],'needs_write_window');self.assertTrue(result['saved_word_changed']);sync_.assert_not_called()
 def test_same_length_active_edits_are_detected_by_content_hash(self):
  with patch('workflow_runtime.word_session',side_effect=[{'open':True,'saved':False,'content_sha256':'first'},{'open':True,'saved':False,'content_sha256':'second'}]), patch('workflow_runtime.time.sleep'):
   result=synchronize(self.fixture.root,'push',write_window=True)
  self.assertEqual(result['status'],'needs_write_window');self.assertTrue(result['changing'])
 def test_confirmed_window_imports_latest_saved_word_before_generation(self):
  self.edit_word('原来的正文','Word改正文')
  def session(project,action,expected_hash=None):
   if action=='Inspect':return {'open':True,'saved':True,'content_sha256':'stable'}
   return self.word_session(project,action,expected_hash)
  with patch('workflow_runtime.word_session',side_effect=session), patch('workflow_runtime.time.sleep'):
   result=synchronize(self.fixture.root,'push',write_window=True)
  self.assertEqual(result['status'],'passed');self.assertIn('Word改正文',tw.document_text(self.fixture.out))
 def test_new_edits_after_import_stop_before_compile(self):
  f=self.fixture;f.source.write_text(f.source.read_text(encoding='utf8').replace('第一章','源码改标题'),encoding='utf8')
  reads=0
  def session(project,action,expected_hash=None):
   nonlocal reads
   if action=='Inspect':
    reads+=1;return {'open':True,'saved':reads<=2,'content_sha256':'old' if reads<=2 else 'new'}
   return self.word_session(project,action,expected_hash)
  with patch('workflow_runtime.word_session',side_effect=session), patch('workflow_runtime.compile_code') as compile_, patch('workflow_runtime.time.sleep'):
   result=synchronize(f.root,'push',write_window=True)
  self.assertEqual(result['status'],'needs_write_window');compile_.assert_not_called()
 def test_unchanged_saved_open_word_does_not_need_a_window(self):
  with patch('workflow_runtime.word_session',return_value={'open':True,'saved':True,'content_sha256':'stable'}), patch('workflow_runtime.time.sleep'):
   result=synchronize(self.fixture.root,'push')
  self.assertEqual(result['status'],'passed');self.assertFalse(result['rebuilt_word'])

 def test_noop_delivers_pending_skips_before_clearing_them(self):
  f=self.fixture;config=tw.load(f.root/'workflow.json')
  skipped=[{'proposed':'被覆盖的建议','retained_word':['用户正文']}]
  config['pending_merge_skips']=skipped;tw.dump(f.root/'workflow.json',config)
  result=synchronize(f.root,'push')
  self.assertEqual(result['skipped_latex_edits'],skipped)
  self.assertEqual(tw.load(f.root/'workflow_result.json')['skipped_latex_edits'],skipped)
  self.assertNotIn('pending_merge_skips',tw.load(f.root/'workflow.json'))

 def test_changed_mapping_granularity_does_not_take_noop_path(self):
  f=self.fixture;profile=tw.load(f.adapter/'profile.json');profile['source_mapping_granularity']='section'
  tw.dump(f.adapter/'profile.json',profile)
  self.assertEqual(synchronize(f.root,'pull')['status'],'passed')
  self.assertEqual(read_manifest(Document(f.out))['granularity'],'section')

 def test_source_only_cli_rejects_other_actions_before_word_operations(self):
  from workflow_runtime import main
  for action in ('status','preflight','check','push'):
   with self.subTest(action=action),patch('workflow_runtime.word_session') as session,contextlib.redirect_stderr(io.StringIO()):
    with self.assertRaises(SystemExit) as error:main(self.fixture.root,[action,'--source-only'])
    self.assertEqual(error.exception.code,2);session.assert_not_called()

 def test_source_only_mapping_change_stops_before_save_or_close(self):
  f=self.fixture;profile=tw.load(f.adapter/'profile.json');profile['source_mapping_granularity']='section'
  tw.dump(f.adapter/'profile.json',profile)
  with patch('workflow_runtime.word_session',return_value={'open':False,'saved':None}) as session:
   with self.assertRaisesRegex(tw.ConversionError,'full pull'):synchronize(f.root,'pull',source_only=True)
  session.assert_not_called()

 def test_source_only_import_does_not_interrupt_active_word_or_generate(self):
  f=self.fixture;self.edit_word('原来的正文','已保存的用户修改');word=f.out.read_bytes()
  with patch('workflow_runtime.word_session',side_effect=AssertionError('Preparation must not touch desktop Word')) as session,patch('workflow_runtime.compile_code') as compile_:
   result=synchronize(f.root,'pull',source_only=True)
  session.assert_not_called();compile_.assert_not_called()
  self.assertIn('已保存的用户修改',f.source.read_text(encoding='utf8'))
  self.assertEqual(word,f.out.read_bytes());self.assertFalse(result['writeback_completed'])
  self.assertEqual(result['word_scope'],'saved_file_only')

 def test_source_only_noop_keeps_pending_skips_for_final_writeback(self):
  f=self.fixture;config=tw.load(f.root/'workflow.json')
  skipped=[{'proposed':'未写入的建议','retained_word':['用户正文']}]
  config['pending_merge_skips']=skipped;tw.dump(f.root/'workflow.json',config)
  with patch('workflow_runtime.word_session',side_effect=AssertionError('No desktop operations')):
   result=synchronize(f.root,'pull',source_only=True)
  self.assertEqual(result['skipped_latex_edits'],skipped)
  self.assertEqual(tw.load(f.root/'workflow.json')['pending_merge_skips'],skipped)
  self.assertFalse(result['writeback_completed'])

 def test_preparation_records_the_word_identity_actually_imported(self):
  from workflow_runtime import import_word_edits
  f=self.fixture;self.edit_word('原来的正文','先保存的正文')
  def import_latest(*args,**kwargs):
   self.edit_word('先保存的正文','导入前刚保存的正文')
   return import_word_edits(*args,**kwargs)
  with patch('workflow_runtime.import_word_edits',side_effect=import_latest),patch('workflow_runtime.word_session',side_effect=AssertionError('No desktop operations')):
   result=synchronize(f.root,'pull',source_only=True)
  self.assertEqual(result['status'],'passed')
  self.assertIn('导入前刚保存的正文',f.source.read_text(encoding='utf8'))
  config=tw.load(f.root/'workflow.json')
  self.assertEqual(config['synchronized_word_sha256'],tw.digest(f.out))
  self.assertEqual(config['import_receipt']['word_sha256'],tw.digest(f.out))

 def test_one_confirmed_push_saves_latest_edits_merges_and_reports_completion(self):
  f=self.fixture;synchronize(f.root,'pull',source_only=True)
  f.source.write_text(f.source.read_text(encoding='utf8').replace('第一章','准备好的标题'),encoding='utf8')
  saved=False;actions=[]
  def session(project,action,expected_hash=None):
   nonlocal saved
   actions.append(action)
   if action=='Inspect':return {'open':True,'saved':saved,'content_sha256':'stable'}
   if action=='Save':
    self.edit_word('原来的正文','窗口确认前最新正文');saved=True
   return self.word_session(project,action,expected_hash)
  with patch('workflow_runtime.word_session',side_effect=session),patch('workflow_runtime.time.sleep'):
   result=synchronize(f.root,'push',write_window=True)
  self.assertEqual(actions.count('Save'),1);self.assertEqual(actions.count('Close'),1)
  self.assertTrue(result['writeback_completed']);self.assertIn('resume editing',result['next_action'])
  for text in ('准备好的标题','窗口确认前最新正文'):
   self.assertIn(text,tw.document_text(f.out));self.assertIn(text,f.source.read_text(encoding='utf8'))

 def test_one_confirmed_push_keeps_user_overlap_and_delivers_skips(self):
  f=self.fixture
  self.edit_word('原来的正文','参数100')
  synchronize(f.root,'pull',source_only=True)
  f.source.write_text(f.source.read_text(encoding='utf8').replace('参数100','参数200'),encoding='utf8')
  def session(project,action,expected_hash=None):
   if action=='Inspect':return {'open':True,'saved':True,'content_sha256':'stable'}
   return self.word_session(project,action,expected_hash)
  self.edit_word('参数100','参数300')
  with patch('workflow_runtime.word_session',side_effect=session),patch('workflow_runtime.time.sleep'):
   result=synchronize(f.root,'push',write_window=True)
  self.assertTrue(result['writeback_completed']);self.assertTrue(result['skipped_latex_edits'])
  self.assertIn('参数300',tw.document_text(f.out));self.assertNotIn('参数200',tw.document_text(f.out))
  self.assertEqual(tw.load(f.root/'workflow_result.json')['skipped_latex_edits'],result['skipped_latex_edits'])
  self.assertNotIn('pending_merge_skips',tw.load(f.root/'workflow.json'))


if __name__=='__main__':unittest.main()
