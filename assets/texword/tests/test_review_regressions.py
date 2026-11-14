"""Behavioral regressions for the two independent workflow code reviews.

All documents are synthetic fixtures. No test connects to Microsoft Word.
"""
import copy
import subprocess
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from docx import Document
from docx.oxml import OxmlElement
from lxml import etree as E
import texword as tw
from adoption import simplify_latex
from check_project import check_project
from review_comments import reconcile_review_ids
from roundtrip import sync_word
from roundtrip_model import W,NS,math_model,signature,read_manifest,attach_manifest
from section_mapping import merge_source_edits
from tracked_reviews import bookmark_name
from workflow_runtime import synchronize
import test_sections
import test_tracked_reviews
import test_workflow


class ReviewPureRegressions(unittest.TestCase):
 def test_word_overlap_discards_entire_new_review(self):
  value,skipped=merge_source_edits('温度100。',r'\twrevision{k}{温度100。}{温度110。}{调整温度}','温度120。')
  self.assertEqual(value,'温度120。');self.assertTrue(skipped)

 def test_disjoint_title_edit_keeps_body_review_atomic(self):
  base='标题。正文内容。';local=r'标题。\twrevision{k}{正文内容。}{完善后的正文。}{理由}'
  value,skipped=merge_source_edits(base,local,'新标题。正文内容。')
  self.assertEqual(value,'新'+local);self.assertEqual(skipped,[])

 def test_existing_review_sides_cannot_cross_merge(self):
  base=r'\twrevision{k}{温度100。}{温度110。}{理由}'
  value,skipped=merge_source_edits(base,base.replace('110','115'),'温度120。')
  self.assertEqual(value,'温度120。');self.assertTrue(skipped)

 def test_duplicate_comment_identity_and_renumbering(self):
  baseline=[dict(id=str(i),anchor='的',text='请核对',author='User') for i in (1,2)]
  for ids in ((1,2),(40,41)):
   snapshot=copy.deepcopy(baseline)
   for item,cid in zip(snapshot,ids):item['id']=str(cid)
   normalized,origins=reconcile_review_ids(snapshot,baseline)
   self.assertEqual(normalized,baseline);self.assertEqual(len(origins),2)
  snapshot=copy.deepcopy(baseline);snapshot[1]['text']='第二条新意见'
  self.assertEqual(reconcile_review_ids(snapshot,baseline)[0][1]['id'],'2')
  with self.assertRaisesRegex(ValueError,'Ambiguous'):
   reconcile_review_ids(baseline+[dict(baseline[0],id='3')],baseline)
  for i,item in enumerate(baseline):item['context']=[str(i),'context']
  snapshot=copy.deepcopy(baseline)+[dict(baseline[0],id='3',context=['third','context'])]
  self.assertEqual([x['id'] for x in reconcile_review_ids(snapshot,baseline)[0]],['1','2','new:3'])

 def test_math_styles_are_semantic_but_default_fonts_are_not(self):
  def model(properties):
   return E.fromstring(('<m:oMath xmlns:m="'+NS['m']+'"><m:r><m:rPr>'+properties+'</m:rPr><m:t>x</m:t></m:r></m:oMath>').encode())
  italic=model('');upright=model('<m:sty m:val="p"/>');bold=model('<m:sty m:val="b"/>')
  self.assertNotEqual(math_model(italic),math_model(upright))
  p=E.Element(W+'p');p.append(copy.deepcopy(upright))
  q=E.Element(W+'p');q.append(copy.deepcopy(bold))
  self.assertNotEqual(signature([p],{}),signature([q],{}))
  self.assertNotEqual(math_model(italic),math_model(model('<m:nor/>')))
  self.assertNotEqual(math_model(italic),math_model(model('<m:scr m:val="double-struck"/>')))
  self.assertEqual(math_model(italic),math_model(model('<m:sty m:val="i"/>')))
  combined=model('<m:sty m:val="p"/>');combined.find('.//{'+NS['m']+'}t').text='kg+1'
  split=E.fromstring(('<m:oMath xmlns:m="'+NS['m']+'"><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>kg</m:t></m:r><m:r><m:t>+1</m:t></m:r></m:oMath>').encode())
  self.assertEqual(math_model(combined),math_model(split))

 def test_layout_simplification_preserves_dash_semantics(self):
  source=r'中文—连接，English--title，--flag，file---id，$a--b$'
  self.assertEqual(simplify_latex(source),source)

 def test_equivalent_inline_delimiters_do_not_rewrite_literals_or_formulas(self):
  from roundtrip import same_inline_math_notation
  self.assertTrue(same_inline_math_notation(r'温度\(T\)，单位\(\mathrm{K}\)。',r'温度$T$，单位$\mathrm{K}$。'))
  self.assertFalse(same_inline_math_notation(r'\verb|\(T\)|',r'\verb|$T$|'))
  self.assertFalse(same_inline_math_notation(r'\(T\)',r'$\mathbf{T}$'))

 def test_equivalent_inline_delimiters_with_mixed_newlines(self):
  from roundtrip import same_inline_math_notation
  self.assertTrue(same_inline_math_notation('段落\r\n温度\\(T\\)。\r\n', '段落\n温度$T$。\r\n'))
  self.assertFalse(same_inline_math_notation('段落\r\n温度\\(T\\)。', '段落\n温度$P$。'))


class ReviewSectionRegressions(unittest.TestCase):
 def setUp(self):
  self.case=test_sections.SectionTests();self.case.setUp();self.f=self.case.f
 def tearDown(self):self.case.tearDown()
 def content(self,doc):return self.case.controls(doc)[0].find('w:sdtContent',NS)

 def test_unknown_block_cannot_silently_drop_user_text(self):
  doc=Document(self.f.out);control=OxmlElement('w:sdt');content=OxmlElement('w:sdtContent')
  content.append(self.case.paragraph('USER_BLOCK_CONTROL_TEXT'));control.append(content);self.content(doc).append(control)
  doc.save(self.f.out);before=self.f.source.read_bytes();word=self.f.out.read_bytes()
  with self.assertRaisesRegex(tw.ConversionError,'Section synchronization'):
   sync_word(self.f.out,self.f.source,apply=True)
  self.assertEqual(before,self.f.source.read_bytes());self.assertEqual(word,self.f.out.read_bytes())
  report=tw.load(self.f.out.with_suffix('.sync.json'))
  self.assertIn('container',str(report['conflicts']))

 def test_nested_cell_container_is_detected_and_preserved(self):
  f=self.f;doc=Document(f.out);cell=doc.element.xpath('.//w:sdt//w:tbl//w:tc')[-1]
  control=OxmlElement('w:sdt');content=OxmlElement('w:sdtContent')
  content.append(self.case.paragraph('NESTED_USER_TEXT'));control.append(content);cell.append(control)
  doc.save(f.out);source=f.source.read_bytes();word=f.out.read_bytes()
  with self.assertRaises(tw.ConversionError):sync_word(f.out,f.source,apply=True)
  self.assertIn('container',str(tw.load(f.out.with_suffix('.sync.json'))['conflicts']))
  self.assertEqual(source,f.source.read_bytes());self.assertEqual(word,f.out.read_bytes())

 def test_ordinary_prose_before_uncaptioned_table_is_preserved(self):
  doc=Document(self.f.out);content=self.content(doc)
  content.append(self.case.paragraph('表面温度随时间变化。'))
  table=doc.add_table(rows=2,cols=2)
  for i,row in enumerate(table.rows):
   for j,cell in enumerate(row.cells):cell.text='数据'+str(i)+str(j)
  content.append(table._tbl);doc.save(self.f.out)
  sync_word(self.f.out,self.f.source,apply=True);value=self.f.source.read_text('utf8')
  self.assertIn('表面温度随时间变化。',value)
  self.assertNotIn(r'\caption{面温度',value)

 def test_body_tab_keeps_words_separated(self):
  doc=Document(self.f.out);p=self.case.paragraph('甲')
  run=OxmlElement('w:r');run.append(OxmlElement('w:tab'));p.append(run)
  run=OxmlElement('w:r');text=OxmlElement('w:t');text.text='乙';run.append(text);p.append(run)
  self.content(doc).append(p);doc.save(self.f.out)
  sync_word(self.f.out,self.f.source,apply=True)
  self.assertRegex(self.f.source.read_text('utf8'),r'甲\s+乙')

 def test_math_style_only_edit_is_imported_and_rebuilds(self):
  f=self.f;f.source.write_text(f.source.read_text('utf8').replace('原来的正文',r'原来的正文\(x\)'),encoding='utf8');f.build()
  doc=Document(f.out);run=next(r for r in doc.element.xpath('.//m:r') if ''.join(r.xpath('./m:t/text()',namespaces=NS))=='x')
  properties=run.find('m:rPr',NS)
  if properties is None:properties=OxmlElement('m:rPr');run.insert(0,properties)
  style=properties.find('m:sty',NS)
  if style is None:style=OxmlElement('m:sty');properties.append(style)
  style.set('{'+NS['m']+'}val','b');doc.save(f.out)
  self.assertGreater(sync_word(f.out,f.source,apply=True)['changed_blocks'],0)
  self.assertRegex(f.source.read_text('utf8'),r'\\(?:mathbf|boldsymbol)\{x\}')
  f.build();self.assertEqual(sync_word(f.out,f.source)['changed_blocks'],0)

 def test_section_receipt_merges_subsequent_word_edits_without_dropping_review(self):
  f=self.f;doc=Document(f.out)
  next(t for t in doc.element.xpath('.//w:sdt//w:t') if t.text=='原来的正文').text='Word先改'
  doc.save(f.out);synchronize(f.root,'pull',source_only=True)
  f.source.write_text(f.source.read_text('utf8').replace('Word先改',r'\twrevision{parallel}{Word先改}{助手完善}{理由}'),encoding='utf8')
  doc=Document(f.out);next(t for t in doc.element.xpath('.//w:tbl//w:t') if t.text=='现有描述').text='Word后改表格'
  doc.save(f.out);result=synchronize(f.root,'push')
  self.assertIn(r'\twrevision{parallel}{Word先改}{助手完善}',f.source.read_text('utf8'))
  self.assertIn('Word后改表格',f.source.read_text('utf8'));self.assertEqual(result['skipped_latex_edits'],[])

 def test_source_only_refuses_structural_changes_without_partial_import(self):
  f=self.f;doc=Document(f.out);controls=self.case.controls(doc)
  doc.element.body.insert(doc.element.body.index(controls[1]),self.case.paragraph('新映射外段落'))
  doc.save(f.out);source=f.source.read_bytes();word=f.out.read_bytes()
  with self.assertRaises(tw.ConversionError):synchronize(f.root,'pull',source_only=True)
  self.assertEqual(source,f.source.read_bytes());self.assertEqual(word,f.out.read_bytes())
  self.assertIn('Structural',str(tw.load(f.root/'sync_check.json')['conflicts']))


class ReviewTrackedRegressions(unittest.TestCase):
 def setUp(self):self.case=test_tracked_reviews.TrackedReviewTests();self.case.setUp()
 def tearDown(self):self.case.tearDown()
 def add_comment(self,doc):
  start=next(e for e in doc.element.iter(W+'bookmarkStart') if e.get(W+'name')==bookmark_name('change-one'))
  parent=start.getparent();end=next(e for e in parent if e.tag==W+'bookmarkEnd' and e.get(W+'id')==start.get(W+'id'))
  c=doc.comments.add_comment(text='用户批注必须保留',author='User');cid=str(c.comment_id)
  a=OxmlElement('w:commentRangeStart');a.set(W+'id',cid);parent.insert(parent.index(start)+1,a)
  b=OxmlElement('w:commentRangeEnd');b.set(W+'id',cid);parent.insert(parent.index(end),b)
  run=OxmlElement('w:r');ref=OxmlElement('w:commentReference');ref.set(W+'id',cid);run.append(ref);parent.insert(parent.index(end),run)

 def test_comment_on_pending_revision_preserves_both_by_refusing_overlap(self):
  f=self.case;doc=Document(f.word);self.add_comment(doc);doc.save(f.word)
  before=f.source.read_bytes();word=f.word.read_bytes()
  with self.assertRaisesRegex(tw.ConversionError,'overlaps pending'):
   sync_word(f.word,f.source,apply=True)
  self.assertEqual(before,f.source.read_bytes());self.assertEqual(word,f.word.read_bytes())

 def test_comment_on_accepted_revision_survives_import(self):
  f=self.case;f.decide(True);doc=Document(f.word);self.add_comment(doc);doc.save(f.word)
  sync_word(f.word,f.source,apply=True);value=f.source.read_text('utf8')
  self.assertIn('用户批注必须保留',value);self.assertIn('twcomment[User]',value)
  self.assertIn('建议采用的表述',value);self.assertNotIn('twrevision',value)

 def test_early_error_replaces_stale_report(self):
  f=self.case;report=f.root/'sync.json'
  tw.dump(report,{'conflicts':[],'changed_blocks':99,'obsolete':True})
  doc=Document(f.word);start=next(e for e in doc.element.iter(W+'bookmarkStart') if e.get(W+'name','').startswith('TWREV_'))
  start.getparent().remove(start);doc.save(f.word)
  with self.assertRaisesRegex(tw.ConversionError,'bookmark missing'):
   sync_word(f.word,f.source,report=report)
  current=tw.load(report);self.assertNotIn('obsolete',current)
  self.assertIn('bookmark missing',str(current['conflicts']));self.assertEqual(current['word_sha256'],tw.digest(f.word))
  report.unlink()
  with self.assertRaises(tw.ConversionError):sync_word(f.word,f.source,report=report)
  self.assertIn('bookmark missing',tw.load(report)['error'])


class ReviewRuntimeRegressions(unittest.TestCase):
 def setUp(self):self.case=test_workflow.WorkflowTests();self.case.setUp();self.f=self.case.fixture
 def tearDown(self):self.case.tearDown()

 def test_noop_never_saves_closes_or_opens_word(self):
  calls=[]
  def session(project,action,expected_hash=None):
   calls.append(action);return {'open':True,'saved':True,'content_sha256':'stable'}
  with patch('workflow_runtime.word_session',side_effect=session):
   self.assertFalse(synchronize(self.f.root,'push')['rebuilt_word'])
  self.assertEqual(set(calls),{'Inspect'})

 def test_source_only_import_is_acknowledged_before_prose_edit(self):
  f=self.f;self.case.edit_word('原来的正文','Word改正文')
  with patch('workflow_runtime.compile_code') as compiler:
   result=synchronize(f.root,'pull',source_only=True)
  compiler.assert_not_called();self.assertTrue(result['source_only'])
  source=f.source.read_text('utf8').replace('第一章','新标题');f.source.write_text(source,encoding='utf8')
  with patch('workflow_runtime.sync_word',wraps=sync_word) as importer:
   result=synchronize(f.root,'push')
  # Candidate and final checks read back; the import phase does not reapply Word.
  self.assertFalse(any(c.kwargs.get('apply') for c in importer.call_args_list))
  self.assertEqual(result['status'],'passed');self.assertIn('Word改正文',f.source.read_text('utf8'))

 def test_new_word_edits_after_source_only_use_last_import_as_common_base(self):
  f=self.f;self.case.edit_word('原来的正文','Word先改')
  synchronize(f.root,'pull',source_only=True)
  value=f.source.read_text('utf8').replace('Word先改',r'\twrevision{parallel}{Word先改}{助手完善}{理由}')
  f.source.write_text(value,encoding='utf8');self.case.edit_word('现有描述','Word后改表格')
  result=synchronize(f.root,'push')
  value=f.source.read_text('utf8')
  self.assertIn(r'\twrevision{parallel}{Word先改}{助手完善}',value)
  self.assertIn('Word后改表格',value);self.assertEqual(result['skipped_latex_edits'],[])

 def test_failed_build_retains_concurrent_latex_and_successful_import(self):
  f=self.f;self.case.edit_word('原来的正文','Word改正文');word=f.out.read_bytes();real_run=subprocess.run
  def session(project,action,expected_hash=None):
   if action=='Save':f.source.write_text(f.source.read_text('utf8').replace('第一章','并发标题'),encoding='utf8')
   return self.case.word_session(project,action,expected_hash)
  def run(command,**kw):
   if len(command)>1 and command[1]==str(f.code):return subprocess.CompletedProcess(command,1,'','intentional failure')
   return real_run(command,**kw)
  with patch('workflow_runtime.word_session',side_effect=session),patch('workflow_runtime.subprocess.run',side_effect=run):
   with self.assertRaises(tw.ConversionError):synchronize(f.root,'push')
  self.assertEqual(word,f.out.read_bytes());value=f.source.read_text('utf8')
  self.assertIn('并发标题',value);self.assertIn('Word改正文',value)

 def test_failed_candidate_never_closes_or_replaces_current_word(self):
  f=self.f;f.source.write_text(f.source.read_text('utf8').replace('第一章','新标题'),encoding='utf8')
  word=f.out.read_bytes();calls=[]
  def checker(project,**kw):
   if kw.get('candidate') is not None:return {'status':'needs_attention'}
   return check_project(project,word=False)
  def session(project,action,expected_hash=None):
   calls.append(action);return self.case.word_session(project,action,expected_hash)
  with patch('check_project.check_project',side_effect=checker),patch('workflow_runtime.word_session',side_effect=session):
   with self.assertRaisesRegex(tw.ConversionError,'Candidate readback'):
    synchronize(f.root,'push')
  self.assertNotIn('Close',calls);self.assertEqual(word,f.out.read_bytes())

 def test_unreadable_generated_mapping_is_rejected_before_real_replace(self):
  f=self.f;f.source.write_text(f.source.read_text('utf8').replace('第一章','新标题'),encoding='utf8')
  word=f.out.read_bytes();calls=[];real_run=subprocess.run
  def run(command,**kw):
   result=real_run(command,**kw)
   if len(command)>2 and command[1]==str(f.code) and result.returncode==0:
    candidate=Path(command[2]);doc=Document(candidate)
    doc.element.xpath('.//w:sdtPr/w:tag')[0].set(W+'val','texword:unknown')
    doc.save(candidate)
   return result
  def session(project,action,expected_hash=None):
   calls.append(action);return self.case.word_session(project,action,expected_hash)
  with patch('workflow_runtime.subprocess.run',side_effect=run),patch('workflow_runtime.word_session',side_effect=session):
   with self.assertRaisesRegex(tw.ConversionError,'Candidate readback'):synchronize(f.root,'push')
  self.assertNotIn('Close',calls);self.assertEqual(word,f.out.read_bytes())
  self.assertEqual(tw.load(f.root/'check_result.json')['phase'],'before_replace')

 def test_legacy_mapping_uses_pristine_baseline_without_losing_word_edit(self):
  f=self.f;doc=Document(f.out);manifest=read_manifest(doc);manifest.pop('semantic_version')
  attach_manifest(doc,manifest);doc.save(f.out);self.case.edit_word('原来的正文','旧版Word修改')
  self.assertEqual(synchronize(f.root,'push')['status'],'passed')
  self.assertIn('旧版Word修改',f.source.read_text('utf8'))
  self.assertEqual(read_manifest(Document(f.out))['semantic_version'],2)

 def test_legacy_readonly_check_does_not_rewrite_word(self):
  f=self.f;doc=Document(f.out);manifest=read_manifest(doc);manifest.pop('semantic_version')
  attach_manifest(doc,manifest);doc.save(f.out);before=f.out.read_bytes()
  self.assertEqual(check_project(f.root,word=False)['status'],'passed')
  self.assertEqual(before,f.out.read_bytes())

 def test_postcheck_user_save_is_not_acknowledged(self):
  f=self.f;f.source.write_text(f.source.read_text('utf8').replace('第一章','新标题'),encoding='utf8')
  def checker(project,**kw):
   result=check_project(project,**dict(kw,word=False))
   if kw.get('candidate') is None:
    doc=Document(f.out);t=next(t for t in doc.element.xpath('.//w:sdt//w:t') if '原来的正文' in (t.text or ''))
    t.text+=' USER_SAVE';doc.save(f.out)
   return result
  with patch('check_project.check_project',side_effect=checker):
   with self.assertRaisesRegex(tw.ConversionError,'has not been acknowledged'):
    synchronize(f.root,'push')
  self.assertIn('USER_SAVE',tw.document_text(f.out))
  self.assertNotEqual(tw.load(f.root/'workflow.json')['synchronized_word_sha256'],tw.digest(f.out))


if __name__=='__main__':unittest.main()
