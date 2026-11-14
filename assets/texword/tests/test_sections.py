"""Section-level edits retain unchanged native objects and support structural edits."""
import json,sys,unittest
from pathlib import Path
from unittest.mock import patch
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import test_adoption as fixtures
import texword as tw
from docx import Document
from docx.oxml import OxmlElement
from roundtrip_model import W,NS,read_manifest
from check_project import check_project
from workflow_runtime import synchronize
from section_mapping import merge_source_edits,_restore_saved_emphasis
from roundtrip_model import signature


class SavedEmphasisTests(unittest.TestCase):
 def fixture(self):
  d=Document();d.styles['Normal'].font.bold=True
  p=d.add_paragraph();p.add_run('表题').bold=True
  record={'units':[{'node_signatures':[signature([p._p],{})]}]}
  p.runs[0].bold=None
  return d,p,record
 def test_inherited_bold_restores_exact_baseline(self):
  d,p,r=self.fixture();elements=[p._p]
  _restore_saved_emphasis(d,elements,r,{})
  self.assertEqual(signature(elements,{}),r['units'][0]['node_signatures'][0])
 def test_real_text_or_explicit_format_changes_are_retained(self):
  for change in ('text','bold'):
   d,p,r=self.fixture()
   if change=='text':p.runs[0].text='修改后的表题'
   else:p.runs[0].bold=False
   elements=[p._p];before=signature(elements,{})
   _restore_saved_emphasis(d,elements,r,{})
   self.assertEqual(signature(elements,{}),before)
 def test_table_restores_header_without_changing_body(self):
  d=Document();d.styles['Normal'].font.bold=True;t=d.add_table(rows=2,cols=1)
  t.cell(0,0).text='表头';t.cell(1,0).text='正文'
  t.cell(0,0).paragraphs[0].runs[0].bold=True
  r={'units':[{'node_signatures':[signature([t._tbl],{})]}]}
  t.cell(0,0).paragraphs[0].runs[0].bold=None;elements=[t._tbl]
  _restore_saved_emphasis(d,elements,r,{})
  self.assertEqual(signature(elements,{}),r['units'][0]['node_signatures'][0])


class SourceMergeTests(unittest.TestCase):
 def test_disjoint_changes_in_one_sentence(self):
  value,skipped=merge_source_edits('温度100，时间60。','温度110，时间60。','温度100，时间80。')
  self.assertEqual(value,'温度110，时间80。');self.assertEqual(skipped,[])
 def test_overlapping_replacement_uses_word(self):
  value,skipped=merge_source_edits('温度100。','温度110。','温度120。')
  self.assertEqual(value,'温度120。');self.assertEqual(skipped[0]['proposed'],'110')
 def test_identical_edits_are_not_conflicts_or_duplicates(self):
  self.assertEqual(merge_source_edits('甲乙','甲新增乙','甲新增乙'),('甲新增乙',[]))
 def test_conflicting_insertions_use_word(self):
  value,skipped=merge_source_edits('甲乙','甲助手乙','甲用户乙')
  self.assertEqual(value,'甲用户乙');self.assertTrue(skipped)
 def test_math_is_an_atomic_source_unit(self):
  value,skipped=merge_source_edits(r'式\(T=1\)。',r'式\(T=2\)。',r'式\(T=3\)。')
  self.assertEqual(value,r'式\(T=3\)。');self.assertTrue(skipped)


class AcceptedReviewMergeTests(unittest.TestCase):
 def test_accepted_review_wins_changed_local_proposal_after_relocation(self):
  raw=r'\twrevision{a}{旧}{用户已接受的建议}{理由}'
  local=r'\twrevision{a}{旧}{助手后来改写}{理由}'
  value,skipped=merge_source_edits(raw,'移到这里：'+local,raw,
      accepted_reviews={'a':{'raw':raw,'source':'用户已接受的建议'}})
  self.assertEqual(value,'移到这里：用户已接受的建议')
  self.assertEqual(skipped[0]['retained_word'],['用户已接受的建议'])


class SectionTests(unittest.TestCase):
 def setUp(self):
  self.f=fixtures.AdoptionTests();self.f.setUp();f=self.f
  profile=tw.load(f.adapter/'profile.json');profile['source_mapping_granularity']='section';tw.dump(f.adapter/'profile.json',profile)
  f.build();tw.dump(f.root/'workflow.json',{'source':str(f.source),'adapter':str(f.adapter),'code':str(f.code),'word':str(f.out)})
  self.bridge=patch('workflow_runtime.word_session',return_value={'open':False,'saved':None});self.bridge.start()
  self.check=patch('check_project.check_project',side_effect=lambda root,**kw:check_project(root,**dict(kw,word=False)));self.check.start()
 def tearDown(self):self.check.stop();self.bridge.stop();self.f.tearDown()
 def controls(self,d):return [e for e in d.element.body if e.tag==W+'sdt']
 def text(self,e):return ''.join(t.text or '' for t in e.iter(W+'t'))
 def paragraph(self,text,level=None):
  p=OxmlElement('w:p')
  if level:
   pr=OxmlElement('w:pPr');style=OxmlElement('w:pStyle');style.set(W+'val','Heading'+str(level));pr.append(style);p.append(pr)
  r=OxmlElement('w:r');t=OxmlElement('w:t');t.text=text;r.append(t);p.append(r);return p
 def test_user_column_widths_survive_body_and_table_rebuilds(self):
  f=self.f;d=Document(f.out);table=d.element.xpath('.//w:sdt//w:tbl')[0]
  widths=[1250,5750]
  for column,width in zip(table.findall('w:tblGrid/w:gridCol',NS),widths):column.set(W+'w',str(width))
  for row in table.findall('w:tr',NS):
   for cell,width in zip(row.findall('w:tc',NS),widths):cell.find('w:tcPr/w:tcW',NS).set(W+'w',str(width))
  table.find('w:tblPr/w:tblW',NS).set(W+'type','dxa');table.find('w:tblPr/w:tblW',NS).set(W+'w',str(sum(widths)))
  d.save(f.out)
  for old,new in [('原来的正文','助手修改的正文'),('现有描述','助手修改的表格内容')]:
   f.source.write_text(f.source.read_text(encoding='utf8').replace(old,new),encoding='utf8')
   result=synchronize(f.root,'push')
   self.assertEqual(result['status'],'passed');self.assertTrue(result['writeback_completed'])
   rebuilt=Document(f.out).element.xpath('.//w:sdt//w:tbl')[0]
   self.assertEqual([int(x.get(W+'w')) for x in rebuilt.findall('w:tblGrid/w:gridCol',NS)],widths)
   self.assertEqual(int(rebuilt.find('w:tblPr/w:tblW',NS).get(W+'w')),sum(widths))
   for row in rebuilt.findall('w:tr',NS):
    self.assertEqual([int(c.find('w:tcPr/w:tcW',NS).get(W+'w')) for c in row.findall('w:tc',NS)],widths)
 def test_native_blank_paragraphs_survive_rebuild_and_deletion(self):
  f=self.f;d=Document(f.out);content=self.controls(d)[0].find('w:sdtContent',NS)
  paragraph=next(e for e in content if '原来的正文' in self.text(e))
  paragraph.addnext(self.paragraph(''))
  d.element.body.insert(0,self.paragraph(''))
  d.element.body.insert(len(d.element.body)-1,self.paragraph(''))
  d.save(f.out)
  def count_empty(doc):
   return sum(1 for p in doc.element.xpath('.//w:p') if not self.text(p).strip() and not p.xpath('.//w:drawing | .//m:oMath'))
  expected=count_empty(d)
  f.source.write_text(f.source.read_text('utf8').replace('原来的正文','修改后的正文'),encoding='utf8')
  self.assertEqual(synchronize(f.root,'push')['status'],'passed')
  d=Document(f.out);self.assertEqual(count_empty(d),expected)
  empty=next(p for p in d.element.xpath('.//w:sdt/w:sdtContent/w:p') if not self.text(p).strip())
  empty.getparent().remove(empty);d.save(f.out)
  f.source.write_text(f.source.read_text('utf8').replace('现有描述','修改后的表格'),encoding='utf8')
  self.assertEqual(synchronize(f.root,'push')['status'],'passed')
  self.assertEqual(count_empty(Document(f.out)),expected-1)
 def test_template_generated_blank_is_not_restored_after_user_deletion(self):
  from word_runtime import TemplateDocument,layout_blank
  native=Document(self.f.out)
  runtime=object.__new__(TemplateDocument)
  runtime._capture_spacing(native,read_manifest(native))
  runtime.doc=Document(self.f.out);runtime.body=runtime.doc.element.body
  content=self.controls(runtime.doc)[0].find('w:sdtContent',NS)
  left=next(e for e in content if '原来的正文' in self.text(e))
  unexpected=self.paragraph('');left.addnext(unexpected)
  runtime._restore_spacing()
  self.assertIsNone(unexpected.getparent())
  self.assertFalse(any(layout_blank(e) for e in content))
 def test_native_equation_gap_survives_word_default_math_changes(self):
  from word_runtime import TemplateDocument,layout_blank
  from lxml import etree as E
  m='{'+NS['m']+'}'
  for saved_gap in (False,True):
   with self.subTest(saved_gap=saved_gap):
    native=Document(self.f.out);content=self.controls(native)[0].find('w:sdtContent',NS)
    lead=self.paragraph('公式为');content.append(lead)
    if saved_gap:content.append(self.paragraph(''))
    for name in ('x','y'):
     p=self.paragraph('');math=E.SubElement(p,m+'oMathPara');math=E.SubElement(math,m+'oMath')
     sup=E.SubElement(math,m+'sSup');base=E.SubElement(sup,m+'e');r=E.SubElement(base,m+'r');E.SubElement(r,m+'t').text=name
     exp=E.SubElement(sup,m+'sup');r=E.SubElement(exp,m+'r');E.SubElement(r,m+'t').text='2';content.append(p)
    runtime=object.__new__(TemplateDocument);runtime._capture_spacing(native,read_manifest(native))
    runtime.doc=native;runtime.body=native.element.body
    # Simulate a template gap plus defaults Word may add to two consecutive
    # formulas: exact signatures differ, but their equations remain equivalent.
    for p in list(content):
     if layout_blank(p):content.remove(p)
    lead.addnext(self.paragraph(''))
    for sup in native.element.xpath('.//m:sSup'):
     pr=E.Element(m+'sSupPr');default=E.SubElement(pr,m+'supHide');default.set(m+'val','0');sup.insert(0,pr)
    # The real report has the lead-in and equation in separate source controls.
    next_control=OxmlElement('w:sdt');next_content=OxmlElement('w:sdtContent');next_control.append(next_content)
    self.controls(native)[0].addnext(next_control)
    for p in list(content):
     if p.find('m:oMathPara',NS) is not None:next_content.append(p)
    from roundtrip_model import math_model
    formula=next_content[0];old_math=[math_model(m) for m in formula.findall('.//m:oMath',NS)]
    baseline=signature([formula],{})
    exponent=formula.xpath('.//m:sup//m:t')[0];exponent.text='3'
    self.assertNotEqual(signature([formula],{},math_baselines=old_math),baseline)
    exponent.text='2'
    runtime._restore_spacing()
    from word_runtime import layout_flow
    flow=list(layout_flow(native.element.body));i=flow.index(lead)
    self.assertEqual(layout_blank(flow[i+1]),saved_gap)
 def test_multi_equation_gap_keeps_each_native_semantic_label(self):
  from roundtrip import sync_word
  f=self.f
  fragment=r'''\section{Model}
Lead text.
\begin{equation}a=1\label{eq:a}\end{equation}
Between equations.
\begin{equation}b=2\label{eq:b}\end{equation}
See Eqs.~\ref{eq:a} and~\ref{eq:b}.
'''
  existing=f.source.read_text('utf8')
  f.source.write_text(existing.replace(r'\begin{thebibliography}',fragment+r'\begin{thebibliography}',1),encoding='utf8')
  f.build();doc=Document(f.out)
  for node in doc.element.xpath('.//m:t'):
   if node.text=='1':node.text='3'
   elif node.text=='2':node.text='4'
  for node in doc.element.xpath('.//w:t'):
   if 'Between equations.' in (node.text or ''):node.text='Changed connecting explanation.'
  doc.save(f.out);out=f.root/'multiple_equations.tex'
  sync_word(f.out,f.source,output=out)
  result=out.read_text('utf8')
  self.assertIn(r'\label{eq:a}',result);self.assertIn(r'\label{eq:b}',result)
  self.assertEqual(result.count(r'\begin{equation}'),2)
  self.assertIn('3',result);self.assertIn('4',result)
 def test_template_inherited_bold_retains_content_protection(self):
  from roundtrip import sync_word
  from roundtrip_model import attach_manifest,doc_relationships
  f=self.f;d=Document(f.out);manifest=read_manifest(d)
  cover=d.paragraphs[0];d.styles['Normal'].font.bold=True
  cover.runs[0].bold=True
  manifest['outside_nodes'][0]=signature([cover._p],doc_relationships(d))
  attach_manifest(d,manifest);cover.runs[0].bold=None;d.save(f.out)
  self.assertEqual(sync_word(f.out,f.source)['changed_blocks'],0)
  for change in ('text','bold'):
   d=Document(f.out);cover=d.paragraphs[0]
   if change=='text':cover.runs[0].text='USER COVER EDIT'
   else:cover.runs[0].bold=False
   d.save(f.out)
   with self.assertRaises(tw.ConversionError):sync_word(f.out,f.source)
   if change=='text':
    d=Document(f.out);d.paragraphs[0].runs[0].text='COVER';d.save(f.out)
 def test_sections_replace_paragraph_controls_without_content_loss(self):
  d=Document(self.f.out);m=read_manifest(d)
  self.assertEqual(len(self.controls(d)),3)
  self.assertEqual(m['granularity'],'section')
  self.assertEqual(len(d.element.xpath('.//w:sdt//w:sdt')),0)
  self.assertIn('原来的正文',tw.document_text(self.f.out))
  self.assertEqual(check_project(self.f.root)['status'],'passed')
 def test_native_cover_edit_survives_source_only_import_and_body_rebuild(self):
  f=self.f;doc=Document(f.out);doc.paragraphs[0].runs[0].text='USER COVER';doc.save(f.out)
  result=synchronize(f.root,'pull',source_only=True)
  self.assertTrue(tw.load(f.root/'workflow.json')['pending_template_refresh'])
  self.assertFalse(result['rebuilt_word'])
  f.source.write_text(f.source.read_text('utf8').replace('现有描述','新的正文'),encoding='utf8')
  result=synchronize(f.root,'push')
  self.assertEqual(result['status'],'passed')
  self.assertEqual(Document(f.out).paragraphs[0].text,'USER COVER')
  self.assertIn('新的正文',tw.document_text(f.out))
  self.assertNotIn('pending_template_refresh',tw.load(f.root/'workflow.json'))
  self.assertEqual(check_project(f.root)['status'],'passed')
  f.source.write_text(f.source.read_text('utf8').replace('新的正文','再次修改'),encoding='utf8')
  synchronize(f.root,'push')
  self.assertEqual(Document(f.out).paragraphs[0].text,'USER COVER')
 def test_cover_image_and_new_footer_relationships_survive_body_rebuild(self):
  from base64 import b64decode
  from docx.shared import Cm
  f=self.f;image=f.root/'cover.png'
  image.write_bytes(b64decode('iVBORw0KGgoAAAANSUhEUgAAABQAAAAMCAIAAADtbgqsAAAAHElEQVR4nGP8z0A+YKJAL8OoZhIBE6kaGEa4ZgD9bwEXwe2qmwAAAABJRU5ErkJggg=='))
  doc=Document(f.out);doc.paragraphs[0].add_run().add_picture(str(image),width=Cm(1))
  doc.sections[0].footer.paragraphs[0].text='USER FOOTER';doc.save(f.out)
  result=synchronize(f.root,'push')
  self.assertEqual(result['status'],'passed')
  doc=Document(f.out);self.assertEqual(doc.sections[0].footer.paragraphs[0].text,'USER FOOTER')
  self.assertEqual(len(doc.paragraphs[0]._p.xpath('.//w:drawing')),1)
  self.assertEqual(check_project(f.root)['status'],'passed')
 def test_insert_delete_and_merge_paragraphs_preserve_table_and_citation(self):
  f=self.f;d=Document(f.out);content=self.controls(d)[0].find('w:sdtContent',NS)
  original=next(e for e in content if '原来的正文' in self.text(e))
  added=self.paragraph('新增加的段落。');content.insert(content.index(original)+1,added)
  d.save(f.out);synchronize(f.root,'pull')
  self.assertIn('新增加的段落。',f.source.read_text(encoding='utf8'))
  d=Document(f.out);content=self.controls(d)[0].find('w:sdtContent',NS)
  added=next(e for e in content if '新增加的段落。' in self.text(e));content.remove(added)
  original=next(e for e in content if '原来的正文' in self.text(e));extra=self.paragraph('合并后的内容。')
  for r in list(extra):original.append(r)
  d.save(f.out);synchronize(f.root,'pull')
  value=f.source.read_text(encoding='utf8');self.assertNotIn('新增加的段落。',value)
  self.assertIn('合并后的内容。',value);self.assertIn(r'\cite{BIB_one}',value)
  self.assertEqual(len(Document(f.out).element.xpath('.//w:tbl')),1)
 def test_new_section_outside_control_then_delete_whole_section(self):
  f=self.f;d=Document(f.out);second=self.controls(d)[1];i=d.element.body.index(second)
  d.element.body.insert(i,self.paragraph('新增章节',1));d.element.body.insert(i+1,self.paragraph('新增章节的正文。'));d.save(f.out)
  synchronize(f.root,'pull');value=f.source.read_text(encoding='utf8')
  self.assertIn(r'\section{新增章节}',value);self.assertIn('新增章节的正文。',value)
  d=Document(f.out);self.assertEqual(len(self.controls(d)),4)
  added=next(e for e in self.controls(d) if '新增章节' in self.text(e));d.element.body.remove(added);d.save(f.out)
  synchronize(f.root,'pull');self.assertNotIn('新增章节',f.source.read_text(encoding='utf8'))
  self.assertEqual(len(self.controls(Document(f.out))),3)
 def test_new_heading_inside_section_is_split_on_rebuild(self):
  f=self.f;d=Document(f.out);content=self.controls(d)[0].find('w:sdtContent',NS)
  content.append(self.paragraph('新增小节',2));content.append(self.paragraph('新增小节中的文字。'));d.save(f.out)
  synchronize(f.root,'pull');self.assertIn(r'\subsection{新增小节}',f.source.read_text(encoding='utf8'))
  self.assertEqual(len(self.controls(Document(f.out))),4)
 def test_reordered_sections_follow_word_order(self):
  f=self.f;d=Document(f.out);controls=self.controls(d);last=controls[-1]
  d.element.body.remove(last);d.element.body.insert(d.element.body.index(controls[0]),last);d.save(f.out)
  synchronize(f.root,'pull');value=f.source.read_text(encoding='utf8')
  self.assertLess(value.index('第二章'),value.index('第一章'))
 def test_word_added_empty_carriers_are_not_text_edits(self):
  f=self.f;d=Document(f.out);second=self.controls(d)[1];d.element.body.insert(d.element.body.index(second),self.paragraph(''));d.save(f.out)
  self.assertEqual(check_project(f.root)['status'],'passed')
 def test_same_section_disjoint_changes_are_merged(self):
  f=self.f;d=Document(f.out);t=next(t for t in d.element.xpath('.//w:sdt//w:t') if '原来的正文' in (t.text or ''));t.text=t.text.replace('原来的正文','Word修改');d.save(f.out)
  f.source.write_text(f.source.read_text(encoding='utf8').replace('现有描述','LaTeX修改'),encoding='utf8')
  result=synchronize(f.root,'pull');self.assertEqual(result['status'],'passed')
  value=f.source.read_text('utf8');self.assertIn('Word修改',value);self.assertIn('LaTeX修改',value)
 def test_inserted_table_renumbers_original_caption_and_preserves_native_table(self):
  f=self.f
  new=r'\begin{table}\caption{新增表}\label{tab:new}\begin{tabular}{ll}甲 & 乙\\一 & 二\\\end{tabular}\end{table}'+'\n\n'
  value=f.source.read_text(encoding='utf8').replace(r'\begin{table}',new+r'\begin{table}',1)
  f.source.write_text(value,encoding='utf8');synchronize(f.root,'push')
  d=Document(f.out)
  texts=[''.join(t.text or '' for t in p.iter(W+'t')) for p in d.element.xpath('.//w:sdtContent/w:p')]
  self.assertTrue(any('表2' in t and '示例表格' in t for t in texts))
  self.assertEqual(len(d.element.xpath('.//w:tbl')),2)
 def test_upgrade_imports_edits_and_a_deleted_old_paragraph(self):
  f=self.f;profile=tw.load(f.adapter/'profile.json');profile.pop('source_mapping_granularity');tw.dump(f.adapter/'profile.json',profile);f.build()
  d=Document(f.out);old=next(e for e in self.controls(d) if '原来的正文' in self.text(e));d.element.body.remove(old)
  title=next(t for t in d.element.xpath('.//w:sdt//w:t') if '第二章' in (t.text or ''));title.text=title.text.replace('第二章','修改后的章名');d.save(f.out)
  profile['source_mapping_granularity']='section';tw.dump(f.adapter/'profile.json',profile)
  synchronize(f.root,'pull');value=f.source.read_text(encoding='utf8')
  self.assertNotIn('原来的正文',value);self.assertIn('修改后的章名',value)
  self.assertEqual(read_manifest(Document(f.out))['granularity'],'section')


if __name__=='__main__':unittest.main()
