"""Behavioral tests for source-preserving Word -> LaTeX patches."""
import copy,json,subprocess,sys,unittest,zipfile
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import texword as tw
import test_texword as fixtures
from roundtrip import sync_word,ensure_import_preamble
from roundtrip_model import NS,W,M,read_manifest,math_model
from docx import Document
from docx.oxml import OxmlElement
from lxml import etree as E

class ImportPreambleTests(unittest.TestCase):
 def test_new_table_and_comment_have_minimal_idempotent_dependencies(self):
  source=r'''\documentclass{article}
% \usepackage{longtable} is not active.
\usepackage{booktabs,array}
\begin{document}
\twcomment{note}{Unchanged anchor}{Check the sensor.}
\begin{longtable}{l}\toprule Value\\\bottomrule\end{longtable}
\end{document}'''
  updated,added=ensure_import_preamble(source)
  self.assertEqual(added,[r'\usepackage{longtable}',r'\providecommand{\twcomment}[4][Codex]{#3}'])
  self.assertIn(r'\twcomment{note}{Unchanged anchor}{Check the sensor.}',updated)
  self.assertEqual(ensure_import_preamble(updated),(updated,[]))

 def test_existing_preamble_and_crlf_are_preserved(self):
  source='\\documentclass{article}\r\n\\usepackage{array,longtable,booktabs}\r\n\\begin{document}\r\n\\begin{longtable}{l}x\\\\\\end{longtable}\r\n\\end{document}\r\n'
  self.assertEqual(ensure_import_preamble(source),(source,[]))

class RoundtripTests(unittest.TestCase):
 def setUp(self):
  self.fixture=fixtures.ConversionTests();self.fixture.setUp();self.root=self.fixture.root
  self.fixture.adapt();self.source=self.fixture.source;self.code=self.root/'build.py';self.word=self.root/'edited.docx'
  tw.compile_code(self.fixture.adapter,self.source,self.code,bibliography=self.fixture.bib,roundtrip=True)
  self.build()
 def tearDown(self):self.fixture.tearDown()
 def build(self,source=None,code=None,word=None):
  if source:tw.compile_code(self.fixture.adapter,source,code,bibliography=self.fixture.bib,roundtrip=True)
  subprocess.run([sys.executable,str(code or self.code),str(word or self.word),'--overwrite'],check=True,capture_output=True)
 def edit(self,old,new):
  d=Document(self.word)
  nodes=d.element.xpath('.//w:sdt/w:sdtContent//w:t')
  hit=next(t for t in nodes if old in (t.text or ''));hit.text=hit.text.replace(old,new);d.save(self.word)
 def test_no_change_preserves_exact_source_bytes(self):
  original=self.source.read_bytes();out=self.root/'unchanged.tex'
  result=sync_word(self.word,self.source,output=out)
  self.assertEqual(result['changed_blocks'],0);self.assertEqual(out.read_bytes(),original)
 def test_report_path_cannot_overwrite_source(self):
  original=self.source.read_bytes()
  with self.assertRaisesRegex(tw.ConversionError,'separate .json'):sync_word(self.word,self.source,report=self.source)
  self.assertEqual(self.source.read_bytes(),original)
 def test_text_heading_table_math_and_citations_roundtrip(self):
  self.edit('正文含','修改后的正文含');self.edit('转换验证','修改后的标题');self.edit('冰与液态水','冰与未冻水')
  d=Document(self.word);maths=d.element.xpath('.//w:sdt/w:sdtContent//m:t')
  next(t for t in maths if t.text=='ρ').text='2ρ';d.save(self.word)
  out=self.root/'mapped.tex';result=sync_word(self.word,self.source,output=out)
  value=out.read_text(encoding='utf8')
  self.assertEqual(result['changed_blocks'],4)
  self.assertIn('修改后的正文含',value);self.assertIn('修改后的标题',value);self.assertIn('冰与未冻水',value)
  self.assertIn(r'\section{修改后的标题}',value)
  self.assertIn(r'\cite{pennes}',value);self.assertIn(r'\label{tab:test}',value);self.assertIn('2',value)
  self.build(out,self.root/'rebuilt.py',self.root/'rebuilt.docx')
  newdoc=Document(self.root/'rebuilt.docx');text=''.join(newdoc.element.xpath('.//w:t/text() | .//m:t/text()'))
  self.assertIn('修改后的正文含',text);self.assertIn('冰与未冻水',text)
 def test_field_cache_updates_do_not_change_source(self):
  d=Document(self.word)
  for t in d.element.xpath('.//w:sdt/w:sdtContent//w:r[w:fldChar[@w:fldCharType="separate"]]/following-sibling::w:r[1]/w:t'):t.text='999'
  d.save(self.word);out=self.root/'cache.tex'
  result=sync_word(self.word,self.source,output=out)
  self.assertEqual(result['changed_blocks'],0);self.assertEqual(out.read_bytes(),self.source.read_bytes())
 def test_overlap_writes_word_priority_output_and_skips_without_mutating_input(self):
  self.edit('正文含','Word修改含');self.source.write_text(fixtures.SOURCE.replace('正文含','LaTeX修改含'),encoding='utf8')
  original=self.source.read_bytes();out=self.root/'conflict.tex'
  result=sync_word(self.word,self.source,output=out)
  self.assertTrue(out.exists());self.assertEqual(original,self.source.read_bytes())
  self.assertIn('Word修改含',out.read_text('utf8'));self.assertNotIn('LaTeX修改含',out.read_text('utf8'))
  self.assertTrue(result['skipped_latex_edits'])
 def test_disjoint_source_and_word_edits_merge(self):
  self.edit('正文含','Word修改含');self.source.write_text(fixtures.SOURCE.replace('水分数模型','LaTeX修改模型'),encoding='utf8')
  sync_word(self.word,self.source,apply=True);value=self.source.read_text(encoding='utf8')
  self.assertIn('Word修改含',value);self.assertIn('LaTeX修改模型',value)
 def test_removed_tag_and_unmapped_insertions_rejected(self):
  d=Document(self.word);d.add_paragraph('Unmapped text must not vanish');d.save(self.word)
  with self.assertRaises(tw.ConversionError):sync_word(self.word,self.source)
  self.assertIn('outside',tw.load(self.word.with_suffix('.sync.json'))['conflicts'][0]['reason'])
 def test_new_paragraph_inside_source_block_imports(self):
  d=Document(self.word);control=next(c for c in d.element.xpath('.//w:sdtContent') if '正文含' in ''.join(E._Element.xpath(c,'.//w:t/text()',namespaces=NS)))
  p=OxmlElement('w:p');r=OxmlElement('w:r');t=OxmlElement('w:t');t.text='追加的段落。';r.append(t);p.append(r);control.append(p);d.save(self.word)
  out=self.root/'inserted.tex';sync_word(self.word,self.source,output=out)
  self.assertIn('追加的段落。',out.read_text(encoding='utf8'))
 def test_cross_reference_edit_keeps_semantics(self):
  self.edit('公式','变更公式');out=self.root/'references.tex';sync_word(self.word,self.source,output=out)
  value=out.read_text(encoding='utf8');self.assertIn(r'\eqref{eq:heat}',value);self.assertIn(r'\ref{tab:test}',value)
  self.assertNotIn('TWRT',value)
 def test_unchanged_inline_math_keeps_its_original_latex_when_prose_changes(self):
  self.edit('正文含','修改正文含');out=self.root/'math-preserved.tex';sync_word(self.word,self.source,output=out)
  self.assertIn(r'$Q_e=\sigma|\nabla\phi|^2$',out.read_text(encoding='utf8'))
 def test_word_math_reserialization_is_not_a_content_edit(self):
  d=Document(self.word)
  for fraction in d.element.xpath('.//m:f'):
   for prop in list(fraction):
    if prop.tag==M+'fPr':fraction.remove(prop)
  for root in d.element.xpath('.//m:oMath'):
   for run in root.findall('.//m:r',NS):
    prop=OxmlElement('w:rPr');font=OxmlElement('w:rFonts');font.set(W+'ascii','Cambria Math');prop.append(font);run.insert(0,prop)
  # Word also emits top-level hidden bookmarks when updating fields.
  bookmark=OxmlElement('w:bookmarkStart');bookmark.set(W+'name','_RefTest');bookmark.set(W+'id','9922');d.element.body.insert(1,bookmark)
  d.save(self.word);out=self.root/'word-normalized.tex';result=sync_word(self.word,self.source,output=out)
  self.assertEqual(result['changed_blocks'],0);self.assertEqual(out.read_bytes(),self.source.read_bytes())
 def test_changed_block_with_source_comment_is_not_silently_rewritten(self):
  self.source.write_text(fixtures.SOURCE.replace('正文含','% retain this comment\n正文含'),encoding='utf8')
  tw.compile_code(self.fixture.adapter,self.source,self.code,bibliography=self.fixture.bib,overwrite=True,roundtrip=True);self.build()
  self.edit('正文含','Word修改正文含')
  with self.assertRaises(tw.ConversionError):sync_word(self.word,self.source,apply=True)
  self.assertIn('retain this comment',self.source.read_text(encoding='utf8'))
 def test_source_preamble_and_unedited_comments_are_preserved(self):
  value=self.source.read_bytes().replace(b'\\begin{document}',b'% exact preamble comment\r\n\\begin{document}')
  self.source.write_bytes(value);tw.compile_code(self.fixture.adapter,self.source,self.code,bibliography=self.fixture.bib,overwrite=True,roundtrip=True);self.build()
  self.edit('正文含','修改正文含');out=self.root/'comments.tex';sync_word(self.word,self.source,output=out)
  self.assertIn(b'% exact preamble comment\r\n',out.read_bytes())

if __name__=='__main__':unittest.main()
