"""Native comments and pending suggestions must never become approved prose."""
import os,subprocess,sys,unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import texword as tw
import test_texword as fixtures
from roundtrip import sync_word
from review_comments import read_comments
from docx import Document
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement
from lxml import etree as E
from roundtrip_model import W,NS

SOURCE=r'''\documentclass{article}
\begin{document}
\section{批注验证}
直接写入的确定内容。\twsuggest{efficiency}{原来的表述}{建议采用的表述}{需要人工确认的理由。}
另有\twcomment{evidence}{待补证据的论述}{请核对原始文献。}，文献\cite{pennes}。
\end{document}
'''

class ReviewTests(unittest.TestCase):
 def setUp(self):
  self.fixture=fixtures.ConversionTests();self.fixture.setUp();self.root=self.fixture.root
  self.fixture.adapt();self.source=self.fixture.source;self.source.write_text(SOURCE,encoding='utf8')
  self.code=self.root/'review_build.py';self.word=self.root/'review.docx';self.build()
 def tearDown(self):self.fixture.tearDown()
 def build(self,roundtrip=True):
  tw.compile_code(self.fixture.adapter,self.source,self.code,bibliography=self.fixture.bib,roundtrip=roundtrip,overwrite=True)
  subprocess.run([sys.executable,str(self.code),str(self.word),'--overwrite'],check=True,capture_output=True,env=dict(os.environ,PANDOC='INTENTIONALLY_UNAVAILABLE'))
 def test_native_comments_leave_original_prose_unchanged(self):
  d=Document(self.word);body=tw.document_text(self.word);comments=read_comments(d)
  self.assertIn('原来的表述',body);self.assertIn('直接写入的确定内容',body)
  self.assertNotIn('建议采用的表述',body);self.assertNotIn('需要人工确认的理由',body)
  self.assertEqual(len(comments),2);self.assertIn('建议采用的表述',''.join(c['text'] for c in comments.values()))
  self.assertEqual(len(d.element.xpath('.//w:commentRangeStart')),2)
 def test_unchanged_comments_preserve_source_exactly(self):
  out=self.root/'no-change.tex';result=sync_word(self.word,self.source,output=out)
  self.assertEqual(result['changed_blocks'],0);self.assertEqual(out.read_bytes(),self.source.read_bytes())
 def test_word_renumbered_comment_ids_do_not_create_source_edits(self):
  d=Document(self.word)
  for c in d.comments:c._element.set(W+'id',str(c.comment_id+100))
  for e in d.element.xpath('.//w:commentRangeStart | .//w:commentRangeEnd | .//w:commentReference'):e.set(W+'id',str(int(e.get(W+'id'))+100))
  d.save(self.word);out=self.root/'renumbered.tex';result=sync_word(self.word,self.source,output=out)
  self.assertEqual(result['changed_blocks'],0);self.assertEqual(out.read_bytes(),self.source.read_bytes())
 def test_modified_comment_is_synced_as_pending_annotation(self):
  d=Document(self.word);c=list(d.comments)[1];c.paragraphs[0].text='请补充测量条件。';d.save(self.word)
  out=self.root/'comment-edit.tex';result=sync_word(self.word,self.source,output=out);value=out.read_text(encoding='utf8')
  self.assertEqual(result['changed_blocks'],1);self.assertIn('请补充测量条件',value)
  self.assertIn(r'\twcomment{evidence}',value);self.assertIn(r'\twsuggest{efficiency}{原来的表述}',value)
  self.source.write_bytes(out.read_bytes());self.build()
  self.assertNotIn('建议采用的表述',tw.document_text(self.word))
 def test_deleting_comment_does_not_accept_suggestion(self):
  d=Document(self.word);cid=str(list(d.comments)[0].comment_id)
  for e in d.element.xpath('.//w:commentRangeStart | .//w:commentRangeEnd | .//w:commentReference'):
   if e.get(W+'id')==cid:e.getparent().remove(e)
  d.save(self.word);out=self.root/'deleted-comment.tex';sync_word(self.word,self.source,output=out)
  value=out.read_text(encoding='utf8');self.assertIn('原来的表述',value);self.assertNotIn('建议采用的表述',value);self.assertNotIn(r'\twsuggest{efficiency}',value)
 def test_body_edit_preserves_pending_comments(self):
  d=Document(self.word);node=next(t for t in d.element.xpath('.//w:sdt//w:t') if '确定内容' in (t.text or ''));node.text=node.text.replace('确定内容','更新后的确定内容');d.save(self.word)
  out=self.root/'prose-edit.tex';sync_word(self.word,self.source,output=out);value=out.read_text(encoding='utf8')
  self.assertIn('更新后的确定内容',value);self.assertIn(r'\twsuggest{efficiency}{原来的表述}{建议采用的表述}',value)
  self.assertIn(r'\cite{pennes}',value)
 def test_new_word_comment_is_imported_with_author(self):
  d=Document(self.word);p=next(p for p in d.element.xpath('.//w:sdtContent/w:p') if '确定内容' in ''.join(E._Element.xpath(p,'.//w:t/text()',namespaces=NS)))
  run=Paragraph(p,d._body).runs[0];d.add_comment(run,text='人工新增意见。',author='审阅者');d.save(self.word)
  out=self.root/'new-comment.tex';sync_word(self.word,self.source,output=out);value=out.read_text(encoding='utf8')
  self.assertIn(r'\twcomment[审阅者]{word-comment-',value);self.assertIn('人工新增意见',value)
 def test_non_roundtrip_generation_still_creates_native_comments(self):
  self.build(roundtrip=False);self.assertEqual(len(read_comments(Document(self.word))),2)
 def test_comment_in_table_cell_remains_pending_after_cell_edit(self):
  extra=r'''\begin{table}\caption{批注表格}\label{tab:review}
\begin{tabular}{ll}项目 & 描述\\
方法 & \twsuggest{cell-note}{现有描述}{建议描述}{需要核对。}\\
\end{tabular}\end{table}
'''
  self.source.write_text(SOURCE.replace(r'\end{document}',extra+r'\end{document}'),encoding='utf8');self.build()
  d=Document(self.word);t=next(t for t in d.element.xpath('.//w:tbl//w:t') if t.text=='方法');t.text='方法分类';d.save(self.word)
  out=self.root/'table-note.tex';sync_word(self.word,self.source,output=out);value=out.read_text(encoding='utf8')
  self.assertIn(r'\twsuggest{cell-note}{现有描述}{建议描述}',value)
  self.source.write_bytes(out.read_bytes());self.build();self.assertEqual(len(read_comments(Document(self.word))),3)
 def test_marking_resolved_never_applies_proposed_text(self):
  # A modern Word resolution flag is an acknowledgement, not text acceptance.
  from docx.opc.part import Part
  from docx.opc.packuri import PackURI
  d=Document(self.word);comment=list(d.comments)[0];p=comment._element.find('w:p',NS)
  p.set('{http://schemas.microsoft.com/office/word/2010/wordml}paraId','12345678')
  data=b'<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="12345678" w15:done="1"/></w15:commentsEx>'
  part=Part(PackURI('/word/commentsExtended.xml'),'application/vnd.ms-word.commentsExtended+xml',data,d.part.package)
  d.part.relate_to(part,'http://schemas.microsoft.com/office/2011/relationships/commentsExtended');d.save(self.word)
  before=self.source.read_bytes()
  with self.assertRaises(tw.ConversionError):sync_word(self.word,self.source,apply=True)
  self.assertEqual(self.source.read_bytes(),before);self.assertNotIn('建议采用的表述',tw.document_text(self.word))

if __name__=='__main__':unittest.main()
