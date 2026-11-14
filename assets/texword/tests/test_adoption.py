"""Preserve imported Word blocks, while edits still flow from plain LaTeX."""
import json,subprocess,sys,tempfile,unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from docx import Document
from docx.oxml import parse_xml
from lxml import etree as E
import texword as tw
from roundtrip import sync_word
from roundtrip_model import field_parts,NS,W

class AdoptionTests(unittest.TestCase):
 def setUp(self):
  self.temp=tempfile.TemporaryDirectory();self.root=Path(self.temp.name);self.template=self.root/'original.docx';d=Document()
  d.add_paragraph('COVER')
  d.add_heading('一、第一章',level=1)
  p=d.add_paragraph('原来的正文')
  for x in E.fromstring(('<root xmlns:w="'+NS['w']+'">'+tw.run('[')+tw.field('REF BIB_one \\h',1)+tw.run(']。')+'</root>').encode()):p._p.append(x)
  d.add_paragraph('表1 示例表格');t=d.add_table(rows=2,cols=2)
  t.cell(0,0).text='项目';t.cell(0,1).text='内容';t.cell(1,0).text='阶段';t.cell(1,1).text='现有描述'
  d.add_heading('参考文献',level=1);p=d.add_paragraph()
  xml=tw.run('[')+E.tostring(tw.tag('bookmarkStart',name='BIB_one',id=42),encoding='unicode')+tw.field('SEQ AblBibRef \\* ARABIC',1)+E.tostring(tw.tag('bookmarkEnd',id=42),encoding='unicode')+tw.run('] Author. Original reference. 2020.')
  for x in E.fromstring(('<root xmlns:w="'+NS['w']+'">'+xml+'</root>').encode()):p._p.append(x)
  d.add_heading('二、第二章',level=1);d.add_paragraph('TAIL');d.save(self.template)
  self.adapter=self.root/'adapter';tw.adapt(self.template,self.adapter,body_start=1,body_end=8,samples={'body':2})
  pieces=[('header',r'\section{第一章}',[1]),('paragraph',r'原来的正文\cite{BIB_one}。',[2]),('table',r'\begin{table}\caption{示例表格}\label{tab:test}\begin{tabular}{ll}项目 & 内容\\阶段 & 现有描述\\\end{tabular}\end{table}',[3,4]),('header',r'\section*{参考文献}',[5]),('bibliography',r'\begin{thebibliography}{999}\bibitem{BIB_one}Author. Original reference. 2020.\end{thebibliography}',[6]),('header',r'\section{第二章}',[7])]
  source='\\documentclass{article}\n\\begin{document}\n\n';blocks=[]
  for kind,latex,indices in pieces:
   a=len(source);source+=latex;b=len(source);source+='\n\n';blocks.append({'kind':kind,'latex':latex,'start':a,'end':b,'original_indices':indices})
  source+='\\end{document}\n';self.source=self.root/'report.tex';self.source.write_text(source,encoding='utf8')
  tw.dump(self.adapter/'imported.json',{'source':source,'blocks':blocks,'layouts':[],'bibliography':{'BIB_one':{'number':1,'text':'Author. Original reference. 2020.'}}})
  self.code=self.root/'build.py';self.out=self.root/'six.docx'
 def tearDown(self):self.temp.cleanup()
 def build(self):
  tw.compile_code(self.adapter,self.source,self.code,roundtrip=True,overwrite=True)
  r=subprocess.run([sys.executable,str(self.code),str(self.out),'--overwrite'],capture_output=True,text=True)
  if r.returncode:raise AssertionError(r.stderr)
 def test_heading_number_detection_preserves_explicit_and_inherited_numbering(self):
  from adopt_extract import _heading_has_number,_style_table
  d=Document();p=d.add_heading('Chapter',level=1)
  self.assertFalse(_heading_has_number(E.fromstring(E.tostring(p._p)),1,_style_table(d.styles.element)))
  p.text='一、Chapter'
  self.assertTrue(_heading_has_number(E.fromstring(E.tostring(p._p)),1,_style_table(d.styles.element)))
  p.text='Chapter';p._p.get_or_add_pPr().append(parse_xml('<w:numPr xmlns:w="'+NS['w']+'"><w:numId w:val="1"/></w:numPr>'))
  self.assertTrue(_heading_has_number(E.fromstring(E.tostring(p._p)),1,_style_table(d.styles.element)))
  p._p.pPr.remove(p._p.pPr.find(W+'numPr'))
  parent=d.styles.add_style('NumberedParent',1)
  parent.element.get_or_add_pPr().append(parse_xml('<w:numPr xmlns:w="'+NS['w']+'"><w:numId w:val="1"/></w:numPr>'))
  child=d.styles.add_style('NumberedChild',1);child.base_style=parent;p.style=child
  self.assertTrue(_heading_has_number(E.fromstring(E.tostring(p._p)),1,_style_table(d.styles.element)))

 def test_fresh_standard_word_heading_adoption_and_baseline(self):
  from adoption import create_project,write_workflow
  from check_project import check_project
  original=self.root/'standard.docx';d=Document()
  d.add_paragraph('COVER');d.add_heading('Chapter',level=1)
  d.add_paragraph('Editable body');d.add_paragraph('TAIL');d.save(original)
  original_bytes=original.read_bytes();project=self.root/'fresh';out=self.root/'working.docx'
  create_project(original,project,body_start=1,body_end=3);write_workflow(project,out)
  code=project/'generate_word.py'
  tw.compile_code(project/'template_adapter',project/'report.tex',code,roundtrip=True)
  generated=subprocess.run([sys.executable,str(code),str(out)],capture_output=True,text=True)
  self.assertEqual(generated.returncode,0,generated.stderr)
  checked=check_project(project,baseline=True)
  self.assertEqual(checked['status'],'passed',checked)
  self.assertEqual(original.read_bytes(),original_bytes)
  self.assertIn(r'\section*{Chapter}',(project/'report.tex').read_text(encoding='utf8'))

 def test_initial_adoption_preserves_text_and_native_references(self):
  self.build();self.assertEqual(tw.document_text(self.template),tw.document_text(self.out))
  self.assertIn('original_blocks(',self.code.read_text(encoding='utf8'))
  result=sync_word(self.out,self.source);self.assertEqual(result['changed_blocks'],0)
 def test_source_edit_regenerates_only_changed_block(self):
  self.source.write_text(self.source.read_text(encoding='utf8').replace('原来的正文','修改后的正文'),encoding='utf8');self.build()
  text=tw.document_text(self.out);self.assertIn('修改后的正文',text);self.assertIn('二、第二章',text)
  doc=Document(self.out);fields=[v for p in doc.element.xpath('.//w:p') for k,v,f in field_parts(p) if k=='field']
  self.assertIn('REF BIB_one \\h',fields)
 def test_word_edit_keeps_citation_and_source_compiles_again(self):
  self.build();doc=Document(self.out);t=next(t for t in doc.element.xpath('.//w:sdt//w:t') if t.text=='原来的正文');t.text='Word更改的正文';doc.save(self.out)
  sync_word(self.out,self.source,apply=True);self.assertIn(r'\cite{BIB_one}',self.source.read_text(encoding='utf8'));self.build()
  self.assertIn('Word更改的正文',tw.document_text(self.out))
 def test_word_bibliography_edit_preserves_keys(self):
  self.build();doc=Document(self.out);t=next(t for t in doc.element.xpath('.//w:sdt//w:t') if 'Original reference' in (t.text or ''));t.text=t.text.replace('Original reference','Updated reference');doc.save(self.out)
  sync_word(self.out,self.source,apply=True);value=self.source.read_text(encoding='utf8');self.assertIn(r'\bibitem{BIB_one}',value);self.build()
  self.assertIn('Updated reference',tw.document_text(self.out))

class TableImportTests(unittest.TestCase):
 def test_generated_empty_rule_suffixes_preserve_literal_parentheses(self):
  from adoption import simplify_latex
  value='\\begin{longtable}{ll}\n\\toprule()\nA & B\\\\\n\\midrule()\n() & keep \\toprule() here\\\\\n\\bottomrule()\n\\end{longtable}'
  cleaned=simplify_latex(value,columns=2)
  for rule in ('toprule','midrule','bottomrule'):
   self.assertNotIn('\\'+rule+'()\n',cleaned)
  self.assertIn(r'() & keep \toprule() here',cleaned)
  self.assertEqual(simplify_latex(value),value)
  self.assertIn(r'\midrule[0.5pt]',simplify_latex(r'\midrule[0.5pt]',columns=2))
  self.assertIn(r'\bottomrule(note)',simplify_latex(r'\bottomrule(note)',columns=2))

 def test_word_table_roundtrip_keeps_rows_and_real_empty_parentheses(self):
  from roundtrip import block_to_latex
  doc=Document();table=doc.add_table(rows=2,cols=2)
  table.cell(0,0).text='Item';table.cell(0,1).text='Description'
  table.cell(1,0).text='()';table.cell(1,1).text='Literal parentheses ()'
  record={'id':'table','kind':'table','latex':r'\begin{longtable}{ll}Item & Description\\() & Literal parentheses ()\\\end{longtable}'}
  imported=block_to_latex(doc,[table._tbl],record,{'references':{}})
  self.assertIn('Literal parentheses ()',imported)
  self.assertNotRegex(imported,r'\\(?:toprule|midrule|bottomrule)\(\)')
  with tempfile.TemporaryDirectory() as tmp:
   rebuilt=Path(tmp)/'rebuilt.docx'
   tw.pandoc(['-f','latex','-o',rebuilt],input=imported)
   actual=Document(rebuilt).tables[0]
   self.assertEqual(len(actual.rows),2)
   self.assertEqual(len(actual.columns),2)
   self.assertEqual(actual.cell(1,0).text,'()')
   self.assertEqual(actual.cell(1,1).text,'Literal parentheses ()')

if __name__=='__main__':unittest.main()
