import copy,hashlib,json,sys,tempfile,unittest,zipfile,subprocess,os
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import texword as tw
from docx import Document
from docx.shared import Pt,Cm

SOURCE=r'''\documentclass{article}
\usepackage{amsmath,booktabs}
\begin{document}
\section{转换验证}\label{sec:test}
正文含行内公式 $Q_e=\sigma|\nabla\phi|^2$，文献\cite{pennes}。
\subsection{水分数模型}
\begin{equation}\label{eq:heat}
\rho c\frac{\partial T}{\partial t}=\nabla\cdot(k\nabla T)+Q_e
\end{equation}
公式\eqref{eq:heat}与表\ref{tab:test}，再次引用\cite{pennes}。
\begin{table}[htbp]\centering
\caption{测试表格}\label{tab:test}
\begin{tabular}{ll}\toprule
阶段 & 状态\\\midrule
冷冻 & 冰与液态水\\
射频 & 液态水与水蒸气\\\bottomrule
\end{tabular}\end{table}
\end{document}
'''
BIB='''@article{pennes, author={Pennes, Harry H.}, title={Analysis of tissue and arterial blood temperatures in the resting human forearm}, journal={Journal of Applied Physiology}, year={1948}, volume={1}, number={2}, pages={93--122}, doi={10.1152/jappl.1948.1.2.93}}'''

class ConversionTests(unittest.TestCase):
 def setUp(self):
  self.temp=tempfile.TemporaryDirectory();self.root=Path(self.temp.name)
  self.template=self.root/'template.docx';d=Document()
  d.sections[0].page_width=Cm(21);d.sections[0].page_height=Cm(29.7)
  d.sections[0].header.paragraphs[0].text='TEST HEADER'
  d.add_paragraph('COVER UNCHANGED')
  d.add_heading('Body starts',level=1)
  d.add_heading('Subheading example',level=2)
  p=d.add_paragraph('Body example');p.style=d.styles['Normal'];p.paragraph_format.first_line_indent=Cm(.74)
  p.runs[0].font.name='Calibri';p.runs[0].font.size=Pt(12)
  d.add_paragraph('TAIL UNCHANGED');d.save(self.template)
  self.source=self.root/'draft.tex';self.source.write_text(SOURCE,encoding='utf8')
  self.bib=self.root/'refs.bib';self.bib.write_text(BIB,encoding='utf8')
  self.before=self.template.read_bytes()
  self.adapter=self.root/'adapter'
 def tearDown(self):self.temp.cleanup()
 def adapt(self):
  return tw.adapt(self.template,self.adapter,body_start='Body starts',body_end='TAIL UNCHANGED',samples={'body':3})
 def test_profile_and_generated_driver(self):
  p=self.adapt()
  self.assertEqual(p['roles']['body']['sample_index'],3)
  self.assertTrue((self.adapter/'template_adapter.py').is_file())
  self.assertTrue((self.adapter/'reference.docx').is_file())
  self.assertEqual(self.template.read_bytes(),self.before)
 def test_native_math_fields_template_and_rebuild(self):
  self.adapt();out=self.root/'trial.docx'
  code=self.root/'build_word.py'
  compiled=tw.compile_code(self.adapter,self.source,code,bibliography=self.bib)
  self.assertFalse(out.exists())
  self.assertIn('document.table(',code.read_text(encoding='utf8'))
  self.assertIn('document.equation(',code.read_text(encoding='utf8'))
  env=dict(os.environ,PANDOC='INTENTIONALLY_UNAVAILABLE')
  self.source.rename(self.source.with_suffix('.hidden'));self.bib.rename(self.bib.with_suffix('.hidden'))
  subprocess.run([sys.executable,str(code),str(out)],check=True,env=env,stdout=subprocess.PIPE,stderr=subprocess.PIPE)
  report=tw.load(out.with_suffix('.validation.json'))
  self.source.with_suffix('.hidden').rename(self.source);self.bib.with_suffix('.hidden').rename(self.bib)
  self.assertGreaterEqual(report['math_objects'],2)
  self.assertGreaterEqual(report['cross_reference_fields'],4)
  self.assertEqual(report['missing_bookmarks'],[])
  self.assertTrue(report['prefix_preserved']);self.assertTrue(report['suffix_preserved'])
  with zipfile.ZipFile(out) as z:
   self.assertIn('TEST HEADER',z.read('word/header1.xml').decode())
   self.assertEqual(z.read('word/header1.xml'),zipfile.ZipFile(self.template).read('word/header1.xml'))
  self.assertEqual(self.template.read_bytes(),self.before)
  self.source.write_text(SOURCE.replace('正文含','第二次修改含'),encoding='utf8')
  tw.compile_code(self.adapter,self.source,code,bibliography=self.bib,overwrite=True)
  subprocess.run([sys.executable,str(code),str(out),'--overwrite'],check=True,stdout=subprocess.PIPE,stderr=subprocess.PIPE)
  self.assertIn('第二次修改含',tw.document_text(out))
  self.assertNotIn('正文含',tw.document_text(out))
  code.write_text(code.read_text(encoding='utf8').replace('第二次修改含','直接修改生成代码含'),encoding='utf8')
  subprocess.run([sys.executable,str(code),str(out),'--overwrite'],check=True,env=env,stdout=subprocess.PIPE,stderr=subprocess.PIPE)
  self.assertIn('直接修改生成代码含',tw.document_text(out))
 def test_unknown_reference_rejected(self):
  self.adapt();self.source.write_text(SOURCE.replace('eq:heat}', 'eq:missing}',1),encoding='utf8')
  with self.assertRaisesRegex(tw.ConversionError,'Unresolved reference'):tw.compile_code(self.adapter,self.source,self.root/'bad.py',bibliography=self.bib)
 def test_unknown_citation_rejected(self):
  self.adapt();self.source.write_text(SOURCE.replace('cite{pennes}','cite{missing}'),encoding='utf8')
  with self.assertRaises(tw.ConversionError):tw.compile_code(self.adapter,self.source,self.root/'bad.py',bibliography=self.bib)
 def test_citation_group_limit(self):
  self.adapt();self.source.write_text(SOURCE.replace('cite{pennes}','cite{a,b,c,d}'),encoding='utf8')
  with self.assertRaisesRegex(tw.ConversionError,'three'):tw.compile_code(self.adapter,self.source,self.root/'bad.py',bibliography=self.bib)
 def test_unsupported_raw_content_rejected(self):
  self.adapt();self.source.write_text(SOURCE.replace('正文含',r'\begin{tikzpicture}\draw (0,0)--(1,1);\end{tikzpicture} 正文含'),encoding='utf8')
  with self.assertRaisesRegex(tw.ConversionError,'Unsupported'):tw.compile_code(self.adapter,self.source,self.root/'bad.py',bibliography=self.bib)
 def test_template_change_detected(self):
  self.adapt();p=self.adapter/'template.docx';p.write_bytes(p.read_bytes()+b'changed')
  with self.assertRaisesRegex(tw.ConversionError,'Template changed'):tw.compile_code(self.adapter,self.source,self.root/'bad.py',bibliography=self.bib)
 def test_no_original_overwrite(self):
  self.adapt()
  with self.assertRaises(tw.ConversionError):tw.compile_code(self.adapter,self.source,self.template,bibliography=self.bib,overwrite=True)
 def test_generated_runtime_protects_source(self):
  self.adapt();code=self.root/'build_word.py'
  tw.compile_code(self.adapter,self.source,code,bibliography=self.bib)
  original=self.source.read_bytes()
  result=subprocess.run([sys.executable,str(code),str(self.source),'--overwrite'],stdout=subprocess.PIPE,stderr=subprocess.PIPE)
  self.assertNotEqual(result.returncode,0)
  self.assertIn(b'Output must use .docx',result.stderr)
  self.assertEqual(original,self.source.read_bytes())
 def test_adapter_links_shared_runtime_and_compiled_bundle_is_standalone(self):
  self.adapt();code=self.root/'build_word.py';out=self.root/'portable.docx'
  link=(self.adapter/'word_runtime.py').read_text(encoding='utf8')
  self.assertNotIn('class TemplateDocument',link)
  probe='import sys;sys.path.insert(0,sys.argv[1]);import template_adapter,word_runtime;template_adapter.Report();print(word_runtime.__file__)'
  result=subprocess.run([sys.executable,'-c',probe,str(self.adapter)],capture_output=True,text=True,check=True)
  self.assertEqual(Path(result.stdout.strip()).resolve(),(tw.ROOT/'word_runtime.py').resolve())
  tw.compile_code(self.adapter,self.source,code,bibliography=self.bib)
  assets=self.root/'build_word_assets'
  self.assertEqual((assets/'word_runtime.py').read_bytes(),(tw.ROOT/'word_runtime.py').read_bytes())
  import shutil
  shutil.rmtree(self.adapter)
  subprocess.run([sys.executable,str(code),str(out)],check=True,capture_output=True)
  self.assertIn('正文含',tw.document_text(out))
 def test_template_specific_code_calibration_is_used(self):
  self.adapt();code=self.root/'build_word.py';out=self.root/'calibrated.docx'
  adapter_code=self.adapter/'template_adapter.py'
  with adapter_code.open('a',encoding='utf8') as f:
   f.write('\nPROFILE["roles"]["body"]["rPr"] = \'<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:sz w:val="38"/></w:rPr>\'\n')
  tw.compile_code(self.adapter,self.source,code,bibliography=self.bib)
  subprocess.run([sys.executable,str(code),str(out)],check=True,stdout=subprocess.PIPE,stderr=subprocess.PIPE)
  from lxml import etree as E
  with zipfile.ZipFile(out) as z:styles=E.fromstring(z.read('word/styles.xml'))
  self.assertEqual(styles.xpath('string(w:style[@w:styleId="BodyText"]/w:rPr/w:sz/@w:val)',namespaces=tw.NS),'38')

if __name__=='__main__':unittest.main(verbosity=2)
