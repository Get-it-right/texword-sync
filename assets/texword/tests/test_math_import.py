"""Regression for Word's empty run after an invisible closing delimiter."""
import copy,sys,tempfile,unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from docx import Document
from docx.oxml import OxmlElement
from roundtrip import block_to_latex
from roundtrip_model import NS,math_model,equivalent_math
from lxml import etree as E
from texword import pandoc


class MathImportTests(unittest.TestCase):
 def styled_document(self,root,paragraphs,sources):
  path=Path(root)/'styled.docx'
  pandoc(['-f','latex','-o',path],input='\n\n'.join(paragraphs))
  doc=Document(path);records=[]
  for paragraph,latex,maths in zip(doc.paragraphs,paragraphs,sources):
   records.append({'id':'styled','kind':'paragraph','latex':latex,
    'math_sources':[{'latex':m,'display':False} for m in maths],
    'math_baselines':[math_model(m) for m in paragraph._p.findall('.//m:oMath',NS)]})
  return doc,records

 def assert_upright_and_no_placeholder(self,root,latex):
  path=Path(root)/'rebuilt.docx';pandoc(['-f','latex','-o',path],input=latex)
  doc=Document(path);m='{'+NS['m']+'}'
  for run in doc.element.xpath('.//m:r'):
   payload=''.join(t.text or '' for t in run.findall('m:t',NS))
   if any(c.isalpha() or c=='°' for c in payload):
    style=run.find('m:rPr/m:sty',NS)
    self.assertIsNotNone(style,payload);self.assertEqual(style.get(m+'val'),'p',payload)
  for slot in doc.element.xpath('.//m:sSup/m:e'):
   self.assertNotEqual(''.join(slot.itertext()),'\u200b')

 def test_adjacent_edited_paragraphs_preserve_unit_sources(self):
  from section_mapping import _convert,section_to_latex
  from roundtrip_model import signature,doc_relationships
  with tempfile.TemporaryDirectory() as tmp:
   sources=[[r'\mathrm{Pa}',r'\mathrm{K}',r'\mathrm{kg}\cdot\mathrm{m}^{-3}',r'\mathrm{°C}'],[r'\mathrm{s}',r'\mathrm{J}']]
   paragraphs=['Alpha '+' '.join('$'+m+'$' for m in sources[0]),'Beta '+' '.join('$'+m+'$' for m in sources[1])]
   doc,old=self.styled_document(tmp,paragraphs,sources)
   for unit,p in zip(old,doc.paragraphs):
    unit['node_signatures']=[signature([p._p],doc_relationships(doc))];unit['review_baseline']=[]
   for p in doc.paragraphs:p.runs[0].text='Edited'+p.runs[0].text
   imported=_convert(doc,[p._p for p in doc.paragraphs],old,{'references':{}},{})
   for group in sources:
    for formula in group:self.assertIn(formula,imported)
   self.assert_upright_and_no_placeholder(tmp,imported)
   record={'units':old,'math_baselines':[m for unit in old for m in unit['math_baselines']]}
   imported=section_to_latex(doc,[p._p for p in doc.paragraphs],record,{'references':{}},{},{})
   for group in sources:
    for formula in group:self.assertIn(formula,imported)

 def test_formula_insertion_or_deletion_keeps_later_roman_units(self):
  for operation in ('insert','delete'):
   with self.subTest(operation=operation),tempfile.TemporaryDirectory() as tmp:
    formulas=[r'\mathrm{Pa}',r'\mathrm{K}',r'\mathrm{°C}']
    doc,records=self.styled_document(tmp,['Alpha '+' '.join('$'+m+'$' for m in formulas)],[formulas])
    paragraph=doc.paragraphs[0]._p;first=paragraph.find('m:oMath',NS)
    if operation=='delete':paragraph.remove(first)
    else:
     new=copy.deepcopy(first);new.find('.//m:t',NS).text='x';paragraph.insert(1,new)
    imported=block_to_latex(doc,[paragraph],records[0],{'references':{}})
    self.assertIn(r'\mathrm{K}',imported);self.assertIn(r'\mathrm{°C}',imported)
    self.assert_upright_and_no_placeholder(tmp,imported)

 def test_real_numeric_edit_retains_explicit_roman_style(self):
  with tempfile.TemporaryDirectory() as tmp:
   doc,records=self.styled_document(tmp,[r'Unit $2\mathrm{K}$; temp $\mathrm{°C}$.'],[[r'2\mathrm{K}',r'\mathrm{°C}']])
   next(t for t in doc.element.xpath('.//m:t') if t.text=='2').text='3'
   imported=block_to_latex(doc,[doc.paragraphs[0]._p],records[0],{'references':{}})
   self.assertIn('3',imported);self.assertNotIn('2',imported)
   self.assertIn(r'\mathrm{K}',imported);self.assert_upright_and_no_placeholder(tmp,imported)

 def test_new_formula_retains_roman_units_without_baseline(self):
  with tempfile.TemporaryDirectory() as tmp:
   doc,records=self.styled_document(tmp,[r'Units $\mathrm{Pa}$ $\mathrm{K}$ $\mathrm{°C}$.'],[[r'\mathrm{Pa}',r'\mathrm{K}',r'\mathrm{°C}']])
   records[0]['math_sources']=[];records[0]['math_baselines']=[]
   imported=block_to_latex(doc,[doc.paragraphs[0]._p],records[0],{'references':{}})
   self.assertIn(r'\mathrm{K}',imported);self.assertNotIn(r'\circ',imported)
   self.assert_upright_and_no_placeholder(tmp,imported)

 def test_roman_and_italic_same_letter_remain_distinct_after_insertion(self):
  with tempfile.TemporaryDirectory() as tmp:
   doc,records=self.styled_document(tmp,[r'Variable $K$ unit $\mathrm{K}$.'],[['K',r'\mathrm{K}']])
   p=doc.paragraphs[0]._p;new=copy.deepcopy(p.find('m:oMath',NS));new.find('.//m:t',NS).text='x';p.insert(1,new)
   imported=block_to_latex(doc,[p],records[0],{'references':{}})
   self.assertEqual(imported.count(r'\mathrm{K}'),1);self.assertIn('$K$',imported)
   # A genuine font edit is imported rather than matched to the old roman K.
   run=p.findall('m:oMath',NS)[-1].find('.//m:r',NS);run.find('m:rPr',NS).remove(run.find('m:rPr/m:sty',NS))
   imported=block_to_latex(doc,[p],records[0],{'references':{}})
   self.assertNotIn(r'\mathrm{K}',imported)

 def test_inline_formula_does_not_reuse_identical_display_source(self):
  with tempfile.TemporaryDirectory() as tmp:
   doc,records=self.styled_document(tmp,[r'Variable $K$.'],[['K']])
   record=records[0];record['math_sources'].insert(0,{'latex':'K','display':True})
   record['math_baselines'].insert(0,record['math_baselines'][0])
   imported=block_to_latex(doc,[doc.paragraphs[0]._p],record,{'references':{}})
   self.assertNotIn(r'\[',imported);self.assertIn('$K$',imported)

 def test_parenthesized_units_upright_without_changing_variables(self):
  from roundtrip import _normalize_imported_math
  unit={'t':'Math','c':[{'t':'InlineMath'},r'kg \cdot m^{- 3}']}
  variable={'t':'Math','c':[{'t':'InlineMath'},'K']}
  values=[{'t':'Str','c':'密度（'},unit,{'t':'Str','c':'），渗透率'},variable]
  _normalize_imported_math(values)
  # Parentheses alone do not identify a unit: s or K may be a variable.
  self.assertEqual(unit['c'][1],r'kg \cdot m^{- 3}')
  self.assertEqual(variable['c'][1],'K')
 def test_hidden_limit_cleanup_preserves_real_limits_and_spacing(self):
  from word_runtime import normalize_hidden_math_limits
  m='{'+NS['m']+'}'
  def operator(hidden,value):
   return E.fromstring(('<m:nary xmlns:m="'+NS['m']+'"><m:naryPr><m:supHide m:val="'+hidden+'"/></m:naryPr>'
    '<m:sub><m:r><m:t>j</m:t></m:r></m:sub><m:sup><m:r><m:t>'+value+'</m:t></m:r></m:sup>'
    '<m:e><m:r><m:t>c\u2001c</m:t></m:r></m:e></m:nary>').encode())
  hidden=operator('1','\u200b');visible=operator('0','\u200b');real=operator('1','2')
  root=E.Element(m+'oMath');root.extend([hidden,visible,real])
  self.assertEqual(normalize_hidden_math_limits(root),1)
  self.assertEqual(len(hidden.find('m:sup',NS)),0)
  self.assertEqual(''.join(visible.find('m:sup',NS).itertext()),'\u200b')
  self.assertEqual(''.join(real.find('m:sup',NS).itertext()),'2')
  self.assertEqual(''.join(hidden.find('m:e',NS).itertext()),'c\u2001c')
  self.assertEqual(normalize_hidden_math_limits(root),0)

 def test_unused_empty_delimiter_separator_preserves_real_edits(self):
  def model(properties,arguments):
   body='<m:d><m:dPr>'+properties+'</m:dPr>'
   body+=''.join('<m:e><m:r><m:t>'+value+'</m:t></m:r></m:e>' for value in arguments)
   return math_model(E.fromstring(('<m:oMath xmlns:m="'+NS['m']+'">'+body+'</m:d></m:oMath>').encode()))
  empty='<m:sepChr m:val=""/>'
  self.assertTrue(equivalent_math(model(empty,['x']),model('',['x'])))
  self.assertFalse(equivalent_math(model('<m:sepChr m:val="|"/>',['x']),model('',['x'])))
  self.assertFalse(equivalent_math(model(empty,['x']),model('',['y'])))
  self.assertFalse(equivalent_math(model(empty,['x']),model('',['x','y'])))
  self.assertFalse(equivalent_math(model(empty,['x','y']),model('',['x','y'])))
  self.assertFalse(equivalent_math(model('<m:sepChr m:val=","/>',['x','y']),model('',['x','y'])))
  self.assertFalse(equivalent_math(model(empty,['x']),model('<m:endChr m:val="]"/>',['x'])))

 def test_word_em_spaces_remain_math_spacing_and_recompile(self):
  with tempfile.TemporaryDirectory() as tmp:
   word=Path(tmp)/'math.docx'
   pandoc(['-f','latex','-o',word],input=r'\[x+1\]')
   doc=Document(word);root=doc.element.xpath('.//m:oMath')[0]
   run=OxmlElement('m:r');text=OxmlElement('m:t');text.text='\u2005\u2001\u2001\u200b';run.append(text);root.append(run)
   following=OxmlElement('m:r');label=OxmlElement('m:t');label.text='D';following.append(label);root.append(following)
   record={'id':'equation','kind':'equation','latex':r'\[x+1\]'}
   imported=block_to_latex(doc,[root.xpath('ancestor::w:p',namespaces=NS)[0]],record,{'references':{}})
   self.assertNotIn(r'\text{',imported)
   self.assertIn(r'\quad\quad',imported)
   rebuilt=Path(tmp)/'rebuilt.docx'
   pandoc(['-f','latex','-o',rebuilt],input=imported)
   self.assertEqual(len(Document(rebuilt).element.xpath('.//m:oMath')),1)
   self.assertTrue(any(t.text=='D' for t in Document(rebuilt).element.xpath('.//m:t')))
   self.assertNotIn('\u200b',''.join(t.text or '' for t in Document(rebuilt).element.xpath('.//m:t')))

 def test_word_defaults_preserve_integral_and_matrix_but_detect_real_edits(self):
  def model(body):
   return math_model(E.fromstring(('<m:oMath xmlns:m="'+NS['m']+'">'+body+'</m:oMath>').encode()))
  integral='<m:nary><m:naryPr>{}</m:naryPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:nary>'
  explicit=model(integral.format('<m:chr m:val="∫"/><m:subHide m:val="0"/><m:supHide m:val="0"/>'))
  self.assertTrue(equivalent_math(explicit,model(integral.format(''))))
  self.assertFalse(equivalent_math(explicit,model(integral.format('<m:chr m:val="∑"/>'))))
  column='<m:mc><m:mcPr><m:count m:val="{}"/><m:mcJc m:val="{}"/></m:mcPr></m:mc>'
  def matrix(columns,base=''):
   return model('<m:m><m:mPr>'+base+'<m:mcs>'+columns+'</m:mcs></m:mPr><m:mr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:mr></m:m>')
  separate=matrix(column.format(1,'center')*2,'<m:baseJc m:val="center"/>')
  self.assertTrue(equivalent_math(separate,matrix(column.format(2,'center'))))
  self.assertFalse(equivalent_math(separate,matrix(column.format(1,'center'))))
  self.assertFalse(equivalent_math(separate,matrix(column.format(2,'left'))))
  self.assertFalse(equivalent_math(separate,matrix(column.format(2,'center'),'<m:baseJc m:val="top"/>')))

 def test_terminal_invisible_delimiter_space_and_real_edits(self):
  formula=r'H(T)=\left\{\begin{matrix}0&T<0\\1&T\geq0\end{matrix}\right.'
  latex='\\begin{equation*}\n'+formula+'\n\\end{equation*}'
  with tempfile.TemporaryDirectory() as tmp:
   word=Path(tmp)/'math.docx'
   pandoc(['-f','latex','-o',word],input=latex)
   doc=Document(word);root=doc.element.xpath('.//m:oMath')[0]
   run=OxmlElement('m:r');text=OxmlElement('m:t');text.text=' ';run.append(text);root.append(run)
   record={'id':'equation','kind':'equation','latex':latex}
   def convert():
    return block_to_latex(doc,[root.xpath('ancestor::w:p',namespaces=NS)[0]],record,{'references':{}})
   imported=convert()
   self.assertNotIn('\\right.\\ ',imported)
   self.assertIn('\\right.\n',imported)
   # An actual coefficient edit must still be imported.
   next(t for t in root.findall('.//m:t',NS) if t.text=='1').text='2'
   self.assertIn('2',convert())
   record['latex']='\\['+formula+'\\]'
   self.assertTrue(convert().startswith('\\['))
   self.assertTrue(convert().endswith('\\]'))
   # A space before a visible closing delimiter is outside the repair scope.
   case_properties=next(p for p in root.findall('.//m:dPr',NS)
    if p.find('m:begChr',NS) is not None and p.find('m:begChr',NS).get('{'+NS['m']+'}val')=='{')
   case_properties.find('m:endChr',NS).set('{'+NS['m']+'}val',']')
   self.assertIn('\\right\\rbrack\\ ',convert())


if __name__=='__main__':unittest.main()
