"""Native comparison boundaries: content decisions versus formatting history."""
import sys,unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from docx import Document
from lxml import etree as E
from native_comparison import accept_formatting_revisions,_remove_comparison_carriers,_structural_roles,_project_structural_roles,project_document,PARA_ID
from roundtrip_model import W,signature,attach_manifest,read_manifest,RT_NS,sha

class NativeComparisonTests(unittest.TestCase):
 def test_deleted_row_does_not_merge_its_detached_cell_paragraphs(self):
  from roundtrip_model import M
  doc=Document();table=doc.add_table(rows=2,cols=1);table.cell(0,0).text='keep'
  row=table.rows[1]._tr;properties=E.Element(W+'trPr');E.SubElement(properties,W+'del');row.insert(0,properties)
  paragraph=table.cell(1,0).paragraphs[0]._p
  E.SubElement(E.SubElement(paragraph.get_or_add_pPr(),W+'rPr'),W+'del')
  math=E.SubElement(E.SubElement(paragraph,M+'oMathPara'),M+'oMath')
  deleted=E.SubElement(math,W+'del');run=E.SubElement(deleted,M+'r');E.SubElement(run,M+'t').text='1'
  project_document(doc,accept=True)
  self.assertEqual(len(table.rows),1)
  self.assertEqual(table.cell(0,0).text,'keep')
 def test_rejected_body_fragment_retains_body_role_without_format_revision(self):
  original=Document();original.add_paragraph('The simulations identified a condition-dependent near-saturation region and its interpretation.')
  comparison=Document();paragraph=comparison.add_paragraph(style='Heading 2');paragraph._p.set(PARA_ID,'00000001')
  deleted=E.SubElement(paragraph._p,W+'del');run=E.SubElement(deleted,W+'r')
  E.SubElement(run,W+'delText').text='The simulations identified a condition-dependent near-saturation '
  comparison.add_paragraph('region and its interpretation.')
  rules=_structural_roles(comparison,original)
  self.assertEqual(len(rules),1)
  project_document(comparison,accept=False);_project_structural_roles(comparison,rules)
  from section_mapping import _heading_level
  self.assertIsNone(_heading_level(comparison,paragraph._p))
  self.assertFalse(comparison.element.xpath('.//w:pPrChange'))
 def duplicate_map(self,doc,value,name='duplicate'):
  from docx.opc.part import Part
  from docx.opc.packuri import PackURI
  from docx.opc.constants import RELATIONSHIP_TYPE as RT
  import json
  root=E.Element('{'+RT_NS+'}roundtrip');root.text=json.dumps(value)
  part=Part(PackURI('/customXml/'+name+'.xml'),'application/xml',E.tostring(root),doc.part.package)
  doc.part.relate_to(part,RT.CUSTOM_XML)

 def test_legacy_duplicate_maps_reconcile_then_collapse_owned_parts(self):
  from docx.opc.part import Part
  from docx.opc.packuri import PackURI
  from docx.opc.constants import RELATIONSHIP_TYPE as RT
  doc=Document();core={'source':'original','source_sha256':sha('original'),'spans':[]}
  attach_manifest(doc,core)
  self.duplicate_map(doc,dict(core,native_comparison={'schema':1,'author':'Codex','base_sha256':'a','proposal_sha256':'b'}))
  unrelated=Part(PackURI('/customXml/unrelated.xml'),'application/xml',b'<other>keep</other>',doc.part.package)
  doc.part.relate_to(unrelated,RT.CUSTOM_XML)
  merged=read_manifest(doc);self.assertEqual(merged['native_comparison']['schema'],1)
  attach_manifest(doc,merged)
  roots=[E.fromstring(rel.target_part.blob).tag for rel in doc.part.rels.values() if not rel.is_external and rel.reltype==RT.CUSTOM_XML]
  self.assertEqual(roots.count('{'+RT_NS+'}roundtrip'),1)
  self.assertIn('other',roots)

 def test_duplicate_maps_with_incompatible_content_stop(self):
  doc=Document();attach_manifest(doc,{'source':'a','spans':[]})
  self.duplicate_map(doc,{'source':'b','spans':[]})
  with self.assertRaisesRegex(ValueError,'incompatible content baselines'):read_manifest(doc)

 def test_format_balloon_cleanup_keeps_content_revisions_and_real_comments(self):
  doc=Document();paragraph=doc.add_paragraph('Measured temperature needs confirmation.')
  doc.add_comment(paragraph.runs,text='Confirm the sensor position.',author='Codex')
  style=doc.styles['Heading 1'].element
  properties=E.SubElement(style,W+'pPr')
  E.SubElement(properties,W+'pPrChange').set(W+'author','Codex')
  E.SubElement(paragraph._p.get_or_add_pPr(),W+'pPrChange')
  insertion=E.SubElement(paragraph._p,W+'ins');run=E.SubElement(insertion,W+'r');E.SubElement(run,W+'t').text=' New evidence.'
  self.assertEqual(accept_formatting_revisions(doc),2)
  self.assertEqual(len(doc.element.xpath('.//w:ins')),1)
  self.assertEqual(len(doc.comments),1)
  self.assertIn('Confirm the sensor position.',next(iter(doc.comments)).text)
  self.assertFalse(style.findall('.//'+W+'pPrChange'))
  self.assertIsNotNone(doc.settings.element.find(W+'doNotTrackFormatting'))

 def test_carrier_cleanup_keeps_anchored_and_original_blank_paragraphs(self):
  doc=Document();control=E.Element(W+'sdt');properties=E.SubElement(control,W+'sdtPr')
  E.SubElement(properties,W+'tag').set(W+'val','texword:section')
  content=E.SubElement(control,W+'sdtContent')
  extra=E.SubElement(content,W+'p');original=E.SubElement(content,W+'p')
  anchor=E.SubElement(content,W+'p');E.SubElement(anchor,W+'bookmarkStart').set(W+'name','real_target')
  doc.element.body.insert(0,control)
  blank=signature([E.Element(W+'p')],{})
  _remove_comparison_carriers(doc,{'spans':[{'id':'section','units':[{'node_signatures':[blank]}]}]})
  self.assertIsNone(extra.getparent())
  self.assertIsNotNone(original.getparent())
  self.assertIsNotNone(anchor.getparent())

if __name__=='__main__':unittest.main()
