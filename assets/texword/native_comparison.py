"""Native Word comparison at the existing generator/import boundary.

Word review state advances with author decisions. Publisher PDF baselines are
owned by the caller and are deliberately independent of this module.
"""
from __future__ import annotations
import base64,copy,json,subprocess,tempfile,zipfile
from pathlib import Path
from lxml import etree as E
from docx import Document
from roundtrip_model import W,M,NS,read_manifest,attach_manifest,sha,signature,doc_relationships,content_model,equivalent_math,field_parts
import re

P='{http://schemas.microsoft.com/office/2006/xmlPackage}'
C='{http://schemas.openxmlformats.org/package/2006/content-types}'
PARA_ID='{http://schemas.microsoft.com/office/word/2010/wordml}paraId'

def _style_names(doc):
    return {s.get(W+'styleId'):s.find(W+'name').get(W+'val') for s in doc.styles.element if s.find(W+'name') is not None}

def _structural_roles(doc,original):
    """Remember heading/body roles without retaining format-review balloons.

    Word may align an old body fragment with a newly inserted heading. Text
    rejection must restore that fragment's body role at the next import.
    Only exact known rejected states that reconstruct one original paragraph
    qualify; matching a common phrase or paragraph prefix is insufficient.
    """
    from section_mapping import _heading_level
    names=_style_names(doc);original_names=_style_names(original);rules=[];rels=doc_relationships(doc)
    def prose(paragraph):return ''.join(v for k,v,_ in field_parts(paragraph) if k=='text')
    original_paragraphs=[p for p in original.element.xpath('.//w:p') if not p.findall('.//m:oMath',NS)]
    rejected_flow=None
    for paragraph in doc.element.xpath('.//w:p'):
        change=paragraph.find('w:pPr/w:pPrChange',NS);identity=paragraph.get(PARA_ID)
        if not identity or paragraph.findall('.//m:oMath',NS):continue
        before=change.find(W+'pPr') if change is not None else None
        after=paragraph.find(W+'pPr')
        def style(properties):
            node=properties.find(W+'pStyle') if properties is not None else None
            return names.get(node.get(W+'val'),'') if node is not None else 'Normal'
        old_style,new_style=style(before),style(after)
        if change is None and _heading_level(doc,paragraph) is None:continue
        views=[];old_paragraph=None;new_prose=''
        for accept in (False,True):
            temporary=Document();cloned=copy.deepcopy(paragraph)
            for marker in cloned.xpath('./w:pPr/w:rPr/w:ins|./w:pPr/w:rPr/w:del'):marker.getparent().remove(marker)
            temporary.element.body.insert(0,cloned);project_document(temporary,accept=accept)
            views.append(sha(json.dumps(content_model([cloned],rels),ensure_ascii=False)))
            if not accept:old_paragraph=cloned
            else:new_prose=prose(cloned)
        if change is None:
            if new_prose.strip():continue  # Only a deletion-only fragment.
            if rejected_flow is None:
                rejected=Document()
                for element in list(rejected.element.body):rejected.element.body.remove(element)
                for element in doc.element.body:rejected.element.body.append(copy.deepcopy(element))
                project_document(rejected,accept=False)
                rejected_flow=[p for p in _flow(rejected.element.body) if p.tag==W+'p' and any(k!='text' or v.strip() for k,v,_ in field_parts(p))]
            prior_node=next((p for p in rejected_flow if p.get(PARA_ID)==identity),None)
            if prior_node is None:continue
            reconstructed=[copy.deepcopy(prior_node)]
            position=rejected_flow.index(prior_node)
            if position+1<len(rejected_flow):
                combined=copy.deepcopy(prior_node)
                for child in rejected_flow[position+1]:
                    if child.tag!=W+'pPr':combined.append(copy.deepcopy(child))
                reconstructed.append(combined)
            models=[content_model([p],rels) for p in reconstructed]
            original_rels=doc_relationships(original)
            matches=[p for p in original_paragraphs if content_model([p],original_rels) in models]
            if len(matches)!=1:continue
            prior=matches[0].find('w:pPr/w:pStyle',NS)
            old_style=original_names.get(prior.get(W+'val'),'') if prior is not None else 'Normal'
        if old_style==new_style or not any(re.fullmatch(r'heading\s+[1-9]',s,re.I) for s in (old_style,new_style)):continue
        if views[0]!=views[1]:rules.append({'para_id':identity,'old_style':old_style,'new_style':new_style,'old_content':views[0],'new_content':views[1]})
    return rules

def _project_structural_roles(doc,rules):
    names=_style_names(doc);rels=doc_relationships(doc)
    paragraphs={p.get(PARA_ID):p for p in doc.element.xpath('.//w:p') if p.get(PARA_ID)}
    for rule in rules:
        paragraph=paragraphs.get(rule['para_id'])
        if paragraph is None:continue
        properties=paragraph.find(W+'pPr')
        node=properties.find(W+'pStyle') if properties is not None else None
        current=names.get(node.get(W+'val'),'') if node is not None else 'Normal'
        if current!=rule['new_style']:continue
        actual=sha(json.dumps(content_model([paragraph],rels),ensure_ascii=False))
        if actual!=rule['old_content']:continue
        matches=[key for key,name in names.items() if name==rule['old_style']]
        if len(matches)!=1:raise ValueError('Rejected paragraph role has no unique style')
        if properties is None:properties=E.SubElement(paragraph,W+'pPr')
        if node is None:node=E.SubElement(properties,W+'pStyle')
        node.set(W+'val',matches[0])

def accept_formatting_revisions(doc):
    """Keep current formatting, removing its review history in all XML parts.

    Style-definition revisions live in styles.xml, outside the document body.
    Text, equation, row/cell insertions/deletions and real comments are retained.
    """
    removed=0
    for part in doc.part.package.parts:
        root=getattr(part,'_element',None)
        if root is None:continue
        for element in list(root.iter()):
            local=E.QName(element).localname
            if element.tag.startswith(W) and (local.endswith('PrChange') or local in ('tblGridChange','numberingChange')):
                element.getparent().remove(element);removed+=1
    if doc.settings.element.find(W+'doNotTrackFormatting') is None:E.SubElement(doc.settings.element,W+'doNotTrackFormatting')
    return removed

def flat_opc_to_docx(source,output):
    tree=E.parse(str(source),E.XMLParser(huge_tree=True,resolve_entities=False))
    types=E.Element(C+'Types',nsmap={None:C[1:-1]})
    with zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as archive:
        for part in tree.getroot():
            name=part.get(P+'name')
            if not name or '..' in Path(name).parts:raise ValueError('Invalid Flat OPC part name')
            E.SubElement(types,C+'Override',PartName=name,ContentType=part.get(P+'contentType'))
            xml=part.find(P+'xmlData');binary=part.find(P+'binaryData')
            data=(E.tostring(xml[0],xml_declaration=True,encoding='UTF-8') if xml is not None
                  else base64.b64decode(''.join(binary.itertext())))
            archive.writestr(name.lstrip('/'),data)
        archive.writestr('[Content_Types].xml',E.tostring(types,xml_declaration=True,encoding='UTF-8'))

def _flow(parent):
    for element in parent:
        if element.tag==W+'sdt':yield from _flow(element.find(W+'sdtContent'))
        else:yield element

def project_document(doc,*,accept=True):
    """Project a native comparison in a reader copy, retaining user decisions.

    Covers inline text/fields/OMML, paragraph marks and table-row changes.
    Unknown structural revision forms stop instead of silently dropping data.
    The user's file is never modified by this projection.
    """
    deleted_marks=[]
    keep='ins' if accept else 'del'
    discard='del' if accept else 'ins'
    def visit(parent):
        for element in list(parent):
            if element.getparent() is not parent:continue
            local=E.QName(element).localname
            if element.tag in (W+'ins',W+'del',W+'moveFrom',W+'moveTo'):
                side=('del' if local=='moveFrom' else 'ins' if local=='moveTo' else local)
                grand=parent.getparent()
                if parent.tag==W+'rPr' and grand is not None and grand.tag==W+'pPr':
                    if side==discard:deleted_marks.append(grand.getparent())
                    parent.remove(element);continue
                if parent.tag==W+'trPr':
                    if side==discard:
                        row=parent.getparent();row.getparent().remove(row);return
                    parent.remove(element);continue
                if parent.tag==M+'ctrlPr' and side==discard:
                    properties=parent.getparent();operator=properties.getparent()
                    operator.getparent().remove(operator);return
                at=parent.index(element)
                if side==keep:
                    for child in list(element):
                        for t in child.iter(W+'delText'):t.tag=W+'t'
                        for t in child.iter(W+'delInstrText'):t.tag=W+'instrText'
                        parent.insert(at,child);at+=1
                parent.remove(element)
                # Process children moved out of a revision, including OMML.
                continue
            if element.tag.startswith(W) and local.endswith('PrChange'):
                if accept:parent.remove(element)
                else:
                    original=element.find(parent.tag)
                    if original is None:raise ValueError('Missing original properties in native revision')
                    for child in list(parent):parent.remove(child)
                    for child in original:parent.append(copy.deepcopy(child))
                continue
            if element.tag.startswith(W) and local in ('moveFromRangeStart','moveFromRangeEnd','moveToRangeStart','moveToRangeEnd'):
                parent.remove(element);continue
            if element.tag.startswith(W) and local in ('cellIns','cellDel','cellMerge'):
                raise ValueError('Native cell-structure revision needs an explicit projection: '+local)
            visit(element)
    # Native Word can put revision containers inside math runs/properties.
    # Repeat until every moved descendant has been processed.
    while doc.element.xpath('.//w:ins|.//w:del|.//w:moveFrom|.//w:moveTo|.//w:rPrChange|.//w:pPrChange|.//w:tblPrChange|.//w:trPrChange|.//w:tcPrChange|.//w:sectPrChange'):
        visit(doc.element)
    for math in list(doc.element.iter(M+'oMath')):
        if not any((t.text or '').strip() for t in math.iter(M+'t')) and not any((t.get(M+'val') or '').strip() for t in math.iter(M+'chr')):
            parent=math.getparent();parent.remove(math)
            if parent.tag==M+'oMathPara' and not parent.findall(M+'oMath'):parent.getparent().remove(parent)
    for paragraph in reversed(deleted_marks):
        parent=paragraph.getparent()
        if parent is None:continue
        root=paragraph
        while root.getparent() is not None:root=root.getparent()
        if root is not doc.element:continue  # Its containing table row was removed.
        visible=paragraph.xpath('.//w:t|.//m:oMath|.//w:drawing|.//w:fldChar')
        if not any((n.text or '').strip() or n.tag in (M+'oMath',W+'drawing',W+'fldChar') for n in visible):
            parent.remove(paragraph);continue
        flow=list(_flow(doc.element.body))
        if paragraph not in flow:
            siblings=list(parent);index=siblings.index(paragraph)
            following=siblings[index+1] if index+1<len(siblings) else None
        else:
            index=flow.index(paragraph);following=flow[index+1] if index+1<len(flow) else None
        if following is None or following.tag!=W+'p':raise ValueError('Deleted paragraph mark meets a nonparagraph boundary')
        for child in list(following):
            if child.tag!=W+'pPr':paragraph.append(child)
        following.getparent().remove(following)
    return doc

def project_reader(doc,manifest):
    if manifest.get('native_comparison',{}).get('schema')==1:
        project_document(doc,accept=True)
        _project_structural_roles(doc,manifest['native_comparison'].get('structural_roles',[]))
        _remove_comparison_carriers(doc,manifest)
        _project_derived_bibliography(doc,manifest['native_comparison'].get('derived_bibliography'))
    return doc

def _remove_comparison_carriers(doc,manifest):
    """Drop excess empty layout paragraphs Word adds at control boundaries."""
    records={r['id']:r for r in manifest['spans']};blank=signature([E.Element(W+'p')],{})
    for control in doc.element.body:
        if control.tag!=W+'sdt':continue
        tag=control.find('w:sdtPr/w:tag',NS)
        record=records.get(tag.get(W+'val','')[8:]) if tag is not None else None
        if not record or 'units' not in record:continue
        expected=sum(value==blank for unit in record['units'] for value in unit['node_signatures'])
        content=control.find(W+'sdtContent')
        carriers=[e for e in content if e.tag==W+'p' and not e.findall('.//w:bookmarkStart',NS) and all(k=='text' and not v.strip() for k,v,_ in field_parts(e))]
        for element in carriers[:max(0,len(carriers)-expected)]:content.remove(element)

def visible_model(doc):
    """Semantic document flow, independent of invisible control boundaries."""
    elements=[]
    def flatten(parent):
        for element in parent:
            if element.tag==W+'sdt':flatten(element.find(W+'sdtContent'))
            else:elements.append(element)
    flatten(doc.element.body)
    model=content_model(elements,doc_relationships(doc))
    # DOCX-to-LaTeX transport discards paragraph-edge whitespace. It is layout,
    # and must not reappear as a new content proposal after an author decision.
    def edges(value):
        if not isinstance(value,list):return
        if len(value)==2 and value[0]=='p' and isinstance(value[1],list):
            parts=value[1]
            if parts and parts[0][0]=='text':parts[0][1]=parts[0][1].lstrip()
            if parts and parts[-1][0]=='text':parts[-1][1]=parts[-1][1].rstrip()
            while parts and parts[0][0]=='text' and not parts[0][1]:parts.pop(0)
            while parts and parts[-1][0]=='text' and not parts[-1][1]:parts.pop()
        for item in value:edges(item)
    edges(model)
    return [block for block in model if block is not None and not (block[0]=='p' and all(part[0]=='text' and not part[1].strip() for part in block[1]))]

def same_visible_content(left,right):
    def same(a,b):
        if isinstance(a,list) and isinstance(b,list):
            if len(a)!=len(b):return False
            if len(a)==3 and a[0]=='math' and b[0]=='math':return equivalent_math(a[1],b[1]) and a[2]==b[2]
            return all(same(x,y) for x,y in zip(a,b))
        return a==b
    return same(visible_model(left),visible_model(right))

def _suffix(doc):
    children=list(doc.element.body)
    positions=[i for i,e in enumerate(children) if e.tag==W+'sdt' and any(t.get(W+'val','').startswith('texword:') for t in e.findall('w:sdtPr/w:tag',NS))]
    return children[positions[-1]+1:] if positions else []

def _entry_identity(element,document):
    model=content_model([element],doc_relationships(document))
    if model and model[0][1] and model[0][1][0][0]=='text':
        model[0][1][0][1]=re.sub(r'^\[\d+\]\s*','',model[0][1][0][1])
    return sha(json.dumps(model,ensure_ascii=False))

def _bibliography_snapshot(original,proposal,manifest):
    # External .bib entries are a generated projection, never a second editable
    # bibliography database. Restrict normalization to known complete entries.
    if not re.search(r'\\bibliography\s*\{',manifest['source']):return None
    old=_suffix(original);new=_suffix(proposal)
    if not new or any(e.tag not in (W+'p',W+'sectPr',W+'bookmarkStart',W+'bookmarkEnd',W+'proofErr') for e in new):return None
    allowed={}
    for document,elements in ((original,old),(proposal,new)):
        for element in elements:
            if element.tag!=W+'p':continue
            # Only the generated entry number is presentation. It can change
            # when a new citation is inserted earlier in the manuscript.
            identity=_entry_identity(element,document)
            allowed[signature([element],doc_relationships(document))]=identity
    paragraphs=[e for e in new if e.tag==W+'p']
    used={x.get('{'+NS['r']+'}id') for e in paragraphs for x in e.iter(W+'hyperlink')}
    style_ids={x.get(W+'val') for e in paragraphs for x in e.iter() if x.tag in (W+'pStyle',W+'rStyle')}
    style_names={s.get(W+'styleId'):s.find(W+'name').get(W+'val') for s in proposal.styles.element if s.get(W+'styleId') in style_ids and s.find(W+'name') is not None}
    # Pandoc can reference Word's latent built-in Hyperlink character style
    # without materializing its definition until the first Word save.
    if 'Hyperlink' in style_ids:style_names.setdefault('Hyperlink','Hyperlink')
    return {'allowed_entry_signatures':allowed,
            'proposal_xml':[E.tostring(e,encoding='unicode') for e in paragraphs],
            'style_names':style_names,
            'external_relationships':{k:v.target_ref for k,v in proposal.part.rels.items() if v.is_external and k in used}}

def _project_derived_bibliography(doc,snapshot):
    if not snapshot:return
    suffix=_suffix(doc)
    if not suffix:return
    # Unknown prose, edited reference metadata, comments, or new containers keep
    # their existing conflict semantics. In particular, body text displaced
    # between controls by Word comparison is never classified as bibliography.
    allowed=set(snapshot['allowed_entry_signatures'].values());seen=set()
    for element in suffix:
        if element.tag in (W+'sectPr',W+'bookmarkStart',W+'bookmarkEnd',W+'proofErr'):continue
        if element.tag!=W+'p' or element.xpath('.//w:commentRangeStart|.//w:commentReference'):return
        identity=_entry_identity(element,doc)
        if identity not in allowed or identity in seen:return
        seen.add(identity)
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    styles={}
    for style in doc.styles.element:
        name=style.find(W+'name')
        if name is not None:styles.setdefault(name.get(W+'val','').casefold(),[]).append(style.get(W+'styleId'))
    replacements=[]
    for xml in snapshot['proposal_xml']:
        element=E.fromstring(xml.encode())
        for item in element.iter():
            if item.tag in (W+'pStyle',W+'rStyle'):
                name=snapshot.get('style_names',{}).get(item.get(W+'val'))
                matches=styles.get(name.casefold(),[]) if name else []
                if not matches and item.tag==W+'rStyle' and item.get(W+'val')=='Hyperlink':
                    # This latent style is undefined in both packages; retain
                    # the explicit run formatting and native hyperlink target.
                    item.getparent().remove(item);continue
                if len(matches)!=1:raise ValueError('Generated bibliography style cannot be resolved uniquely')
                item.set(W+'val',matches[0])
        for link in element.iter(W+'hyperlink'):
            key='{'+NS['r']+'}id';rid=link.get(key)
            if rid:
                url=snapshot['external_relationships'].get(rid)
                if url is None:raise ValueError('Missing generated bibliography hyperlink')
                link.set(key,doc.part.relate_to(url,RT.HYPERLINK,is_external=True))
        replacements.append(element)
    paragraphs=[e for e in suffix if e.tag==W+'p']
    if not paragraphs:return
    body=doc.element.body;at=body.index(paragraphs[0])
    for element in paragraphs:body.remove(element)
    for element in replacements:body.insert(at,element);at+=1

def compare_packages(original,revised,output,*,author='Codex',evidence=None):
    """Use Word's comparison and its read-only Flat OPC export, without UI.

    A private Word instance owns only these disposable input copies. In-memory
    comparison documents are exported directly, avoiding SaveAs on Word's
    special unnamed comparison document.
    """
    original=Path(original).resolve();revised=Path(revised).resolve();output=Path(output).resolve()
    with tempfile.TemporaryDirectory(prefix='texword_native_',dir=output.parent) as temporary:
        folder=Path(temporary);flat=folder/'comparison.xml';script=folder/'compare.ps1'
        progress=Path(evidence) if evidence else folder/'progress.txt'
        script.write_text(r'''
param([string]$Original,[string]$Revised,[string]$Flat,[string]$Progress,[string]$Author)
$ErrorActionPreference='Stop'
function Mark([string]$Value){[IO.File]::AppendAllText($Progress,([DateTime]::UtcNow.ToString('o')+' '+$Value+"`r`n"))}
$app=$null;$left=$null;$right=$null;$diff=$null
try {
 $app=New-Object -ComObject Word.Application
 $app.DisplayAlerts=0
 Mark 'private Word instance created'
 $left=$app.Documents.Open($Original,$false,$true)
 $right=$app.Documents.Open($Revised,$false,$true)
 Mark 'inputs opened read-only'
 # Track content, equations, tables and fields; style-definition churn is not
 # an author comment and must not fill the review margin with format balloons.
 $diff=$app.CompareDocuments($left,$right,2,1,$true,$true,$true,$true,$false,$false,$false,$true,$false,$false,$Author,$true)
 Mark 'native comparison returned'
 $xml=$diff.WordOpenXML
 [IO.File]::WriteAllText($Flat,$xml,[Text.Encoding]::UTF8)
 Mark 'Flat OPC exported'
} finally {
 if($app){$discard=0;$app.Quit([ref]$discard)}
 Mark 'private Word instance closed'
}
''',encoding='utf-8-sig')
        command=['C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe','-NoProfile','-ExecutionPolicy','Bypass','-File',str(script),'-Original',str(original),'-Revised',str(revised),'-Flat',str(flat),'-Progress',str(progress),'-Author',author]
        result=subprocess.run(command,capture_output=True,text=True,encoding='utf-8',errors='replace',timeout=240)
        if result.returncode:raise ValueError('Native Word comparison failed: '+result.stderr[-3000:])
        flat_opc_to_docx(flat,output)
    doc=Document(output);roles=_structural_roles(doc,Document(original));accept_formatting_revisions(doc)
    proposal=Document(revised);manifest=read_manifest(proposal)
    manifest['native_comparison']={'schema':1,'base_sha256':sha(original.read_bytes()),'proposal_sha256':sha(revised.read_bytes()),'author':author}
    if roles:manifest['native_comparison']['structural_roles']=roles
    bibliography=_bibliography_snapshot(Document(original),proposal,manifest)
    if bibliography:manifest['native_comparison']['derived_bibliography']=bibliography
    attach_manifest(doc,manifest)
    view=doc.settings.element.find(W+'revisionView')
    if view is None:view=E.SubElement(doc.settings.element,W+'revisionView')
    for name in ('markup','insDel'):view.set(W+name,'1')
    view.set(W+'formatting','0')
    if doc.settings.element.find(W+'trackRevisions') is None:E.SubElement(doc.settings.element,W+'trackRevisions')
    doc.save(output)
    return {'native_revision_nodes':len(doc.element.xpath('.//w:ins|.//w:del')),'source_sha256':manifest['source_sha256']}

def prepare_candidate(config,staged,current):
    """Keep settled decisions; reoffer only still-pending and newly added edits."""
    options=config.get('native_review')
    if not options:return None
    staged=Path(staged);current=Path(current)
    with tempfile.TemporaryDirectory(prefix='review_base_',dir=staged.parent) as temporary:
        folder=Path(temporary);base=folder/'base.docx';proposal=folder/'proposal.docx'
        existing=Document(current);manifest=read_manifest(existing)
        refresh=options.get('initial_refresh_word_sha256')
        if refresh:
            if sha(current.read_bytes())!=refresh:raise ValueError('Word changed after initial comparison repair was prepared')
            base.write_bytes(Path(options['baseline_word']).read_bytes())
        elif manifest.get('native_comparison',{}).get('schema')==1:
            # Reject only pending revisions in a copy. Already accepted or
            # rejected changes are ordinary text and remain part of this base.
            project_document(existing,accept=False)
            _project_structural_roles(existing,manifest['native_comparison'].get('structural_roles',[]))
            existing.save(base)
        else:
            base.write_bytes(Path(options['baseline_word']).read_bytes())
        proposal.write_bytes(staged.read_bytes())
        if same_visible_content(Document(base),Document(proposal)):
            # Settled decisions are the new Word baseline. Repackaging control
            # boundaries or cached field values must not reoffer those edits.
            document=Document(proposal);accept_formatting_revisions(document);metadata=read_manifest(document)
            metadata['native_comparison']={'schema':1,'base_sha256':sha(base.read_bytes()),'proposal_sha256':sha(proposal.read_bytes()),'author':options.get('author','Codex')}
            attach_manifest(document,metadata)
            if document.settings.element.find(W+'trackRevisions') is None:E.SubElement(document.settings.element,W+'trackRevisions')
            document.save(staged)
            return {'native_revision_nodes':0,'settled_decisions_preserved':True,'source_sha256':metadata['source_sha256']}
        return compare_packages(base,proposal,staged,author=options.get('author','Codex'))
