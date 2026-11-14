"""Word-native accept/reject decisions, separately from ordinary comments."""
import os
import subprocess
import sys
import unittest
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import texword as tw
import test_texword as fixtures
from docx import Document
from docx.oxml import OxmlElement
from roundtrip import sync_word
from roundtrip_model import W, M, math_model, equivalent_math, signature, read_manifest
from tracked_reviews import project_revisions

SOURCE = r'''\documentclass{article}
\begin{document}
\section{修订验证}
已经确定的内容。\twrevision{change-one}{原来的表述}{建议采用的表述}{核对理由。}
保留\twcomment{note}{证据说明}{普通批注没有批准含义。}，文献\cite{pennes}。
\end{document}
'''

SOURCE_CITATION_REVISIONS = r'''\documentclass{article}
\begin{document}
\section{引用修订验证}
\twrevision{citation-change}{原引用\cite{pennes}}{建议引用\cite{future}}{更新到新证据。}
\twrevision{delete-citation}{待删除引用\cite{pennes}}{}{删除理由。}
在此处\twrevision{insert-citation}{}{新增证据\cite{future}}{补充理由。}之后继续。
\end{document}
'''


class TrackedReviewTests(unittest.TestCase):
    def setUp(self):
        self.fixture = fixtures.ConversionTests(); self.fixture.setUp()
        self.root = self.fixture.root; self.fixture.adapt()
        self.source = self.fixture.source; self.source.write_text(SOURCE, encoding='utf8')
        self.code = self.root/'tracked_build.py'; self.word = self.root/'tracked.docx'
        self.build()

    def tearDown(self):
        self.fixture.tearDown()

    def build(self):
        tw.compile_code(self.fixture.adapter, self.source, self.code, bibliography=self.fixture.bib, roundtrip=True, overwrite=True)
        subprocess.run([sys.executable, str(self.code), str(self.word), '--overwrite'], check=True, capture_output=True)

    def decide(self, accept):
        doc = Document(self.word)
        for node in list(doc.element.xpath('.//w:ins | .//w:del')):
            parent = node.getparent(); at = parent.index(node)
            if (node.tag == W+'ins') == accept:
                for child in list(node):
                    for text in child.iter(W+'delText'):text.tag = W+'t'
                    for instruction in child.iter(W+'delInstrText'):instruction.tag = W+'instrText'
                    parent.insert(at, child); at += 1
            parent.remove(node)
        doc.save(self.word)

    def test_pending_is_real_native_revision_and_source_unchanged(self):
        doc = Document(self.word)
        self.assertEqual(len(doc.element.xpath('.//w:ins')), 1)
        self.assertEqual(len(doc.element.xpath('.//w:del')), 1)
        self.assertEqual(doc.element.xpath('.//w:ins//w:t/text()'), ['建议采用的表述'])
        self.assertEqual(len(doc.comments), 1)  # Only the independent ordinary comment.
        self.assertNotIn('核对理由', ''.join(c.text for c in doc.comments))
        out = self.root/'pending.tex'
        self.assertEqual(sync_word(self.word, self.source, output=out)['changed_blocks'], 0)
        self.assertEqual(out.read_bytes(), self.source.read_bytes())

    def test_native_acceptance_updates_source_and_removes_wrapper(self):
        self.decide(True); out = self.root/'accepted.tex'
        sync_word(self.word, self.source, output=out); value = out.read_text('utf8')
        self.assertIn('建议采用的表述', value); self.assertNotIn('原来的表述', value)
        self.assertNotIn(r'\twrevision', value); self.assertIn(r'\twcomment{note}', value)
        self.source.write_bytes(out.read_bytes()); self.build()
        self.assertEqual(len(Document(self.word).element.xpath('.//w:ins | .//w:del')), 0)

    def test_native_rejection_keeps_original_and_removes_wrapper(self):
        self.decide(False); out = self.root/'rejected.tex'
        sync_word(self.word, self.source, output=out); value = out.read_text('utf8')
        self.assertIn('原来的表述', value); self.assertNotIn('建议采用的表述', value)
        self.assertNotIn(r'\twrevision', value)

    def test_accepted_caption_keeps_table_move_and_pending_note_unique(self):
        profile = tw.load(self.fixture.adapter/'profile.json')
        profile['source_mapping_granularity'] = 'section'
        tw.dump(self.fixture.adapter/'profile.json', profile)
        table = r'''\begin{longtable}{ll}
\caption{\twrevision{caption}{旧表题}{新表题}{表题措辞。}}\label{tab:moved}\tabularnewline
\textbf{模型} & \textbf{结果} \\
甲 & 0.9 \\
\end{longtable}

注：\twrevision{pending-note}{测试集}{评价集}{资料用途。}保持待定。'''
        # Keep substantial section anchors, as in the report, so the source
        # locator identifies a body move rather than moving section headings.
        def retained(prefix):
            return '\n\n'.join(f'{prefix}保留段落{i}。' for i in range(12))
        template = ('\\documentclass{article}\n\\begin{document}\n\\section{方法}\n'
                    + retained('方法前置') + '\n\n%s\n\n' + retained('方法后置')
                    + '\n\\section{基础}\n' + retained('基础前置')
                    + '\n\n%s\n\n' + retained('基础后置') + '\n\\end{document}\n')
        self.source.write_text(template % ('', table), encoding='utf8')
        self.build()
        doc = Document(self.word)
        for node in list(doc.element.xpath('.//w:ins | .//w:del')):
            text = ''.join(t.text or '' for t in node.iter()
                           if t.tag in (W+'t', W+'delText'))
            if text not in ('旧表题', '新表题'):
                continue
            parent = node.getparent(); at = parent.index(node)
            if node.tag == W+'ins':
                for child in list(node):
                    parent.insert(at, child); at += 1
            parent.remove(node)
        doc.save(self.word)
        original_word = self.word.read_bytes()
        self.source.write_text(template % (table, '仅保留基础说明。'), encoding='utf8')
        out = self.root/'moved-table.tex'
        try:
            sync_word(self.word, self.source, output=out)
        except tw.ConversionError:
            import json
            self.fail(str(json.loads(self.word.with_suffix('.sync.json').read_text('utf8'))['conflicts']))
        value = out.read_text('utf8')
        self.assertEqual(value.count(r'\label{tab:moved}'), 1)
        self.assertEqual(value.count(r'\twrevision{pending-note}'), 1)
        self.assertNotIn(r'\twrevision{caption}', value)
        self.assertIn('新表题', value)
        self.assertLess(value.index(r'\label{tab:moved}'), value.index(r'\section{基础}'))
        self.assertEqual(self.word.read_bytes(), original_word)
        self.source.write_bytes(out.read_bytes()); self.build()
        self.assertEqual(sync_word(self.word, self.source)['changed_blocks'], 0)

    def test_resolved_cross_paragraph_bookmark_keeps_user_text(self):
        from tracked_reviews import bookmark_name
        self.decide(True);doc=Document(self.word);manifest=read_manifest(doc)
        start=next(e for e in doc.element.iter(W+'bookmarkStart') if e.get(W+'name')==bookmark_name('change-one'))
        end=next(e for e in doc.element.iter(W+'bookmarkEnd') if e.get(W+'id')==start.get(W+'id'))
        parent=start.getparent();following=OxmlElement('w:p');run=OxmlElement('w:r');text=OxmlElement('w:t');text.text='用户拆分后的后段';run.append(text);following.append(run)
        parent.getparent().insert(parent.getparent().index(parent)+1,following);following.append(end)
        before=doc.element.xpath('.//w:t/text() | .//w:instrText/text()')
        states=project_revisions(doc,manifest)
        self.assertEqual(states['change-one'],'resolved_or_edited')
        self.assertEqual(doc.element.xpath('.//w:t/text() | .//w:instrText/text()'),before)

    def test_pending_or_cross_scope_bookmarks_remain_protected(self):
        from tracked_reviews import bookmark_name
        for case in ('ins','del','rPrChange','other_scope'):
            with self.subTest(case=case):
                doc=Document(self.word);manifest=read_manifest(doc)
                # Resolve the replacement only in this synthetic reader copy.
                for revision in list(doc.element.xpath('.//w:ins | .//w:del')):
                    parent=revision.getparent();at=parent.index(revision)
                    if revision.tag==W+'ins':
                        for child in list(revision):parent.insert(at,child);at+=1
                    parent.remove(revision)
                start=next(e for e in doc.element.iter(W+'bookmarkStart') if e.get(W+'name')==bookmark_name('change-one'))
                end=next(e for e in doc.element.iter(W+'bookmarkEnd') if e.get(W+'id')==start.get(W+'id'))
                parent=start.getparent();following=OxmlElement('w:p');parent.getparent().insert(parent.getparent().index(parent)+1,following)
                if case=='other_scope':
                    control=OxmlElement('w:sdt');properties=OxmlElement('w:sdtPr');tag=OxmlElement('w:tag');tag.set(W+'val','texword:other')
                    properties.append(tag);control.append(properties);other=OxmlElement('w:sdtContent');control.append(other)
                    doc.element.body.insert(len(doc.element.body)-1,control);other.append(following)
                else:following.append(OxmlElement('w:'+case))
                following.append(end)
                with self.assertRaisesRegex(ValueError,'crosses paragraphs'):project_revisions(doc,manifest)

    def test_settled_bookmark_in_math_preserves_equation_and_pending_guards(self):
        from tracked_reviews import bookmark_name
        self.decide(True)
        for case in ('resolved', 'ins', 'rPrChange'):
            with self.subTest(case=case):
                doc=Document(self.word);manifest=read_manifest(doc)
                start=next(e for e in doc.element.iter(W+'bookmarkStart')
                           if e.get(W+'name')==bookmark_name('change-one'))
                paragraph=start.getparent()
                equation=OxmlElement('m:oMath');run=OxmlElement('m:r')
                text=OxmlElement('m:t');text.text='1';run.append(text)
                paragraph.insert(paragraph.index(start),equation)
                equation.append(start);equation.append(run)
                if case!='resolved':equation.append(OxmlElement('w:'+case))
                before=doc.element.xpath('.//w:t/text() | .//m:t/text() | .//w:instrText/text()')
                if case=='resolved':
                    self.assertEqual(project_revisions(doc,manifest)['change-one'],'resolved_or_edited')
                    self.assertEqual(doc.element.xpath('.//w:t/text() | .//m:t/text() | .//w:instrText/text()'),before)
                else:
                    with self.assertRaisesRegex(ValueError,'crosses paragraphs'):project_revisions(doc,manifest)

    def test_resolved_final_table_cell_bookmark_keeps_text_and_guards_pending(self):
        from tracked_reviews import bookmark_name
        self.decide(True)
        for case in ('resolved','ins','rPrChange','later_cell','later_paragraph'):
            with self.subTest(case=case):
                doc=Document(self.word);manifest=read_manifest(doc)
                start=next(e for e in doc.element.iter(W+'bookmarkStart') if e.get(W+'name')==bookmark_name('change-one'))
                end=next(e for e in doc.element.iter(W+'bookmarkEnd') if e.get(W+'id')==start.get(W+'id'))
                paragraph=start.getparent();container=paragraph.getparent();at=container.index(paragraph)
                table=OxmlElement('w:tbl');row=OxmlElement('w:tr');cell=OxmlElement('w:tc')
                table.append(row);row.append(cell);cell.append(paragraph);container.insert(at,table)
                container.insert(container.index(table)+1,end)
                if case in ('ins','rPrChange'):paragraph.append(OxmlElement('w:'+case))
                elif case=='later_cell':
                    another=OxmlElement('w:tc');another.append(OxmlElement('w:p'));row.append(another)
                elif case=='later_paragraph':cell.append(OxmlElement('w:p'))
                before=doc.element.xpath('.//w:t/text() | .//w:instrText/text()')
                if case=='resolved':
                    self.assertEqual(project_revisions(doc,manifest)['change-one'],'resolved_or_edited')
                    self.assertEqual(doc.element.xpath('.//w:t/text() | .//w:instrText/text()'),before)
                else:
                    with self.assertRaisesRegex(ValueError,'crosses paragraphs'):project_revisions(doc,manifest)

    def test_pending_final_table_cell_displaced_end_preserves_revision_and_guards_neighbors(self):
        from tracked_reviews import bookmark_name
        self.source.write_text(SOURCE.replace('核对理由。}\n保留', '核对理由。}\n\n保留'), encoding='utf8')
        self.build()
        for case in ('pending','later_cell','later_paragraph','extra_text'):
            with self.subTest(case=case):
                doc=Document(self.word);manifest=read_manifest(doc)
                start=next(e for e in doc.element.iter(W+'bookmarkStart')
                           if e.get(W+'name')==bookmark_name('change-one'))
                end=next(e for e in doc.element.iter(W+'bookmarkEnd')
                         if e.get(W+'id')==start.get(W+'id'))
                paragraph=start.getparent();container=paragraph.getparent();at=container.index(paragraph)
                # Model a revision occupying the entire terminal cell.
                for node in list(paragraph)[:paragraph.index(start)]:
                    if node.tag!=W+'pPr':paragraph.remove(node)
                table=OxmlElement('w:tbl');row=OxmlElement('w:tr');cell=OxmlElement('w:tc')
                table.append(row);row.append(cell);cell.append(paragraph);container.insert(at,table)
                container.insert(container.index(table)+1,end)
                if case=='later_cell':
                    another=OxmlElement('w:tc');another.append(OxmlElement('w:p'));row.append(another)
                elif case=='later_paragraph':cell.append(OxmlElement('w:p'))
                elif case=='extra_text':
                    run=OxmlElement('w:r');text=OxmlElement('w:t');text.text='neighbor';run.append(text);paragraph.append(run)
                if case=='pending':
                    self.assertEqual(project_revisions(doc,manifest)['change-one'],'pending')
                    self.assertEqual(len(doc.comments),2)
                else:
                    with self.assertRaisesRegex(ValueError,'crosses paragraphs'):
                        project_revisions(doc,manifest)

    def _build_citation_revisions(self):
        self.source.write_text(SOURCE_CITATION_REVISIONS, encoding='utf8')
        self.fixture.bib.write_text(
            fixtures.BIB + '\n' +
            '@article{future, author={Future, Evidence}, title={A relevant study}, '
            'journal={Evidence Journal}, year={2025}, volume={2}, number={1}, pages={1--9}}',
            encoding='utf8',
        )
        self.build()

    def test_pending_citations_are_native_fields_and_empty_sides_are_one_sided(self):
        self._build_citation_revisions()
        doc = Document(self.word)
        self.assertEqual(len(doc.element.xpath('.//w:ins')), 2)
        self.assertEqual(len(doc.element.xpath('.//w:del')), 2)
        manifest = read_manifest(doc)
        references = manifest['references']
        pennes = next(target for target, item in references.items()
                      if item.get('kind') == 'citation' and item.get('key') == 'pennes')
        future = next(target for target, item in references.items()
                      if item.get('kind') == 'citation' and item.get('key') == 'future')
        deletion_fields = ' '.join(doc.element.xpath('.//w:del//w:delInstrText/text()'))
        insertion_fields = ' '.join(doc.element.xpath('.//w:ins//w:instrText/text()'))
        self.assertIn('REF '+pennes, deletion_fields)
        self.assertFalse(doc.element.xpath('.//w:del//w:instrText'))
        self.assertIn('REF '+future, insertion_fields)
        self.assertTrue(all(e.get(W+'fldLock') == 'true'
                            for e in doc.element.xpath('.//w:ins//w:fldChar[@w:fldCharType="begin"]')))
        out = self.root/'citation-pending.tex'
        self.assertEqual(sync_word(self.word, self.source, output=out)['changed_blocks'], 0)
        self.assertEqual(out.read_bytes(), self.source.read_bytes())

    def test_accepting_citation_and_empty_side_revisions_roundtrips_cites(self):
        self._build_citation_revisions(); self.decide(True)
        out = self.root/'citation-accepted.tex'
        sync_word(self.word, self.source, output=out)
        value = out.read_text('utf8')
        self.assertIn(r'\cite{future}', value)
        self.assertNotIn(r'\cite{pennes}', value)
        self.assertIn('新增证据', value)
        self.assertNotIn('待删除引用', value)
        self.assertNotIn(r'\twrevision', value)

    def test_rejecting_citation_and_empty_side_revisions_keeps_original(self):
        self._build_citation_revisions(); self.decide(False)
        out = self.root/'citation-rejected.tex'
        sync_word(self.word, self.source, output=out)
        value = out.read_text('utf8')
        self.assertIn(r'\cite{pennes}', value)
        self.assertNotIn(r'\cite{future}', value)
        self.assertIn('待删除引用', value)
        self.assertNotIn('新增证据', value)
        self.assertNotIn(r'\twrevision', value)

    def test_rejected_insertion_keeps_pending_deletion(self):
        doc = Document(self.word); node = doc.element.xpath('.//w:ins')[0]
        node.getparent().remove(node); doc.save(self.word)
        before = self.word.read_bytes(); out = self.root/'partial.tex'
        sync_word(self.word, self.source, output=out)
        self.assertIn(r'\twrevision{change-one}{原来的表述}{}',out.read_text('utf8'))
        self.assertNotIn('建议采用的表述',out.read_text('utf8'))
        self.assertEqual(self.word.read_bytes(),before)

    def test_accepted_and_edited_insertion_is_ordinary_text_beside_pending_deletion(self):
        profile=tw.load(self.fixture.adapter/'profile.json');profile['source_mapping_granularity']='section'
        tw.dump(self.fixture.adapter/'profile.json',profile);self.build()
        doc=Document(self.word);insertion=doc.element.xpath('.//w:ins')[0]
        parent=insertion.getparent();at=parent.index(insertion)
        for child in list(insertion):
            for text in child.iter(W+'t'):text.text='用户接受后改写的新文字'
            parent.insert(at,child);at+=1
        parent.remove(insertion);doc.save(self.word)
        before=self.word.read_bytes();out=self.root/'partial-edited.tex'
        sync_word(self.word,self.source,output=out);value=out.read_text('utf8')
        self.assertIn(r'\twrevision{change-one}{原来的表述}{}',value)
        self.assertIn('用户接受后改写的新文字',value);self.assertNotIn('建议采用的表述',value)
        self.assertEqual(before,self.word.read_bytes())
        self.source.write_bytes(out.read_bytes());self.build();doc=Document(self.word)
        self.assertEqual(len(doc.element.xpath('.//w:del')),1)
        self.assertEqual(len(doc.element.xpath('.//w:ins')),0)
        self.assertIn('用户接受后改写的新文字',tw.document_text(self.word))
        self.assertEqual(sync_word(self.word,self.source)['changed_blocks'],0)

    def test_accepted_deletion_keeps_pending_insertion(self):
        doc=Document(self.word);deletion=doc.element.xpath('.//w:del')[0]
        deletion.getparent().remove(deletion);doc.save(self.word)
        out=self.root/'partial-insertion.tex';sync_word(self.word,self.source,output=out)
        value=out.read_text('utf8');self.assertNotIn('原来的表述',value)
        self.assertIn(r'\twrevision{change-one}{}{建议采用的表述}',value)
        self.source.write_bytes(out.read_bytes());self.build()
        self.assertEqual(len(Document(self.word).element.xpath('.//w:ins')),1)
        self.assertEqual(len(Document(self.word).element.xpath('.//w:del')),0)

    def test_edited_pending_insertion_keeps_user_text_and_pending_state(self):
        self.source.write_text(SOURCE.replace(
            r'\twrevision{change-one}{原来的表述}{建议采用的表述}',
            r'\twrevision{change-one}{}{建议采用的表述}'), encoding='utf8')
        self.build()
        doc = Document(self.word)
        doc.element.xpath('.//w:ins//w:t')[0].text = '用户改后的建议文字'
        doc.save(self.word)
        before = self.word.read_bytes()
        out = self.root/'edited-insertion.tex'
        sync_word(self.word, self.source, output=out)
        value = out.read_text('utf8')
        self.assertIn(r'\twrevision{change-one}{}{用户改后的建议文字}', value)
        self.assertNotIn('建议采用的表述', value)
        self.assertEqual(self.word.read_bytes(), before)
        self.source.write_bytes(out.read_bytes()); self.build()
        doc = Document(self.word)
        self.assertEqual(len(doc.element.xpath('.//w:ins')), 1)
        self.assertEqual(len(doc.element.xpath('.//w:del')), 0)
        self.assertEqual(doc.element.xpath('.//w:ins//w:t/text()'), ['用户改后的建议文字'])
        self.assertEqual(sync_word(self.word, self.source)['changed_blocks'], 0)

    def test_partial_citation_replacement_keeps_both_reference_roles(self):
        self._build_citation_revisions();doc=Document(self.word)
        insertion=doc.element.xpath('.//w:ins')[0];parent=insertion.getparent();at=parent.index(insertion)
        for child in list(insertion):parent.insert(at,child);at+=1
        parent.remove(insertion);doc.save(self.word)
        out=self.root/'partial-citation.tex';sync_word(self.word,self.source,output=out)
        value=out.read_text('utf8')
        self.assertIn(r'\twrevision{citation-change}{原引用\cite{pennes}}{}',value)
        self.assertIn(r'建议引用\cite{future}',value)
        self.assertIn(r'\twrevision{insert-citation}{}{新增证据\cite{future}}',value)
        self.source.write_bytes(out.read_bytes());self.build()
        self.assertEqual(sync_word(self.word,self.source)['changed_blocks'],0)

    def test_deleted_replacement_with_bookmark_imports_as_deletion(self):
        doc = Document(self.word)
        start = next(e for e in doc.element.iter(W+'bookmarkStart')
                     if e.get(W+'name','').startswith('TWREV_'))
        end = next(e for e in doc.element.iter(W+'bookmarkEnd')
                   if e.get(W+'id') == start.get(W+'id'))
        parent = start.getparent()
        for node in list(parent)[parent.index(start):parent.index(end)+1]:
            parent.remove(node)
        doc.save(self.word)
        out = self.root/'deleted-review.tex'
        sync_word(self.word, self.source, output=out)
        value = out.read_text('utf8')
        self.assertNotIn(r'\twrevision', value)
        self.assertNotIn('原来的表述', value)
        self.assertNotIn('建议采用的表述', value)
        self.assertIn('已经确定的内容', value)
        self.assertIn(r'\twcomment{note}', value)

    def test_deleted_one_sided_replacement_imports_as_deletion(self):
        for index, (anchor, proposal) in enumerate([('原来的表述', ''), ('', '建议采用的表述')]):
            with self.subTest(anchor=anchor, proposal=proposal):
                self.source.write_text(SOURCE.replace(
                    r'\twrevision{change-one}{原来的表述}{建议采用的表述}',
                    r'\twrevision{change-one}{'+anchor+'}{'+proposal+'}'), encoding='utf8')
                self.build(); doc = Document(self.word)
                start = next(e for e in doc.element.iter(W+'bookmarkStart')
                             if e.get(W+'name','').startswith('TWREV_'))
                end = next(e for e in doc.element.iter(W+'bookmarkEnd')
                           if e.get(W+'id') == start.get(W+'id'))
                parent = start.getparent()
                for node in list(parent)[parent.index(start):parent.index(end)+1]:
                    parent.remove(node)
                doc.save(self.word); out = self.root/('one-sided-deleted-'+str(index)+'.tex')
                sync_word(self.word, self.source, output=out)
                value = out.read_text('utf8')
                self.assertNotIn(r'\twrevision', value)
                self.assertNotIn('原来的表述', value)
                self.assertNotIn('建议采用的表述', value)
                self.assertIn('已经确定的内容', value)

    def test_lost_bookmark_with_pending_text_still_stops(self):
        doc = Document(self.word)
        start = next(e for e in doc.element.iter(W+'bookmarkStart')
                     if e.get(W+'name','').startswith('TWREV_'))
        start.getparent().remove(start)
        doc.save(self.word)
        out = self.root/'lost-bookmark.tex'
        with self.assertRaisesRegex(tw.ConversionError, 'bookmark missing'):
            sync_word(self.word, self.source, output=out)
        self.assertFalse(out.exists())

    def test_removed_one_sided_review_is_not_confused_with_other_short_revisions(self):
        source=SOURCE.replace(r'\twrevision{change-one}{原来的表述}{建议采用的表述}',
                              r'\twrevision{change-one}{原来的表述}{}')
        source=source.replace('保留\\twcomment',r'\twrevision{other}{原}{新}{}保留\twcomment')
        self.source.write_text(source,encoding='utf8');self.build();doc=Document(self.word)
        from tracked_reviews import bookmark_name
        start=next(e for e in doc.element.iter(W+'bookmarkStart')
                   if e.get(W+'name')==bookmark_name('change-one'))
        parent=start.getparent();end=next(e for e in parent if e.tag==W+'bookmarkEnd'
                                       and e.get(W+'id')==start.get(W+'id'))
        for e in list(parent)[parent.index(start):parent.index(end)+1]:parent.remove(e)
        doc.save(self.word);out=self.root/'removed-short-neighbor.tex'
        sync_word(self.word,self.source,output=out);value=out.read_text('utf8')
        self.assertNotIn('change-one',value);self.assertIn(r'\twrevision{other}{原}{新}',value)

    def test_lost_bookmark_after_acceptance_imports_resolved_text(self):
        self.decide(True)
        doc = Document(self.word)
        start = next(e for e in doc.element.iter(W+'bookmarkStart')
                     if e.get(W+'name','').startswith('TWREV_'))
        start.getparent().remove(start)
        doc.save(self.word)
        sync_word(self.word, self.source, apply=True)
        value=self.source.read_text('utf8')
        self.assertNotIn('twrevision',value)
        self.assertIn('建议采用的表述',value)

    def test_duplicate_short_pending_reviews_keep_stable_identity(self):
        source=SOURCE.replace(r'\twrevision{change-one}{原来的表述}{建议采用的表述}{核对理由。}',
            r'\twrevision{change-one}{的}{}{同一理由}甲\twrevision{other}{的}{}{同一理由}')
        self.source.write_text(source,encoding='utf8');self.build()
        self.assertEqual(sync_word(self.word,self.source)['changed_blocks'],0)

    def test_word_coalesced_adjacent_insertions_roundtrip(self):
        source=SOURCE.replace(r'\twrevision{change-one}{原来的表述}{建议采用的表述}{核对理由。}',
            r'\twrevision{change-one}{原}{新}{理由}\twrevision{other}{}{增补}{理由}')
        self.source.write_text(source,encoding='utf8');self.build();doc=Document(self.word)
        first,second=doc.element.xpath('.//w:ins');parent=first.getparent()
        # Word can merge pending insertions across touching review boundaries.
        for node in list(parent)[parent.index(first)+1:parent.index(second)+1]:
            if node is second:
                for child in list(second):first.append(child)
                parent.remove(second)
            else:first.append(node)
        doc.save(self.word)
        self.assertEqual(sync_word(self.word,self.source)['changed_blocks'],0)

    def test_review_deletion_and_adjacent_table_edit_keep_table_label(self):
        source = r'''\documentclass{article}
\begin{document}
\section{修订验证}
进度见表\ref{tab:plan}。\twrevision{change-one}{原来的表述}{建议采用的表述}{}

\begin{longtable}{ll}
\caption{研究计划}\label{tab:plan}\\
阶段 & 时间\\
准备 & 待定\\
\end{longtable}
\end{document}
'''
        self.source.write_text(source, encoding='utf8')
        profile = tw.load(self.fixture.adapter/'profile.json')
        profile['source_mapping_granularity'] = 'section'
        tw.dump(self.fixture.adapter/'profile.json', profile); self.build()
        doc = Document(self.word)
        start = next(e for e in doc.element.iter(W+'bookmarkStart')
                     if e.get(W+'name','').startswith('TWREV_'))
        end = next(e for e in doc.element.iter(W+'bookmarkEnd')
                   if e.get(W+'id') == start.get(W+'id'))
        parent = start.getparent()
        for node in list(parent)[parent.index(start):parent.index(end)+1]:
            parent.remove(node)
        next(t for t in doc.element.iter(W+'t') if t.text == '待定').text = '2027.02'
        doc.save(self.word)
        out = self.root/'edited-schedule.tex'
        sync_word(self.word, self.source, output=out)
        value = out.read_text('utf8')
        self.assertEqual(value.count(r'\label{tab:plan}'), 1)
        self.assertIn(r'\ref{tab:plan}', value)
        self.assertIn('2027.02', value)
        self.assertNotIn(r'\twrevision', value)
        self.assertNotIn(r'\endfirsthead', value)
        self.assertNotIn(r'\endhead', value)
        self.assertNotIn(r'\caption{\label', value)
        self.source.write_bytes(out.read_bytes()); self.build()
        text = ''.join(t.text or '' for t in Document(self.word).element.iter(W+'t'))
        self.assertNotIn('[tab:plan]', text)
        self.assertEqual(text.count('阶段'), 1)

    def test_comment_deletion_is_never_native_acceptance(self):
        doc = Document(self.word)
        for node in list(doc.element.xpath('.//w:commentRangeStart | .//w:commentRangeEnd | .//w:commentReference')):
            node.getparent().remove(node)
        doc.save(self.word)
        out = self.root/'comment-removed.tex'
        sync_word(self.word, self.source, output=out)
        value = out.read_text('utf8')
        self.assertIn(r'\twrevision{change-one}{原来的表述}{建议采用的表述}', value)
        self.assertNotIn(r'\twcomment{note}', value)

    def test_prose_edit_preserves_pending_revision_without_cards(self):
        doc = Document(self.word)
        text = next(t for t in doc.element.xpath('.//w:t') if '已经确定' in (t.text or ''))
        text.text = text.text.replace('已经确定', '已经核实')
        doc.save(self.word)
        out = self.root/'edited-pending.tex'
        sync_word(self.word, self.source, output=out)
        self.assertIn(r'\twrevision{change-one}{原来的表述}{建议采用的表述}', out.read_text('utf8'))
        self.source.write_bytes(out.read_bytes()); self.build()
        self.assertEqual(len(Document(self.word).comments), 1)

    def test_section_mapping_pending_and_acceptance(self):
        profile = tw.load(self.fixture.adapter/'profile.json'); profile['source_mapping_granularity'] = 'section'
        tw.dump(self.fixture.adapter/'profile.json', profile); self.build()
        self.assertEqual(sync_word(self.word, self.source)['changed_blocks'], 0)
        self.decide(True); out = self.root/'section-accepted.tex'
        sync_word(self.word, self.source, output=out)
        self.assertIn('建议采用的表述', out.read_text('utf8'))
        self.assertNotIn(r'\twrevision', out.read_text('utf8'))

    def test_word_displaced_end_at_paragraph_boundary_is_not_an_edit(self):
        self.source.write_text(SOURCE.replace('核对理由。}\n保留', '核对理由。}\n\n保留'), encoding='utf8')
        self.build()
        doc = Document(self.word)
        end = next(e for e in doc.element.iter(W+'bookmarkEnd') if e.get(W+'id') in {
            s.get(W+'id') for s in doc.element.iter(W+'bookmarkStart') if s.get(W+'name','').startswith('TWREV_')})
        paragraph = end.getparent(); container = paragraph.getparent()
        container.insert(container.index(paragraph)+1, end)
        end.set(W+'displacedByCustomXml', 'next')
        doc.save(self.word)
        self.assertEqual(sync_word(self.word, self.source)['changed_blocks'], 0)

    def test_widened_bookmark_does_not_capture_neighboring_prose(self):
        doc = Document(self.word)
        end = next(e for e in doc.element.iter(W+'bookmarkEnd') if e.get(W+'id') in {
            s.get(W+'id') for s in doc.element.iter(W+'bookmarkStart') if s.get(W+'name','').startswith('TWREV_')})
        paragraph = end.getparent(); container = paragraph.getparent()
        container.insert(container.index(paragraph)+1, end)
        doc.save(self.word)
        before = self.source.read_bytes()
        with self.assertRaisesRegex(tw.ConversionError, 'outside its replacement'):
            sync_word(self.word, self.source, apply=True)
        self.assertEqual(self.source.read_bytes(), before)


class WordMathDefaultsTests(unittest.TestCase):
    def test_empty_comment_carrier_before_subscript_preserves_real_edits(self):
        def formula(value='20', operator='=-', subscript='d,c', carrier=False):
            root = OxmlElement('m:oMath')
            if carrier:
                run = OxmlElement('m:r'); reference = OxmlElement('w:commentReference')
                reference.set(W+'id', '25'); run.append(reference); root.append(run)
            sub = OxmlElement('m:sSub')
            for name, content in [('e', 'T'), ('sub', subscript)]:
                slot = OxmlElement('m:'+name); run = OxmlElement('m:r'); text = OxmlElement('m:t')
                text.text = content; run.append(text); slot.append(run); sub.append(slot)
            root.append(sub); run = OxmlElement('m:r'); text = OxmlElement('m:t')
            text.text = operator+value+'\u2009°C'; run.append(text); root.append(run)
            return root
        original = math_model(formula(operator='=−'))
        saved = math_model(formula(carrier=True))
        self.assertEqual(saved[3][0], ['text', ''])
        self.assertTrue(equivalent_math(original, saved))
        self.assertEqual(saved[3][0], ['text', ''])  # Keep persisted models and old hashes unchanged.
        for change in [{'value': '21'}, {'operator': '>='}, {'subscript': 'd,x'}]:
            with self.subTest(change=change):
                self.assertFalse(equivalent_math(original, math_model(formula(carrier=True, **change))))
        # An actual mathematical space remains content, unlike the empty carrier.
        spaced = formula(carrier=True); text = OxmlElement('m:t'); text.text = ' '; spaced[0].append(text)
        self.assertFalse(equivalent_math(original, math_model(spaced)))
        lower = ['sub', [], '', [['text', '0']]]; upper = ['sup', [], '', [['text', '1']]]
        base = ['nary', [], '', [lower, upper, ['e', [], '', [['text', 'x']]]]]
        self.assertFalse(equivalent_math(base, ['nary', [], '', [lower, ['sup', [], '', [['text', '2']]], base[3][2]]]))
        self.assertFalse(equivalent_math(base, ['nary', [], '', [['sub', [], '', [['text', '2']]], upper, base[3][2]]]))
        self.assertFalse(equivalent_math(base, ['nary', [], '', [upper, base[3][2]]]))

    def test_default_properties_and_minus_do_not_hide_actual_edits(self):
        original = OxmlElement('m:oMath'); run = OxmlElement('m:r'); text = OxmlElement('m:t')
        text.text = '1−q'; run.append(text); original.append(run)
        saved = OxmlElement('m:oMath'); run2 = OxmlElement('m:r'); text2 = OxmlElement('m:t')
        text2.text = '1-q'; run2.append(text2); saved.append(run2)
        self.assertTrue(equivalent_math(math_model(original), math_model(saved)))
        text2.text = '1+q'
        self.assertFalse(equivalent_math(math_model(original), math_model(saved)))
        hidden = ['subHide', [['val', '0']], '', []]
        self.assertTrue(equivalent_math(['naryPr', [], '', [hidden]], ['naryPr', [], '', []]))
        self.assertFalse(equivalent_math(['begChr', [['val', '[']], '', []], ['begChr', [['val', '(']], '', []]))


if __name__ == '__main__':
    unittest.main()
