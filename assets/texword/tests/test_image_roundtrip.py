"""Word picture edits travel through the same automatic pull/push entrypoint."""
import copy, re, sys, unittest
from base64 import b64decode
from pathlib import Path
from unittest.mock import patch
from docx import Document
from docx.shared import Cm
from docx.oxml import OxmlElement

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import texword as tw
from roundtrip import sync_word
from roundtrip_model import W, NS, read_manifest
from image_roundtrip import A, R, ImageImport
from workflow_runtime import synchronize
from test_sections import SectionTests


class ImageRoundtripTests(unittest.TestCase):
    def setUp(self):
        self.case = SectionTests(); self.case.setUp(); self.f = self.case.f
        self.red = self.f.root / 'red.png'; self.blue = self.f.root / 'blue.png'
        self.red.write_bytes(b64decode('iVBORw0KGgoAAAANSUhEUgAAABQAAAAMCAIAAADtbgqsAAAAHElEQVR4nGP8z0A+YKJAL8OoZhIBE6kaGEa4ZgD9bwEXwe2qmwAAAABJRU5ErkJggg=='))
        self.blue.write_bytes(b64decode('iVBORw0KGgoAAAANSUhEUgAAABQAAAAMCAIAAADtbgqsAAAAHklEQVR4nGNkYPjPQC5gIlsnw6hmkgET6VoYRrRmAPtxARchsutuAAAAAElFTkSuQmCC'))

    def tearDown(self):
        self.case.tearDown()

    def content(self, doc):
        return self.case.controls(doc)[0].find('w:sdtContent', NS)

    def add(self, caption='图 99 新的图片。', container=False):
        doc = Document(self.f.out); content = self.content(doc)
        image = doc.add_paragraph(); image.add_run().add_picture(str(self.red), width=Cm(3))
        cap = doc.add_paragraph(caption)
        if container:
            table = doc.add_table(rows=1, cols=1); cell = table.cell(0, 0)._tc
            for p in list(cell.findall('w:p', NS)): cell.remove(p)
            cell.extend((image._p, cap._p)); content.append(table._tbl)
        else:
            content.extend((image._p, cap._p))
        doc.save(self.f.out)

    def hashes(self):
        from hashlib import sha256
        doc = Document(self.f.out)
        return [sha256(doc.part.related_parts[b.get(R + 'embed')].blob).hexdigest()
                for b in doc.element.body.iter(A + 'blip')]

    def delete(self):
        doc = Document(self.f.out); content = self.content(doc)
        image = next(p for p in content if list(p.iter(W + 'drawing')))
        following = image.getnext(); content.remove(image); content.remove(following)
        doc.save(self.f.out)

    def test_preview_then_import_preserves_bytes_caption_size_and_second_pull_is_noop(self):
        self.add(); before = self.f.source.read_bytes(); word = self.f.out.read_bytes()
        preview = sync_word(self.f.out, self.f.source)
        self.assertEqual(preview['changed_blocks'], 1)
        self.assertEqual(len(preview['image_assets']), 1)
        self.assertFalse((self.f.root / 'figures').exists())
        self.assertEqual(before, self.f.source.read_bytes()); self.assertEqual(word, self.f.out.read_bytes())
        synchronize(self.f.root, 'pull')
        value = self.f.source.read_text(encoding='utf8')
        self.assertIn(r'\caption{新的图片。}', value)
        self.assertNotRegex(value, r'\\caption\{(?:图\s*)?99')
        path = re.search(r'\\includegraphics\[[^]]*\]\{([^}]+)\}', value)[1]
        self.assertEqual((self.f.root / path).read_bytes(), self.red.read_bytes())
        self.assertLessEqual(abs(Document(self.f.out).inline_shapes[0].width - Cm(3)), 1)
        self.assertEqual(sync_word(self.f.out, self.f.source)['changed_blocks'], 0)
        self.assertFalse(synchronize(self.f.root, 'pull')['rebuilt_word'])

    def test_replace_retains_label_references_and_original_asset(self):
        self.add(); synchronize(self.f.root, 'pull')
        value = self.f.source.read_text(encoding='utf8'); label = re.search(r'\\label\{(fig:[^}]+)\}', value)[1]
        value = value.replace(r'\end{figure}', r'\end{figure}' + '\n\n见图\\ref{' + label + '}。')
        self.f.source.write_text(value, encoding='utf8'); synchronize(self.f.root, 'push')
        doc = Document(self.f.out); blip = next(doc.element.body.iter(A + 'blip'))
        rid, _ = doc.part.get_or_add_image(str(self.blue)); blip.set(R + 'embed', rid); doc.save(self.f.out)
        synchronize(self.f.root, 'pull'); value = self.f.source.read_text(encoding='utf8')
        self.assertIn('\\label{' + label + '}', value); self.assertIn('\\ref{' + label + '}', value)
        self.assertEqual(len(list((self.f.root / 'figures').glob('*'))), 2)
        self.assertEqual(self.hashes(), [tw.digest(self.blue)])
        self.assertEqual(sync_word(self.f.out, self.f.source)['changed_blocks'], 0)

    def test_delete_picture_and_caption_removes_source_but_preserves_asset(self):
        self.add(); synchronize(self.f.root, 'pull'); self.delete(); synchronize(self.f.root, 'pull')
        self.assertNotIn('includegraphics', self.f.source.read_text(encoding='utf8'))
        self.assertEqual(self.hashes(), [])
        self.assertEqual(len(list((self.f.root / 'figures').glob('*'))), 1)

    def test_deleted_referenced_figure_is_reported_without_partial_write(self):
        self.add(); synchronize(self.f.root, 'pull')
        value = self.f.source.read_text(encoding='utf8'); label = re.search(r'\\label\{(fig:[^}]+)\}', value)[1]
        self.f.source.write_text(value.replace(r'\end{figure}', r'\end{figure}' + '\n见图\\ref{' + label + '}。'), encoding='utf8')
        synchronize(self.f.root, 'push'); self.delete(); before = self.f.source.read_bytes(); word = self.f.out.read_bytes()
        with self.assertRaisesRegex(tw.ConversionError, 'still referenced'):
            sync_word(self.f.out, self.f.source, apply=True)
        self.assertEqual(before, self.f.source.read_bytes()); self.assertEqual(word, self.f.out.read_bytes())

    def test_figure_in_layout_table_and_caption_edit_keep_label(self):
        self.add(container=True); synchronize(self.f.root, 'pull')
        value = self.f.source.read_text(encoding='utf8'); label = re.search(r'\\label\{(fig:[^}]+)\}', value)[1]
        doc = Document(self.f.out)
        for node in doc.element.body.iter(W + 't'):
            if '新的图片' in (node.text or ''): node.text = node.text.replace('新的图片', '修改后的图注')
        doc.save(self.f.out); synchronize(self.f.root, 'pull')
        value = self.f.source.read_text(encoding='utf8')
        self.assertIn(r'\caption{修改后的图注。}', value); self.assertIn('\\label{' + label + '}', value)
        self.assertEqual(sync_word(self.f.out, self.f.source)['changed_blocks'], 0)

    def test_inline_image_preserves_surrounding_prose_and_bytes(self):
        doc = Document(self.f.out); p = doc.add_paragraph('图像前。')
        p.add_run().add_picture(str(self.red), width=Cm(2)); p.add_run('图像后。')
        self.content(doc).append(p._p); doc.save(self.f.out); synchronize(self.f.root, 'pull')
        value = self.f.source.read_text(encoding='utf8')
        self.assertIn('图像前。', value); self.assertIn('图像后。', value)
        self.assertEqual(self.hashes(), [tw.digest(self.red)])
        self.assertEqual(sync_word(self.f.out, self.f.source)['changed_blocks'], 0)

    def test_disjoint_source_review_edit_and_word_image_both_survive(self):
        self.add(); self.f.source.write_text(self.f.source.read_text(encoding='utf8').replace('现有描述', '源文件修改'), encoding='utf8')
        sync_word(self.f.out,self.f.source,apply=True)
        value=self.f.source.read_text('utf8')
        self.assertIn('源文件修改',value);self.assertIn('includegraphics',value)
        self.assertTrue((self.f.root/'figures').exists())

    def test_asset_publish_rolls_back_new_files_on_source_commit_failure(self):
        images = ImageImport(Document(self.f.out), self.f.source, read_manifest(Document(self.f.out)))
        images.pending['figures/test.png'] = self.red.read_bytes()
        staged = self.f.root / 'staged.tex'; staged.write_text('staged')
        before = self.f.source.read_bytes()
        with patch('image_roundtrip.os.replace', side_effect=OSError('commit failed')):
            with self.assertRaisesRegex(OSError, 'commit failed'): images.publish(self.f.source, staged)
        self.assertEqual(before, self.f.source.read_bytes()); self.assertFalse((self.f.root / 'figures/test.png').exists())

    def test_three_pictures_with_empty_caption_carriers_and_anchor_in_previous_caption(self):
        doc = Document(self.f.out); content = self.content(doc)
        if 'ImageCaption' not in doc.styles: doc.styles.add_style('ImageCaption', 1)
        table = doc.add_table(rows=1, cols=1); cell = table.cell(0, 0)._tc
        for p in list(cell.findall('w:p', NS)): cell.remove(p)
        first = doc.add_paragraph(); first.add_run().add_picture(str(self.red), width=Cm(3))
        blank = doc.add_paragraph(style='ImageCaption')
        cap1 = doc.add_paragraph('图 6 第一张图。', style='ImageCaption')
        second = doc.add_paragraph(); second.add_run().add_picture(str(self.blue), width=Cm(3))
        cap2 = doc.add_paragraph('图7 第二张图。', style='ImageCaption')
        cap2.add_run().add_picture(str(self.red), width=Cm(3))
        cap3 = doc.add_paragraph('图8第三张图。', style='ImageCaption')
        cell.extend(p._p for p in (first, blank, cap1, second, cap2, cap3)); content.append(table._tbl)
        doc.save(self.f.out); synchronize(self.f.root, 'pull')
        value = self.f.source.read_text(encoding='utf8')
        self.assertEqual(value.count(r'\begin{figure}'), 3)
        for title in ('第一张图。', '第二张图。', '第三张图。'):
            self.assertIn('\\caption{' + title + '}', value)
        labels = re.findall(r'\\label\{(fig:[^}]+)\}', value)
        self.assertEqual(len(set(labels)), 3)
        self.assertEqual(self.hashes(), [tw.digest(self.red), tw.digest(self.blue), tw.digest(self.red)])
        self.assertEqual(sync_word(self.f.out, self.f.source)['changed_blocks'], 0)

    def test_cropped_image_roundtrip_preserves_crop_dimensions_and_original_bytes(self):
        self.add(); doc = Document(self.f.out)
        picture = doc.inline_shapes[0]
        crop = OxmlElement('a:srcRect');crop.set('t','1588');crop.set('b','3472')
        picture._inline.graphic.graphicData.pic.blipFill.insert(1,crop)
        width, height = picture.width, picture.height
        doc.save(self.f.out); synchronize(self.f.root, 'pull')
        self.assertIn(r'\twimagecrop{0,1588,0,3472}', self.f.source.read_text('utf8'))
        picture = Document(self.f.out).inline_shapes[0]
        crop = next(picture._inline.iter(A+'srcRect'))
        self.assertEqual((crop.get('t'), crop.get('b')), ('1588', '3472'))
        self.assertLessEqual(abs(picture.width-width), 1)
        self.assertLessEqual(abs(picture.height-height), 1)
        self.assertEqual(self.hashes(), [tw.digest(self.red)])
        self.assertEqual(sync_word(self.f.out, self.f.source)['changed_blocks'], 0)

    def test_invalid_crop_preserves_word_and_creates_no_assets(self):
        self.add(); doc = Document(self.f.out); blip = next(doc.element.body.iter(A + 'blip'))
        crop = OxmlElement('a:srcRect'); crop.set('l', '100000'); blip.getparent().append(crop)
        doc.save(self.f.out); before = self.f.source.read_bytes(); word = self.f.out.read_bytes()
        with self.assertRaises(tw.ConversionError): sync_word(self.f.out, self.f.source, apply=True)
        self.assertEqual(before, self.f.source.read_bytes()); self.assertEqual(word, self.f.out.read_bytes())
        self.assertFalse((self.f.root / 'figures').exists())

    def test_anchored_picture_in_prose_with_separate_caption_table_preserves_label(self):
        self.add();synchronize(self.f.root,'pull')
        label=re.search(r'\\label\{(fig:[^}]+)\}',self.f.source.read_text('utf8'))[1]
        doc=Document(self.f.out);content=self.content(doc)
        picture=next(p for p in content if list(p.iter(W+'drawing')))
        cap=picture.getnext();drawing=next(picture.iter(W+'drawing'))
        inline=drawing[0];inline.tag='{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}anchor'
        prose=doc.add_paragraph('图片前的正文。');prose.add_run()._r.append(drawing)
        content.replace(picture,prose._p)
        table=doc.add_table(rows=1,cols=1);cell=table.cell(0,0)._tc
        for p in list(cell.findall('w:p',NS)):cell.remove(p)
        content.replace(cap,table._tbl);cell.append(cap)
        doc.save(self.f.out);synchronize(self.f.root,'pull')
        source=self.f.source.read_text('utf8')
        self.assertIn('图片前的正文。',source)
        self.assertIn('\\label{'+label+'}',source)
        self.assertNotIn('longtable',source[source.index('图片前的正文。'):])
        self.assertEqual(self.hashes(),[tw.digest(self.red)])

    def test_crop_parser_ignores_comments_and_accepts_whitespace(self):
        import json
        source='% '+r'\twimagecrop{1,2,3,4}\includegraphics{red.png}'+'\n'+r'\includegraphics{red.png}'+ '\n\n'+r'\twimagecrop{0,1588,0,3472}'+ '\n'+r'\includegraphics{red.png}'
        ast=json.loads(tw.pandoc(['-f','latex','-t','json'],input=source)[0])
        images=[]
        def visit(value):
            if isinstance(value,list):
                for child in value:visit(child)
            elif isinstance(value,dict):
                if value.get('t')=='Image':images.append(value['c'][2][1])
                for child in value.values():visit(child)
        visit(ast)
        self.assertEqual(images,['','TWIMAGECROP:0,1588,0,3472'])

    def test_existing_picture_dimension_only_edit_is_imported(self):
        self.add();synchronize(self.f.root,'pull');doc=Document(self.f.out)
        extent=doc.element.xpath('.//wp:extent')[0];extent.set('cx',str(4*360000))
        doc.save(self.f.out)
        self.assertGreater(sync_word(self.f.out,self.f.source,apply=True)['changed_blocks'],0)
        self.assertIn('width=4cm',self.f.source.read_text('utf8'))
        self.assertEqual(self.hashes(),[tw.digest(self.red)])

    def test_existing_picture_invalid_crop_and_rotation_are_rejected_before_import(self):
        self.add();synchronize(self.f.root,'pull');baseline=self.f.out.read_bytes()
        for kind in ('crop','rotation'):
            with self.subTest(kind=kind):
                self.f.out.write_bytes(baseline);doc=Document(self.f.out)
                if kind=='crop':
                    blip=next(doc.element.body.iter(A+'blip'));crop=OxmlElement('a:srcRect');crop.set('l','100000');blip.getparent().append(crop)
                else:
                    transform=next(doc.element.body.iter(A+'xfrm'));transform.set('rot','60000')
                doc.save(self.f.out);word=self.f.out.read_bytes();source=self.f.source.read_bytes()
                with self.assertRaises(tw.ConversionError):sync_word(self.f.out,self.f.source,apply=True)
                self.assertEqual(word,self.f.out.read_bytes());self.assertEqual(source,self.f.source.read_bytes())

    def test_derived_output_has_its_own_assets_and_leaves_source_untouched(self):
        self.add(); before = self.f.source.read_bytes(); output = self.f.root / 'derived/result.tex'
        sync_word(self.f.out, self.f.source, output=output)
        path = re.search(r'\\includegraphics\[[^]]*\]\{([^}]+)\}', output.read_text(encoding='utf8'))[1]
        self.assertEqual((output.parent / path).read_bytes(), self.red.read_bytes())
        self.assertEqual(before, self.f.source.read_bytes())

    def test_block_mapping_figure_replacement_also_imports_automatically(self):
        profile = tw.load(self.f.adapter / 'profile.json'); profile['source_mapping_granularity'] = 'block'
        tw.dump(self.f.adapter / 'profile.json', profile)
        value = self.f.source.read_text(encoding='utf8').replace(r'\section*{参考文献}',
            '\\begin{figure}\\includegraphics[width=3cm]{red.png}\\caption{已有图}\\label{fig:existing}\\end{figure}\n\n' + r'\section*{参考文献}')
        self.f.source.write_text(value, encoding='utf8'); self.f.build()
        doc = Document(self.f.out); blip = next(doc.element.body.iter(A + 'blip'))
        rid, _ = doc.part.get_or_add_image(str(self.blue)); blip.set(R + 'embed', rid); doc.save(self.f.out)
        synchronize(self.f.root, 'pull')
        self.assertEqual(self.hashes(), [tw.digest(self.blue)])
        self.assertIn(r'\label{fig:existing}', self.f.source.read_text(encoding='utf8'))
        self.assertEqual(sync_word(self.f.out, self.f.source)['changed_blocks'], 0)

    def test_heading_semantics_survive_template_outline_mismatch_on_rebuild(self):
        from word_runtime import TemplateDocument
        profile = copy.deepcopy(tw.load(self.f.adapter / 'profile.json'))
        profile['roles']['heading3']['pPr'] = '<w:pPr xmlns:w="' + NS['w'] + '"><w:outlineLvl w:val="1"/></w:pPr>'
        built = TemplateDocument(self.f.adapter / 'template.docx', profile, self.f.adapter)
        style = next(s for s in built.doc.styles.element if s.get(W + 'styleId') == 'Heading3')
        self.assertEqual(style.find('w:pPr/w:outlineLvl', NS).get(W + 'val'), '2')

    def test_caption_formatting_and_citation_are_preserved(self):
        self.add(); doc = Document(self.f.out); cap = self.content(doc)[-1]
        for child in list(cap): cap.remove(child)
        from lxml import etree as E
        xml = tw.run('图 99 ') + '<w:r xmlns:w="' + NS['w'] + '"><w:rPr><w:b/></w:rPr><w:t>加粗图注</w:t></w:r>'
        xml += tw.run('，参考[') + tw.field('REF BIB_one \\h', 1) + tw.run(']。')
        for child in E.fromstring(('<root xmlns:w="' + NS['w'] + '">' + xml + '</root>').encode()): cap.append(child)
        doc.save(self.f.out); synchronize(self.f.root, 'pull')
        value = self.f.source.read_text(encoding='utf8')
        self.assertIn(r'\textbf{加粗图注}', value); self.assertIn(r'\cite{BIB_one}', value)
        self.assertEqual(sync_word(self.f.out, self.f.source)['changed_blocks'], 0)

    def test_caption_edit_does_not_discard_invisible_source_comments(self):
        self.add(); synchronize(self.f.root, 'pull')
        value = self.f.source.read_bytes().decode('utf8').replace(r'\centering', '\\centering % source note\n')
        self.f.source.write_bytes(value.encode('utf8')); synchronize(self.f.root, 'push')
        doc = Document(self.f.out)
        for node in doc.element.body.iter(W + 't'):
            if '新的图片' in (node.text or ''): node.text = node.text.replace('新的图片', '修改图注')
        doc.save(self.f.out); before = self.f.source.read_bytes()
        with self.assertRaises(tw.ConversionError): sync_word(self.f.out, self.f.source, apply=True)
        self.assertEqual(before, self.f.source.read_bytes())
        self.assertIn('LaTeX comments', (self.f.out.with_suffix('.sync.json')).read_text(encoding='utf8'))


if __name__ == '__main__': unittest.main()
