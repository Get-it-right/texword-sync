"""Native Word replacement revisions, projected without accepting pending edits.

Only revisions created by twrevision are managed. Stable bookmarks delimit each
proposal independently of Word's revision and comment ID renumbering.
"""
from __future__ import annotations
import copy
import hashlib
import re
from datetime import datetime, timezone
from lxml import etree as E

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XMLSPACE = '{http://www.w3.org/XML/1998/namespace}space'


def bookmark_name(key):
    return 'TWREV_' + hashlib.sha256(key.encode('utf8')).hexdigest()[:26]


def _node(tag, **attrs):
    value = E.Element(W + tag)
    for key, item in attrs.items():
        value.set(W + key, str(item))
    return value


def _plain(nodes, deleted=False):
    tag = W + ('delText' if deleted else 't')
    return ''.join(t.text or '' for node in nodes for t in node.iter(tag))


def _remove_comment(doc, cid):
    """Remove only the redundant card belonging to a managed replacement."""
    for node in list(doc.element.iter()):
        if node.tag in (W+'commentRangeStart', W+'commentRangeEnd', W+'commentReference') and node.get(W+'id') == cid:
            parent = node.getparent(); parent.remove(node)
            if parent.tag == W+'r' and all(c.tag == W+'rPr' for c in parent):
                parent.getparent().remove(parent)
    for comment in list(doc.comments):
        if str(comment.comment_id) == cid:
            element = comment._element; element.getparent().remove(element)


def _restore_pending_anchor(doc, spec, start, end):
    """Create an in-memory mapping anchor; never save a card to the user's file."""
    parent = start.getparent()
    nodes = list(parent)[parent.index(start)+1:parent.index(end)]
    markers=[node for node in nodes if node.tag == W+'commentRangeStart']
    if markers:
        from review_comments import read_comments
        comments=read_comments(doc)
        if len(markers)==1 and _legacy_card(comments,markers[0].get(W+'id'),spec):
            markers[0].set('{urn:texword:roundtrip:1}managedOrigin', str(spec['native_id']))
            return
        raise ValueError('User comment overlaps pending tracked proposal; preserve both in Word: '+spec['key'])
    comment = doc.comments.add_comment(text=spec['rendered_text'], author=spec.get('author', 'Codex'), initials='')
    cid = str(comment.comment_id)
    marker=_node('commentRangeStart', id=cid)
    marker.set('{urn:texword:roundtrip:1}managedOrigin', str(spec['native_id']))
    parent.insert(parent.index(start)+1, marker)
    parent.insert(parent.index(end), _node('commentRangeEnd', id=cid))
    run = _node('r'); run.append(_node('commentReference', id=cid))
    parent.insert(parent.index(end), run)


def _legacy_card(comments,cid,spec):
    comment=comments.get(cid,{})
    return (cid==str(spec['native_id']) and comment.get('text')==spec['rendered_text']
            and comment.get('author')==spec.get('author','Codex'))


def _validate_inline_runs(nodes, allowed_targets, document_targets, key):
    """Allow plain runs and complete native REF fields, rejecting other objects."""
    active = None
    for node in nodes:
        if node.tag != W + 'r':
            raise ValueError('Tracked replacement anchor must contain ordinary text or citation fields: ' + key)
        for child in node:
            if child.tag == W + 'rPr':
                continue
            if child.tag == W + 't':
                if active is not None and not active['separated']:
                    raise ValueError('Malformed citation field in tracked replacement: ' + key)
                continue
            if child.tag == W + 'instrText':
                if active is None or active['separated']:
                    raise ValueError('Malformed citation field in tracked replacement: ' + key)
                active['instruction'] += child.text or ''
                continue
            if child.tag != W + 'fldChar':
                raise ValueError('Tracked replacement anchor must contain ordinary text or citation fields: ' + key)
            kind = child.get(W + 'fldCharType')
            if kind == 'begin' and active is None:
                active = {'instruction': '', 'separated': False}
            elif kind == 'separate' and active is not None and not active['separated']:
                active['separated'] = True
            elif kind == 'end' and active is not None and active['separated']:
                match = re.fullmatch(r'\s*REF\s+(\S+)(?:\s+\\h)?\s*', active['instruction'], re.IGNORECASE)
                if match is None or match.group(1) not in allowed_targets or match.group(1) not in document_targets:
                    raise ValueError('Tracked replacement citation must be a mapped native REF field: ' + key)
                active = None
            else:
                raise ValueError('Malformed or unsupported field in tracked replacement: ' + key)
    if active is not None:
        raise ValueError('Incomplete citation field in tracked replacement: ' + key)


def install_revisions(doc, definitions, render_tokens=None):
    """Wrap rendered anchor runs after the source baseline has been recorded."""
    existing_ids = [int(e.get(W + 'id')) for e in doc.element.iter(W + 'bookmarkStart')]
    next_id = max(existing_ids + [800000]) + 1
    document_targets = set(doc.element.xpath('.//w:bookmarkStart/@w:name'))
    existing_rev = [int(e.get(W + 'id')) for e in doc.element.iter() if e.tag in (W+'ins', W+'del')]
    rev_id = max(existing_rev + [800000]) + 1
    for spec in definitions:
        if spec['kind'] != 'revision':
            continue
        cid = str(spec['native_id'])
        starts = [e for e in doc.element.iter(W+'commentRangeStart') if e.get(W+'id') == cid]
        ends = [e for e in doc.element.iter(W+'commentRangeEnd') if e.get(W+'id') == cid]
        if len(starts) != 1 or len(ends) != 1 or starts[0].getparent() is not ends[0].getparent():
            raise ValueError('Tracked replacement requires one paragraph: ' + spec['key'])
        start, end = starts[0], ends[0]
        parent = start.getparent()
        nodes = list(parent)[parent.index(start)+1:parent.index(end)]
        anchor_empty = bool(spec.get('anchor_empty', not spec.get('anchor', '').strip()))
        proposal_empty = bool(spec.get('proposal_empty', not spec.get('proposal', '').strip()))
        if anchor_empty and proposal_empty:
            raise ValueError('Tracked revision cannot have both sides empty: ' + spec['key'])
        if anchor_empty and nodes:
            raise ValueError('Tracked replacement has content inside its empty anchor: ' + spec['key'])
        if not anchor_empty:
            if not nodes:
                raise ValueError('Tracked replacement anchor rendered empty: ' + spec['key'])
            _validate_inline_runs(nodes, set(spec.get('citation_targets', [])), document_targets, spec['key'])
        if '\n' in spec.get('proposal', '') or '\r' in spec.get('proposal', ''):
            raise ValueError('Tracked replacement proposal must be single-paragraph text')

        proposal_nodes = []
        if not proposal_empty:
            if render_tokens is not None and 'proposal_tokens' in spec:
                holder = E.Element(W + 'p')
                render_tokens(holder, spec['proposal_tokens'])
                proposal_nodes = list(holder)
            else:
                run = _node('r')
                text = _node('t'); text.set(XMLSPACE, 'preserve'); text.text = spec['proposal']
                run.append(text); proposal_nodes = [run]
            _validate_inline_runs(proposal_nodes, set(spec.get('citation_targets', [])), document_targets, spec['key'])
            if not proposal_nodes or _plain(proposal_nodes) != spec['proposal']:
                raise ValueError('Tracked replacement proposal tokens do not match their display text: ' + spec['key'])
            # Word's Fields.Update can move a REF result outside its enclosing
            # insertion revision. Keep proposal citation fields fixed at the
            # source-registry value while the reviewer decides the revision.
            for node in proposal_nodes:
                for field_char in node.iter(W+'fldChar'):
                    if field_char.get(W+'fldCharType') == 'begin':
                        field_char.set(W+'fldLock', 'true')

        attributes = {'author': spec.get('author', 'Codex'), 'date': datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')}
        style = next((node.find(W+'rPr') for node in nodes if node.find(W+'rPr') is not None), None)
        deletion = None
        if not anchor_empty:
            deletion = _node('del', id=rev_id, **attributes); rev_id += 1
            for node in nodes:
                deletion.append(node)
                for t in node.iter(W+'t'):
                    t.tag = W+'delText'
                for instruction in node.iter(W+'instrText'):
                    instruction.tag = W+'delInstrText'
        insertion = None
        if not proposal_empty:
            insertion = _node('ins', id=rev_id, **attributes); rev_id += 1
            for node in proposal_nodes:
                if style is not None and node.tag == W+'r' and node.find(W+'rPr') is None:
                    node.insert(0, copy.deepcopy(style))
                insertion.append(node)
        # Moving anchor runs into <w:del> changes the end bookmark's index.
        # Re-read it after the move so both revision containers stay inside the
        # stable range used by the roundtrip importer.
        at = parent.index(end)
        if deletion is not None:
            parent.insert(at, deletion); at += 1
        if insertion is not None:
            parent.insert(at, insertion)
        parent.insert(parent.index(start), _node('bookmarkStart', id=next_id, name=bookmark_name(spec['key'])))
        parent.insert(parent.index(end)+1, _node('bookmarkEnd', id=next_id))
        next_id += 1
        # Native insertions/deletions already expose the proposed wording.
        # Keep stable bookmarks, but do not duplicate the change in a card.
        _remove_comment(doc, cid)


def _remaining_revision_source(doc, nodes, manifest, key, *, deleted=False):
    """Reuse native text/citation import for the still-pending side only."""
    from roundtrip import block_to_latex
    paragraph = _node('p')
    for revision in nodes:
        for child in revision:
            child = copy.deepcopy(child)
            if deleted:
                for text in child.iter(W+'delText'):
                    text.tag = W+'t'
                for instruction in child.iter(W+'delInstrText'):
                    instruction.tag = W+'instrText'
            paragraph.append(child)
    value = block_to_latex(doc, [paragraph],
        {'id':key, 'kind':'paragraph', 'latex':'', 'math_sources':[]}, manifest, {})
    plain = _plain(nodes, deleted=deleted)
    leading = re.match(r'\s*', plain)[0]
    trailing = re.search(r'\s*$', plain)[0]
    return leading + value.strip() + trailing


def _project_remaining_side(doc, manifest, spec, start, end, revisions, *, deleted):
    """Keep resolved text in place and narrow metadata to the unresolved side."""
    from review_syntax import encode_plain, prepare_reviews, format_review
    parent = start.getparent()
    first, last = revisions[0], revisions[-1]
    middle = list(parent)[parent.index(first):parent.index(last)+1]
    if any(node not in revisions and (_plain([node]) or _plain([node], deleted=True)) for node in middle):
        raise ValueError('Remaining revisions are interleaved with ordinary text: '+spec['key'])
    value = _remaining_revision_source(doc, revisions, manifest, spec['key'], deleted=deleted)
    anchor, proposal = (value, '') if deleted else ('', value)
    raw = (r'\twrevision['+encode_plain(spec.get('author', 'Codex'))+']{'+spec['key']+
           '}{'+anchor+'}{'+proposal+'}{'+encode_plain(spec.get('reason', ''))+'}')
    _, definitions = prepare_reviews(raw)
    spec.update(definitions[0])
    # The reader compares cached native citation numbers, while the source keeps
    # semantic citation keys. The compiler regenerates proposal tokens later.
    spec['proposal'] = '' if deleted else _plain(revisions)
    spec['rendered_text'] = format_review(spec)
    parent.remove(start); parent.insert(parent.index(first), start)
    parent.remove(end); parent.insert(parent.index(last)+1, end)


def project_revisions(doc, manifest, *, project_accepted=False):
    """Project pending edits to their original text for ordinary source mapping.

    Word acceptance/rejection removes ins/del. In that case the remaining body
    text is authoritative and its annotation wrapper is removed during import.
    This only changes the in-memory reader document, never the user's DOCX.
    """
    states = {}
    # Other intact proposal ranges cannot be remnants of a removed proposal.
    # Inspect ownership before projection removes their bookmarks/containers.
    managed_names = {bookmark_name(s['key']) for s in manifest.get('reviews', [])
                     if s['kind'] == 'revision'}
    managed_ids = {e.get(W+'id') for e in doc.element.iter(W+'bookmarkStart')
                   if e.get(W+'name') in managed_names}
    # Word coalesces adjacent insertions and moves their boundary bookmarks
    # inside the combined insertion. Split the reader containers at those exact
    # boundaries; text, revision sides and paragraph membership stay unchanged.
    for revision in list(doc.element.iter()):
        if revision.tag not in (W+'ins', W+'del') or revision.getparent().tag != W+'p':
            continue
        if not any(c.tag in (W+'bookmarkStart', W+'bookmarkEnd')
                   and c.get(W+'id') in managed_ids for c in revision):
            continue
        parent=revision.getparent();at=parent.index(revision)
        piece=E.Element(revision.tag,dict(revision.attrib))
        for child in list(revision):
            if child.tag in (W+'bookmarkStart', W+'bookmarkEnd') and child.get(W+'id') in managed_ids:
                if len(piece):parent.insert(at,piece);at+=1
                parent.insert(at,child);at+=1
                piece=E.Element(revision.tag,dict(revision.attrib))
            else:piece.append(child)
        if len(piece):parent.insert(at,piece)
        parent.remove(revision)
    # Touching ranges may be serialized start-B/end-A at the same character
    # position. Close A before opening B, without moving across any text.
    for end in list(doc.element.iter(W+'bookmarkEnd')):
        if end.get(W+'id') not in managed_ids:continue
        previous=end.getprevious()
        while (previous is not None and previous.tag==W+'bookmarkStart'
               and previous.get(W+'id') in managed_ids
               and previous.get(W+'id')!=end.get(W+'id')):
            previous.addprevious(end);previous=end.getprevious()
    owned = set()
    for marker in doc.element.iter(W+'bookmarkStart'):
        if marker.get(W+'name') not in managed_names:
            continue
        parent = marker.getparent()
        end = next((e for e in parent if e.tag == W+'bookmarkEnd'
                    and e.get(W+'id') == marker.get(W+'id')), None)
        if end is not None:
            for node in list(parent)[parent.index(marker)+1:parent.index(end)]:
                owned.update(e for e in node.iter() if e.tag in (W+'ins', W+'del'))
    unowned = [e for e in doc.element.iter()
               if e.tag in (W+'ins', W+'del') and e not in owned]
    for spec in manifest.get('reviews', []):
        if spec['kind'] != 'revision':
            continue
        starts = [e for e in doc.element.iter(W+'bookmarkStart') if e.get(W+'name') == bookmark_name(spec['key'])]
        if not starts:
            # Accepting/rejecting a short revision may remove its bookmark.
            # Plain copies of common words elsewhere are not pending revisions.
            # The ordinary source importer preserves the user's resulting text.
            position=manifest.get('source','').find(spec.get('raw',''))
            record=next((r for r in manifest.get('spans',[])
                         if position>=0 and r['start']<=position<r['end']),None)
            scope=doc.element
            if record is not None:
                for control in doc.element.iter(W+'sdt'):
                    tag=control.find(W+'sdtPr/'+W+'tag')
                    if tag is not None and tag.get(W+'val')=='texword:'+record['id']:
                        scope=control;break
            local_unowned=[e for e in unowned if e is scope or scope in e.iterancestors()]
            if not local_unowned:
                states[spec['key']]='resolved_or_deleted'
                continue
            # Deleting the complete replacement in Word also removes its
            # bookmark. Let the ordinary source mapping import that deletion;
            # a lost bookmark with surviving original/proposed text must stop.
            text = _plain([doc.element]) + _plain([doc.element], deleted=True)
            pending = [_plain([e]) + _plain([e], deleted=True) for e in unowned]
            if ((not spec['anchor'] or spec['anchor'] not in text)
                    and (not spec['proposal'] or spec['proposal'] not in text)
                    and not any(t and (t in spec['anchor'] or t in spec['proposal']) for t in pending)):
                states[spec['key']] = 'deleted'
                continue
        if len(starts) != 1:
            raise ValueError('Tracked proposal bookmark missing; preserve the Word file for reconciliation: '+spec['key'])
        start = starts[0]; parent = start.getparent(); bid = start.get(W+'id')
        ends = [e for e in doc.element.iter(W+'bookmarkEnd') if e.get(W+'id') == bid]
        if len(ends) != 1:
            raise ValueError('Tracked proposal crosses paragraphs: '+spec['key'])
        end = ends[0]
        # Editing settled text beside an equation can move its old bookmark
        # into OMML. It no longer delimits a pending proposal. Remove only
        # these markers in the reader copy, preserving every formula/text node.
        if end.getparent() is not parent:
            start_paragraph = next(start.iterancestors(W+'p'), None)
            end_paragraph = next(end.iterancestors(W+'p'), None)
            if start_paragraph is not None and start_paragraph is end_paragraph:
                ordered = list(start_paragraph.iter())
                a, b = ordered.index(start), ordered.index(end)
                def pending_boundary(node):
                    return (node.tag in (W+'ins', W+'del', W+'moveFrom', W+'moveTo',
                                         W+'moveFromRangeStart', W+'moveToRangeStart')
                            or E.QName(node).localname.endswith('PrChange'))
                if (a < b and not any(pending_boundary(node) for node in ordered[a:b+1])
                        and not any(pending_boundary(node) for marker in (start,end)
                                    for node in marker.iterancestors())):
                    start.getparent().remove(start)
                    end.getparent().remove(end)
                    states[spec['key']] = 'resolved_or_edited'
                    continue
        # Word can move an end-of-paragraph bookmark immediately outside its
        # paragraph at a content-control boundary (displacedByCustomXml).
        # This spans no additional text; restore only that exact boundary form.
        if end.getparent() is not parent:
            if end.getparent() is parent.getparent() and end.getprevious() is parent:
                parent.append(end)
            else:
                # Accept/reject followed by paragraph editing can leave the old
                # bookmark across paragraphs. With no pending native changes
                # in that exact interval, the user's settled text owns it.
                def mapped_scope(marker):
                    for ancestor in marker.iterancestors(W+'sdt'):
                        tag=ancestor.find(W+'sdtPr/'+W+'tag')
                        if tag is not None and tag.get(W+'val','').startswith('texword:'):
                            return ancestor
                    return None
                scope=mapped_scope(start);other=end.getparent()
                pending=(W+'ins',W+'del',W+'moveFrom',W+'moveTo',W+'moveFromRangeStart',W+'moveToRangeStart')
                def has_pending(node):
                    return node.tag in pending or E.QName(node).localname.endswith('PrChange')
                same_paragraph_container=(parent.tag==W+'p' and other.tag==W+'p'
                                          and parent.getparent() is other.getparent())
                # A resolved range in the final table cell may end immediately
                # after the table at its content-control boundary on Word save.
                cell=parent.getparent();row=cell.getparent() if cell is not None else None
                table=row.getparent() if row is not None else None
                terminal_table_cell=(parent.tag==W+'p' and cell.tag==W+'tc'
                    and row is not None and row.tag==W+'tr' and table is not None and table.tag==W+'tbl'
                    and cell[-1] is parent and row[-1] is cell and table[-1] is row
                    and table.getparent() is other and other.index(end)==other.index(table)+1)
                # Word also displaces this exact boundary while a revision is
                # pending. Reattach only an intact local payload; the range
                # must cover no later cell, paragraph, or ordinary prose.
                tail=list(parent)[parent.index(start)+1:]
                local_del=[node for node in tail if node.tag==W+'del']
                local_ins=[node for node in tail if node.tag==W+'ins']
                intact_pending=(bool(local_del or local_ins)
                    and _plain(local_del,deleted=True)==spec['anchor']
                    and _plain(local_ins)==spec['proposal']
                    and not any(_plain([node]) or _plain([node],deleted=True)
                                for node in tail if node not in local_del+local_ins))
                if (terminal_table_cell and scope is not None
                        and mapped_scope(end) is scope and intact_pending):
                    parent.append(end)
                elif ((same_paragraph_container or terminal_table_cell)
                        and scope is not None and mapped_scope(end) is scope):
                    ordered=list(scope.iter());a=ordered.index(start);b=ordered.index(end)
                    if (a<b and not any(has_pending(node) for node in ordered[a:b+1])
                            and not any(has_pending(node) for marker in (start,end) for node in marker.iterancestors())):
                        parent.remove(start);other.remove(end)
                        states[spec['key']]='resolved_or_edited'
                        continue
                if end.getparent() is not parent:
                    raise ValueError('Tracked proposal crosses paragraphs: '+spec['key'])
        nodes = list(parent)[parent.index(start)+1:parent.index(end)]
        deletions = [e for e in nodes if e.tag == W+'del']
        insertions = [e for e in nodes if e.tag == W+'ins']
        if deletions or insertions:
            expect_deletion = not bool(spec.get('anchor_empty', not spec.get('anchor', '').strip()))
            expect_insertion = not bool(spec.get('proposal_empty', not spec.get('proposal', '').strip()))
            if bool(deletions) != expect_deletion or bool(insertions) != expect_insertion:
                if not (expect_deletion and expect_insertion):
                    raise ValueError('Unexpected remaining revision side: '+spec['key'])
                _project_remaining_side(doc, manifest, spec, start, end,
                    deletions or insertions, deleted=bool(deletions))
                nodes = list(parent)[parent.index(start)+1:parent.index(end)]
            if _plain(insertions) != spec['proposal']:
                if deletions:
                    raise ValueError('Pending proposed text was edited; reconcile before replacing Word: '+spec['key'])
                # Editing an unresolved insertion changes its proposal, not its
                # acceptance state. Reuse the one-sided reader projection so
                # Word text/citations become the new pending source proposal.
                _project_remaining_side(doc, manifest, spec, start, end,
                    insertions, deleted=False)
                nodes = list(parent)[parent.index(start)+1:parent.index(end)]
            if any(_plain([node]) or _plain([node], deleted=True) for node in nodes if node not in deletions+insertions):
                raise ValueError('Pending tracked range contains text outside its replacement: '+spec['key'])
            for deletion in deletions:
                at = parent.index(deletion)
                for child in list(deletion):
                    for t in child.iter(W+'delText'):
                        t.tag = W+'t'
                    for instruction in child.iter(W+'delInstrText'):
                        instruction.tag = W+'instrText'
                    parent.insert(at, child); at += 1
                parent.remove(deletion)
            for insertion in insertions:
                parent.remove(insertion)
            _restore_pending_anchor(doc, spec, start, end)
            states[spec['key']] = 'pending'
        else:
            states[spec['key']] = 'accepted' if _plain(nodes) == spec['proposal'] else 'rejected_or_edited'
            # Exact acceptance of plain text is a review decision, not a table
            # edit. Canonicalize only the reader copy so an accepted caption
            # does not force the unchanged native table through Pandoc. The
            # merge applies the captured decision to all three source sides.
            if (project_accepted and states[spec['key']] == 'accepted'
                    and '\\' not in spec['anchor']
                    and all(node.tag == W+'r' and all(c.tag in (W+'rPr', W+'t')
                            for c in node) for node in nodes)):
                manifest.setdefault('accepted_review_sources', {})[spec['key']] = {
                    'raw': spec['raw'],
                    'source': spec.get('proposal_source', spec['proposal'])}
                style = next((node.find(W+'rPr') for node in nodes
                              if node.find(W+'rPr') is not None), None)
                for node in nodes:
                    parent.remove(node)
                if spec['anchor']:
                    run = _node('r')
                    if style is not None:
                        run.append(copy.deepcopy(style))
                    text = _node('t'); text.set(XMLSPACE, 'preserve')
                    text.text = spec['anchor']; run.append(text)
                    parent.insert(parent.index(end), run)
                _restore_pending_anchor(doc, spec, start, end)
                parent.remove(start); parent.remove(end)
                continue
            # The native revision decision owns acceptance. Removing a comment
            # alone never enters this branch while ins/del remain pending.
            from review_comments import read_comments
            comments=read_comments(doc)
            cids = {e.get(W+'id') for e in nodes if e.tag == W+'commentRangeStart'
                    and _legacy_card(comments,e.get(W+'id'),spec)}
            for child in list(parent):
                if child.tag in (W+'commentRangeStart', W+'commentRangeEnd') and child.get(W+'id') in cids:
                    parent.remove(child)
                elif child.tag == W+'r' and any(t.get(W+'id') in cids for t in child.iter(W+'commentReference')):
                    parent.remove(child)
        parent.remove(start); parent.remove(end)
    return states
