#!/usr/bin/env python3
"""
extract_docx_pii.py

Usage:
    python extract_docx_pii.py path/to/document.docx

Output: prints JSON to stdout (list of entities)
"""

import sys
import zipfile
import re
import json
from lxml import etree

# optional spacy
USE_SPACY = False
try:
    import spacy
    nlp = spacy.load("en_core_web_sm")
    USE_SPACY = True
except Exception:
    USE_SPACY = False

# regexes for common PII
RE_EMAIL = re.compile(r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+')
RE_PHONE = re.compile(r'(\+?\d{1,3}[\s\-\u2011]?)?(?:\(?\d{2,4}\)?[\s\-\/\.]?)?\d{2,4}[\s\-\/\.]?\d{2,4}(?:[\s\-\/\.]?\d{1,4})?')
RE_IBAN = re.compile(r'\b[A-Z]{2}[0-9]{2}[A-Z0-9]{4,30}\b')
RE_DATE = re.compile(r'\b(?:\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}|\d{4}[\/\-.]\d{1,2}[\/\-.]\d{1,2}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2},?\s+\d{4})\b', re.I)
RE_ADDRESS_HINT = re.compile(r'\b(street|st\.|road|rd\.|ave|avenue|blvd|lane|ln\.|weg|strasse|straße|platz|zip|postal|postfach|pf)\b', re.I)

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


def get_text_nodes_and_paths(xml_root):
    """Yield (elem, xpath_like) for all <w:t> text elements in document.xml."""
    for t in xml_root.findall('.//w:t', namespaces=NS):
        # Build XPath-like path
        path_parts = []
        el = t
        while el is not None and el.getparent() is not None:
            parent = el.getparent()
            tag = etree.QName(el).localname
            same = [s for s in parent if etree.QName(s).localname == tag]
            idx = same.index(el) + 1
            path_parts.append(f"w:{tag}[{idx}]")
            el = parent
        path_parts.reverse()
        xpath_like = "/word/document.xml/" + "/".join(path_parts)
        yield t, xpath_like


def find_entities_in_text(text):
    """Return list of (entity_name, match_text, start, end)."""
    found = []
    for m in RE_EMAIL.finditer(text):
        found.append(('email', m.group(0), m.start(), m.end()))
    for m in RE_IBAN.finditer(text):
        found.append(('account', m.group(0), m.start(), m.end()))
    for m in RE_DATE.finditer(text):
        found.append(('date', m.group(0), m.start(), m.end()))
    for m in RE_PHONE.finditer(text):
        s = m.group(0)
        digits = re.sub(r'\D', '', s)
        if 6 <= len(digits) <= 15 and not RE_EMAIL.match(s):
            found.append(('phone', s, m.start(), m.end()))
    for m in RE_ADDRESS_HINT.finditer(text):
        start = max(0, m.start()-30)
        end = min(len(text), m.end()+60)
        cand = text[start:end].strip()
        found.append(('address_hint', cand, start, end))
    if USE_SPACY:
        doc = nlp(text)
        for ent in doc.ents:
            if ent.label_ in ('PERSON', 'GPE', 'LOC', 'ORG'):
                label = 'name' if ent.label_ == 'PERSON' else ('location' if ent.label_ in ('GPE', 'LOC') else 'organization')
                found.append((label, ent.text, ent.start_char, ent.end_char))
    else:
        for m in re.finditer(r'\b([A-ZÄÖÜ][a-zäöüß]+(?:\s+[A-ZÄÖÜ][a-zäöüß]+){1,3})\b', text):
            found.append(('name_heuristic', m.group(0), m.start(), m.end()))
    return found


def main(docx_path):
    with zipfile.ZipFile(docx_path) as z:
        try:
            xml_bytes = z.read('word/document.xml')
        except KeyError:
            sys.exit("Error: word/document.xml not found in docx")

    parser = etree.XMLParser(ns_clean=True, recover=True)
    root = etree.fromstring(xml_bytes, parser=parser)

    results = []
    for elem, xpath_like in get_text_nodes_and_paths(root):
        text = ''.join(elem.itertext())
        if not text:
            continue
        ents = find_entities_in_text(text)
        for ent_name, ent_text, start, end in ents:
            item = {
                "entity_name": ent_name,
                "text": ent_text,
                "xml_path": xpath_like,
                "offset": int(start),
                "length": len(ent_text)   # ← Added field: text length
            }
            results.append(item)

    # Deduplicate
    seen = set()
    unique = []
    for r in results:
        key = (r['xml_path'], r['offset'], r['text'])
        if key not in seen:
            seen.add(key)
            unique.append(r)

    print(json.dumps(unique, indent=2, ensure_ascii=False))


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python extract_docx_pii.py path/to/document.docx")
        sys.exit(1)
    main(sys.argv[1])



import zipfile
from lxml import etree

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

def redact_docx(docx_path, output_path, redactions):
    """
    Redact text across single or multiple <w:t> nodes.

    redactions: list of dicts, each with keys:
      - xml_path: full path like "/word/document.xml/w:body/w:p[14]/w:r[1]/w:t[1]"
      - offset: character offset where replacement starts (int)
      - length: number of characters to replace (int)
    """
    with zipfile.ZipFile(docx_path, 'r') as zin:
        xml_bytes = zin.read('word/document.xml')
        xml = etree.fromstring(xml_bytes)

        # Gather all text nodes in linear order for multi-node search
        text_nodes = list(xml.findall('.//w:t', namespaces=NS))

        for red in redactions:
            xml_path = red['xml_path']
            offset = red['offset']
            length = red['length']

            # Find the target element
            local_path = xml_path.replace('/word/document.xml', '')
            nodes = xml.xpath(local_path, namespaces=NS)
            if not nodes:
                continue
            start_node = nodes[0]

            # Step 1: find index of this node in overall sequence
            try:
                idx = text_nodes.index(start_node)
            except ValueError:
                continue

            remaining = length
            cur_offset = offset

            # Step 2: redact from start node onwards
            while remaining > 0 and idx < len(text_nodes):
                el = text_nodes[idx]
                text = el.text or ""

                # if starting node, skip to offset; else start at 0
                start_pos = cur_offset if idx == text_nodes.index(start_node) else 0
                end_pos = min(len(text), start_pos + remaining)
                if start_pos < len(text):
                    span_len = end_pos - start_pos
                    replacement = "█" * span_len
                    new_text = text[:start_pos] + replacement + text[end_pos:]
                    el.text = new_text
                    remaining -= span_len
                else:
                    # nothing to redact in this node
                    pass

                # after first node, reset offset to 0 for next
                cur_offset = 0
                idx += 1

        # Save back to DOCX
        updated_xml = etree.tostring(xml, xml_declaration=True, encoding='UTF-8')
        with zipfile.ZipFile(output_path, 'w') as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'word/document.xml':
                    data = updated_xml
                zout.writestr(item, data)

    print(f"✅ Redacted file written to: {output_path}")
