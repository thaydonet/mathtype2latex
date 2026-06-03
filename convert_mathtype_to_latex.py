"""
MathType (MTEF) to LaTeX converter for Word (.docx) files.
Replaces MathType equation OLE objects with $ct$ text.

Usage: python convert_mathtype_to_latex.py [input.docx] [output.docx]
"""

import sys
import os
import zipfile
import re
import shutil
import tempfile
import posixpath
from io import BytesIO

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import olefile
from mtef_py.ole_util.helper import Helper
from mtef_py.mtef import MTEF


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
O_NS = "urn:schemas-microsoft-com:office:office"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def resolve_docx_relationship_target(target, base_dir="word"):
    """Resolve a relationship target to the normalized path used inside a DOCX zip."""
    if not target:
        return ""
    normalized = target.replace("\\", "/")
    if normalized.startswith("/"):
        normalized = normalized.lstrip("/")
    elif not normalized.startswith("word/"):
        normalized = posixpath.normpath(posixpath.join(base_dir, normalized))
    else:
        normalized = posixpath.normpath(normalized)
    return normalized


def candidate_docx_target_paths(target):
    candidates = [target]
    if target.startswith("embeddings/"):
        candidates.append(f"word/{target}")
    if "/embeddings/" in target:
        candidates.append(f"word/embeddings/{posixpath.basename(target)}")
    return list(dict.fromkeys(candidates))


def make_latex_run(etree, source_run, latex):
    rpr = source_run.find(f"{{{W_NS}}}rPr")
    if rpr is None:
        rpr = etree.Element(f"{{{W_NS}}}rPr")

    # Remove w:vertAlign (subscript / superscript) so the LaTeX text
    # is not shrunken or raised/lowered relative to the baseline.
    vert_align = rpr.find(f"{{{W_NS}}}vertAlign")
    if vert_align is not None:
        rpr.remove(vert_align)

    # Add or update w:color tag to make the formula green (008000)
    color_tag = rpr.find(f"{{{W_NS}}}color")
    if color_tag is None:
        color_tag = etree.Element(f"{{{W_NS}}}color")
        rpr.append(color_tag)
    color_tag.set(f"{{{W_NS}}}val", "008000")

    source_run.clear()
    source_run.append(rpr)

    t_elem = etree.Element(f"{{{W_NS}}}t")
    t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t_elem.text = f"${latex}$"
    source_run.append(t_elem)


def clean_latex_formula(latex):
    if not latex:
        return latex

    # 1. Clean redundant / duplicate begin/end arrays or matrices (e.g. nested array systems)
    # \begin{array}{l} \begin{array}{} x=-4+2t \\ y=9-3t \\ \end{array} \end{array}
    latex = re.sub(
        r"\\begin\{array\}\s*\{[a-zA-Z]*\}\s*\\begin\{array\}\s*\{[a-zA-Z]*\}",
        r"\\begin{array}{l}",
        latex,
    )
    latex = re.sub(r"\\end\{array\}\s*\\end\{array\}", r"\\end{array}", latex)

    # 2. Remove extra spaces inside parenthesis (), braces {}, and after/before \left, \right
    latex = re.sub(r"\\left\s+", r"\\left", latex)
    latex = re.sub(r"\\right\s+", r"\\right", latex)

    latex = re.sub(r"\(\s+", "(", latex)
    latex = re.sub(r"\{\s+", "{", latex)

    latex = re.sub(r"\s+\)", ")", latex)
    latex = re.sub(r"\s+\}", "}", latex)

    # 3. Collapse multiple spaces to a single space
    latex = re.sub(r"\s+", " ", latex)

    # 4. Strip spaces directly inside curly braces or parentheses recursively
    # e.g., "{ u_{ n } }" -> "{u_{n}}"
    for _ in range(3):
        latex = re.sub(r"\{\s*([^{}\s]+)\s*\}", r"{\1}", latex)
        latex = re.sub(r"\(\s*([^()\s]+)\s*\)", r"(\1)", latex)
        # Handle cases with subscripts: e.g. "u_{ n }" -> "u_{n}"
        latex = re.sub(r"_\s*\{\s*([^{}\s]+)\s*\}", r"_{\1}", latex)
        latex = re.sub(r"\^\s*\{\s*([^{}\s]+)\s*\}", r"^{\1}", latex)

    return latex.strip()


def convert_ole_to_latex(ole_bin_data):
    tmp = os.path.join(tempfile.gettempdir(), "_mtef_tmp.bin")
    try:
        with open(tmp, "wb") as f:
            f.write(ole_bin_data)
        
        # Check if it's a valid OLE file
        if not olefile.isOleFile(tmp):
            print(f"WARNING: Not a valid OLE file (size: {len(ole_bin_data)} bytes)")
            return None
            
        ole = olefile.OleFileIO(tmp)
        
        # Check if "Equation Native" stream exists
        if not ole.exists("Equation Native"):
            print(f"WARNING: No 'Equation Native' stream found")
            # Try to list available streams for debugging
            try:
                streams = ole.listdir()
                print(f"  Available streams: {streams}")
            except:
                pass
            ole.close()
            return None
            
        stream_data = ole.openstream("Equation Native").read()
        
        # Validate header
        if len(stream_data) < 28:
            print(f"WARNING: Stream data too short ({len(stream_data)} bytes)")
            ole.close()
            return None
            
        hdr_reader = BytesIO(stream_data[:28])
        cb_hdr = Helper.bytes2int(hdr_reader.read(2))
        
        if cb_hdr is None or cb_hdr != 28:
            print(f"WARNING: Invalid header size (cb_hdr={cb_hdr}, expected 28)")
            ole.close()
            return None
            
        hdr_reader.seek(4 + 2, 1)
        cb_size = Helper.bytes2int(hdr_reader.read(4))
        
        if cb_size is None or cb_size <= 0:
            print(f"WARNING: Invalid equation body size (cb_size={cb_size})")
            ole.close()
            return None
            
        if cb_hdr + cb_size > len(stream_data):
            print(f"WARNING: Equation body extends beyond stream data")
            ole.close()
            return None
            
        eqn_body = stream_data[cb_hdr : cb_hdr + cb_size]
        eqn = MTEF()
        eqn.reader = BytesIO(eqn_body)
        eqn.readRecord()
        eqn.makeAST()
        latex = eqn.Translate()
        ole.close()
        
        if not latex:
            print(f"WARNING: Translation returned empty/None")
            return None
            
        return latex.strip()
        
    except Exception as e:
        # Log the error to help debugging
        print(f"ERROR: Failed to convert OLE object: {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        
    finally:
        try:
            os.remove(tmp)
        except:
            pass
    
    return None


def process_docx(input_path, output_path):
    with zipfile.ZipFile(input_path, "r") as zin:
        file_list = zin.namelist()
        doc_xml = zin.read("word/document.xml").decode("utf-8")
        rels_xml = (
            zin.read("word/_rels/document.xml.rels")
            if "word/_rels/document.xml.rels" in file_list
            else None
        )

    # Parse relationships: rId -> target
    rid_to_target = {}
    if rels_xml:
        from lxml import etree

        rels_tree = etree.fromstring(rels_xml)
        for rel in rels_tree:
            rid = rel.get("Id", "")
            target = rel.get("Target", "")
            rel_type = rel.get("Type", "")
            target_path = resolve_docx_relationship_target(target)
            if "oleobject" in rel_type.lower() or target.lower().endswith(".bin"):
                rid_to_target[rid] = target_path

    print(f"Relationships: {len(rid_to_target)}")

    # Load OLE objects and convert
    ole_objects = {}
    
    # Method 1: Try numbered oleObject files (oleObject1.bin, oleObject2.bin, ...)
    for i in range(1, 1000):
        fname = f"word/embeddings/oleObject{i}.bin"
        if fname in file_list:
            with zipfile.ZipFile(input_path, "r") as zin:
                ole_objects[fname] = zin.read(fname)
    
    # Method 2: Also try any .bin files in embeddings folder
    for fname in file_list:
        if fname.startswith("word/embeddings/") and fname.endswith(".bin"):
            if fname not in ole_objects:  # Don't duplicate
                with zipfile.ZipFile(input_path, "r") as zin:
                    ole_objects[fname] = zin.read(fname)
    
    print(f"Found {len(ole_objects)} OLE objects")

    latex_map = {}
    conversion_errors = 0
    for fname, ole_data in ole_objects.items():
        latex = convert_ole_to_latex(ole_data)
        if latex:
            latex_map[fname] = latex.strip()
        else:
            conversion_errors += 1

    print(f"Converted: {len(latex_map)}/{len(ole_objects)}")
    if conversion_errors > 0:
        print(f"WARNING: Failed to convert {conversion_errors} objects")

    # Build rId -> LaTeX
    rid_to_latex = {}
    for rid, target in rid_to_target.items():
        for candidate in candidate_docx_target_paths(target):
            if candidate in latex_map:
                rid_to_latex[rid] = latex_map[candidate]
                break

    print(f"rId->LaTeX: {len(rid_to_latex)}")

    # Strategy: Parse word/document.xml with lxml.etree, find all <w:object> elements,
    # locate their inner <o:OLEObject> with a matching rId, and replace the enclosing
    # <w:r> run with a simple run containing the LaTeX formula, preserving formatting <w:rPr>.
    from lxml import etree

    doc_xml_bytes = doc_xml.encode("utf-8") if isinstance(doc_xml, str) else doc_xml
    doc_tree = etree.fromstring(doc_xml_bytes)

    namespaces = {"w": W_NS, "o": O_NS, "r": R_NS}

    ole_xml_objects = doc_tree.xpath("//o:OLEObject", namespaces=namespaces)
    fallback_latex = [latex_map[fname] for fname in ole_objects if fname in latex_map]
    replaced = 0

    for xml_index, ole_obj in enumerate(ole_xml_objects):
        rid = ole_obj.get(
            f"{{{R_NS}}}id"
        )
        latex = rid_to_latex.get(rid)
        if not latex:
            # Some Word files contain MathType OLE objects but the relationship
            # target cannot be mapped cleanly. In that case the XML order usually
            # matches word/embeddings/oleObjectN.bin order, so use it as a fallback.
            latex = fallback_latex[xml_index] if xml_index < len(fallback_latex) else None
        if not latex:
            continue

        latex = clean_latex_formula(latex)

        # Find enclosing w:r
        parent_r = None
        parent = ole_obj.getparent()
        while parent is not None:
            if parent.tag == f"{{{W_NS}}}r":
                parent_r = parent
                break
            parent = parent.getparent()

        if parent_r is None:
            continue

        make_latex_run(etree, parent_r, latex)
        replaced += 1

    print(f"Replaced: {replaced}")

    # Serialize back to XML bytes
    doc_xml_bytes = etree.tostring(
        doc_tree, encoding="utf-8", xml_declaration=True, standalone=True
    )
    tmp_out = output_path + ".tmp"
    with zipfile.ZipFile(input_path, "r") as zin:
        with zipfile.ZipFile(tmp_out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                if item == "word/document.xml":
                    zout.writestr(item, doc_xml_bytes)
                else:
                    zout.writestr(item, zin.read(item))

    try:
        if os.path.exists(output_path):
            os.remove(output_path)
        shutil.move(tmp_out, output_path)
        print(f"Saved: {output_path}")
    except Exception as e:
        print(f"\n[ERROR] KhÃ´ng thá»ƒ ghi file káº¿t quáº£ vÃ o '{output_path}'.")
        print("CÃ³ thá»ƒ file nÃ y Ä‘ang Ä‘Æ°á»£c má»Ÿ báº±ng Word. Vui lÃ²ng Ä‘Ã³ng Word vÃ  thá»­ láº¡i!")
        print(f"Chi tiáº¿t lá»—i: {e}\n")
        # Try to clean up temp file if possible
        try:
            if os.path.exists(tmp_out):
                os.remove(tmp_out)
        except:
            pass
        raise


def format_mcq_answers_in_docx(input_path, output_path):
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    with zipfile.ZipFile(input_path, "r") as zin:
        doc_xml = zin.read("word/document.xml")

    from lxml import etree
    import copy

    doc_xml_bytes = doc_xml.encode("utf-8") if isinstance(doc_xml, str) else doc_xml
    doc_tree = etree.fromstring(doc_xml_bytes)

    namespaces = {"w": W}
    paragraphs = doc_tree.xpath("//w:p", namespaces=namespaces)
    modified = 0

    for p in paragraphs:
        texts = p.xpath(".//w:t", namespaces=namespaces)
        full_text = "".join(t.text or "" for t in texts)

        # Check if the paragraph starts with "CÃ¢u <digit>", option "A.", or option "C."
        # (C.-starting paragraphs are the second row in a 2-row AB/CD layout)
        tag_pattern = re.compile(
            r"\[\s*[Mm]á»©c\s+[ÄÄ‘]á»™\s*[:\-]?\s*\d+\s*\]\s*|^\s*[Mm]á»©c\s+[ÄÄ‘]á»™\s*[:\-]?\s*\d+[\s\.\-:]*"
        )
        clean_text = tag_pattern.sub("", full_text, count=1)

        is_mcq_question = re.match(r"^\s*[Cc]Ã¢u\s+\d+", clean_text)
        is_mcq_options_a = re.match(r"^\s*A\.", clean_text)
        # CD-only row: starts with C. and has NO A. earlier in the paragraph
        is_mcq_options_c = (
            re.match(r"^\s*C\.", clean_text)
            and not re.search(r"\bA\.", full_text)
        )

        # MSQ flags (using a) b) c) d) or a. b. c. d. - lowercase)
        is_msq_options_a = re.match(r"^\s*a\s*[\)\.]", clean_text)
        is_msq_options_c = (
            re.match(r"^\s*c\s*[\)\.]", clean_text)
            and not re.search(r"\ba\s*[\)\.]", full_text)
        )

        if not (is_mcq_question or is_mcq_options_a or is_mcq_options_c or is_msq_options_a or is_msq_options_c):
            continue

        # â”€â”€ Determine target_indices based on which options are present â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        target_indices = set()

        is_msq_active = False
        if is_msq_options_a or (is_mcq_question and re.search(r"\ba\s*[\)\.]", full_text)):
            is_msq_active = True
            match_a = re.search(r"\ba\s*[\)\.]", full_text)
            if match_a:
                idx_a = match_a.start()
                match_b = re.search(r"\bb\s*[\)\.]", full_text[idx_a:])
                if match_b:
                    idx_b = idx_a + match_b.start()
                    target_indices.add(idx_b)

                    match_c = re.search(r"\bc\s*[\)\.]", full_text[idx_b:])
                    if match_c:
                        idx_c = idx_b + match_c.start()
                        target_indices.add(idx_c)
                        match_d = re.search(r"\bd\s*[\)\.]", full_text[idx_c:])
                        if match_d:
                            idx_d = idx_c + match_d.start()
                            target_indices.add(idx_d)

        elif is_msq_options_c and not (is_mcq_question or is_msq_options_a):
            is_msq_active = True
            match_c = re.search(r"\bc\s*[\)\.]", full_text)
            if match_c:
                idx_c = match_c.start()
                match_d = re.search(r"\bd\s*[\)\.]", full_text[idx_c:])
                if match_d:
                    idx_d = idx_c + match_d.start()
                    target_indices.add(idx_d)

        # Fallback to standard MCQ if not MSQ
        if not is_msq_active:
            if is_mcq_options_c and not (is_mcq_question or is_mcq_options_a):
                # CD-only row: only insert a break before D.
                match_c = re.search(r"\bC\.", full_text)
                if not match_c:
                    continue
                idx_c = match_c.start()
                match_d = re.search(r"\bD\.", full_text[idx_c:])
                if not match_d:
                    continue
                idx_d = idx_c + match_d.start()
                target_indices = {idx_d}
            else:
                # A-starting paragraph (question or AB / ABCD row)
                match_a = re.search(r"\bA\.", full_text)
                if not match_a:
                    continue
                idx_a = match_a.start()

                match_b = re.search(r"\bB\.", full_text[idx_a:])
                if not match_b:
                    continue
                idx_b = idx_a + match_b.start()
                target_indices.add(idx_b)

                # C. and D. are optional â€“ may live in a separate paragraph
                match_c = re.search(r"\bC\.", full_text[idx_b:])
                if match_c:
                    idx_c = idx_b + match_c.start()
                    target_indices.add(idx_c)
                    match_d = re.search(r"\bD\.", full_text[idx_c:])
                    if match_d:
                        idx_d = idx_c + match_d.start()
                        target_indices.add(idx_d)

        if not target_indices:
            continue

        # 1. First pass: Flatten elements and calculate global character ranges for text elements
        flat_items = []
        global_idx = 0
        children = list(p)
        p_pr = None

        for child in children:
            if child.tag == f"{{{W}}}pPr":
                p_pr = child
                continue

            if child.tag != f"{{{W}}}r":
                flat_items.append({"type": "generic", "elem": child})
                continue

            r = child
            r_pr = r.find(f"{{{W}}}rPr")
            run_elements = list(r)

            for elem in run_elements:
                if elem.tag == f"{{{W}}}rPr":
                    continue
                if elem.tag == f"{{{W}}}br":
                    flat_items.append({"type": "br", "elem": elem, "r_pr": r_pr})
                elif elem.tag == f"{{{W}}}t":
                    txt = elem.text or ""
                    start = global_idx
                    end = global_idx + len(txt)
                    flat_items.append({
                        "type": "t",
                        "text": txt,
                        "r_pr": r_pr,
                        "start_idx": start,
                        "end_idx": end
                    })
                    global_idx += len(txt)
                else:
                    flat_items.append({"type": "obj", "elem": elem, "r_pr": r_pr})

        # 2. Second pass: Rebuild the paragraph's children with precise splitting
        new_children = []
        if p_pr is not None:
            new_children.append(p_pr)

        text_since_last_break = ""
        # Track whether we've seen any content at all (including across br resets)
        any_content_before = False
        inserted = False

        # Pre-compute: for each flat_items index that is a "br", check if it is
        # immediately followed (possibly after whitespace-only "t" items) by a "t"
        # item that starts at a target_index.  If so, that existing w:br will be
        # replaced by our own inserted break, so we should skip (drop) it to avoid
        # a double blank line.
        def next_text_after_br(idx):
            """Return the start_idx of the first non-empty text after flat_items[idx],
            or None if not found before a non-whitespace item."""
            j = idx + 1
            while j < len(flat_items):
                ni = flat_items[j]
                if ni["type"] == "t":
                    stripped = (ni["text"] or "").lstrip(" \t")
                    if stripped:
                        return ni["start_idx"] + (len(ni["text"]) - len(stripped))
                    j += 1
                else:
                    break
            return None

        skip_br_indices = set()
        for fi, fitem in enumerate(flat_items):
            if fitem["type"] == "br":
                nx = next_text_after_br(fi)
                if nx is not None and nx in target_indices:
                    skip_br_indices.add(fi)

        for fi, item in enumerate(flat_items):
            itype = item["type"]
            if itype == "generic":
                new_children.append(item["elem"])
            elif itype == "br":
                if fi in skip_br_indices:
                    # Drop this existing br â€“ we will insert our own break before
                    # the next target option, avoiding a double line break.
                    if text_since_last_break.strip():
                        any_content_before = True
                    text_since_last_break = ""
                else:
                    # Preserve existing line break; reset tracker but mark that content existed
                    if text_since_last_break.strip():
                        any_content_before = True
                    text_since_last_break = ""
                    new_run = etree.Element(f"{{{W}}}r")
                    if item["r_pr"] is not None:
                        new_run.append(copy.deepcopy(item["r_pr"]))
                    new_run.append(item["elem"])
                    new_children.append(new_run)
            elif itype == "obj":
                new_run = etree.Element(f"{{{W}}}r")
                if item["r_pr"] is not None:
                    new_run.append(copy.deepcopy(item["r_pr"]))
                new_run.append(item["elem"])
                new_children.append(new_run)
                text_since_last_break += " "
                any_content_before = True
            elif itype == "t":
                txt = item["text"]
                r_pr = item["r_pr"]
                start_idx = item["start_idx"]
                end_idx = item["end_idx"]

                targets_in_item = sorted([g for g in target_indices if start_idx <= g < end_idx])

                if not targets_in_item:
                    # Strip leading whitespace that immediately precedes a known option
                    # (this handles tabs in a run that doesn't contain the option itself)
                    txt_out = txt
                    if not text_since_last_break and txt_out.startswith((" ", "\t")):
                        txt_out = txt_out.lstrip(" \t")
                    if txt_out:
                        new_r = etree.Element(f"{{{W}}}r")
                        if r_pr is not None:
                            new_r.append(copy.deepcopy(r_pr))
                        t_elem = etree.SubElement(new_r, f"{{{W}}}t")
                        t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                        t_elem.text = txt_out
                        new_children.append(new_r)
                    text_since_last_break += txt_out
                    if txt_out.strip():
                        any_content_before = True
                else:
                    local_indices = [g - start_idx for g in targets_in_item]

                    parts = []
                    last_pos = 0
                    for pos in local_indices:
                        parts.append(txt[last_pos:pos])
                        last_pos = pos
                    parts.append(txt[last_pos:])

                    first_part = parts[0]
                    # Strip trailing whitespace/tab that would appear before the line break
                    first_part_clean = first_part.rstrip(" \t") if first_part else ""
                    if first_part_clean:
                        new_r = etree.Element(f"{{{W}}}r")
                        if r_pr is not None:
                            new_r.append(copy.deepcopy(r_pr))
                        t_elem = etree.SubElement(new_r, f"{{{W}}}t")
                        t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                        t_elem.text = first_part_clean
                        new_children.append(new_r)
                        text_since_last_break += first_part_clean
                        any_content_before = True

                    for part in parts[1:]:
                        if not part:
                            continue
                        # Insert a line break if there is content on the current line
                        # OR if there was any content before (handles the case where
                        # an existing w:br was dropped and any_content_before is True)
                        if len(text_since_last_break.strip()) > 0 or any_content_before:
                            br_run = etree.Element(f"{{{W}}}r")
                            if r_pr is not None:
                                br_run.append(copy.deepcopy(r_pr))
                            etree.SubElement(br_run, f"{{{W}}}br")
                            new_children.append(br_run)
                            text_since_last_break = ""
                            any_content_before = False
                            inserted = True

                        # Strip leading whitespace/tab from the option part
                        part_clean = part.lstrip(" \t")
                        if not part_clean:
                            continue
                        new_r = etree.Element(f"{{{W}}}r")
                        if r_pr is not None:
                            new_r.append(copy.deepcopy(r_pr))
                        t_elem = etree.SubElement(new_r, f"{{{W}}}t")
                        t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                        t_elem.text = part_clean
                        new_children.append(new_r)
                        text_since_last_break += part_clean
                        any_content_before = True

        if inserted:
            p.clear()
            for child in new_children:
                p.append(child)
            modified += 1

    print(f"Formatted {modified} MCQ paragraph(s)")

    doc_xml_bytes = etree.tostring(
        doc_tree, encoding="utf-8", xml_declaration=True, standalone=True
    )
    tmp_out = output_path + ".tmp"
    with zipfile.ZipFile(input_path, "r") as zin:
        with zipfile.ZipFile(tmp_out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                if item == "word/document.xml":
                    zout.writestr(item, doc_xml_bytes)
                else:
                    zout.writestr(item, zin.read(item))

    try:
        if os.path.exists(output_path):
            os.remove(output_path)
        shutil.move(tmp_out, output_path)
        print(f"Saved: {output_path}")
    except Exception as e:
        print(f"\n[ERROR] Cannot write to '{output_path}'.")
        print(f"Error: {e}\n")
        try:
            if os.path.exists(tmp_out):
                os.remove(tmp_out)
        except:
            pass
        raise


def remove_paragraphs_containing(input_path, output_path, keyword):
    with zipfile.ZipFile(input_path, "r") as zin:
        doc_xml = zin.read("word/document.xml")

    from lxml import etree

    doc_xml_bytes = doc_xml.encode("utf-8") if isinstance(doc_xml, str) else doc_xml
    doc_tree = etree.fromstring(doc_xml_bytes)

    namespaces = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    }

    paragraphs = doc_tree.xpath("//w:p", namespaces=namespaces)
    removed = 0

    for p in paragraphs:
        texts = p.xpath(".//w:t", namespaces=namespaces)
        full_text = "".join(t.text or "" for t in texts)
        if keyword.lower() in full_text.lower():
            p.getparent().remove(p)
            removed += 1

    print(f"Removed {removed} paragraph(s) containing '{keyword}'")

    doc_xml_bytes = etree.tostring(
        doc_tree, encoding="utf-8", xml_declaration=True, standalone=True
    )
    tmp_out = output_path + ".tmp"
    with zipfile.ZipFile(input_path, "r") as zin:
        with zipfile.ZipFile(tmp_out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                if item == "word/document.xml":
                    zout.writestr(item, doc_xml_bytes)
                else:
                    zout.writestr(item, zin.read(item))

    try:
        if os.path.exists(output_path):
            os.remove(output_path)
        shutil.move(tmp_out, output_path)
        print(f"Saved: {output_path}")
    except Exception as e:
        print(f"\n[ERROR] Cannot write to '{output_path}'.")
        print(f"Error: {e}\n")
        try:
            if os.path.exists(tmp_out):
                os.remove(tmp_out)
        except:
            pass
        raise


def remove_level_tags_in_docx(input_path, output_path):
    """
    Remove exactly these 4 strings from document:
    [Mức độ 1], [Mức độ 2], [Mức độ 3], [Mức độ 4]
    Example: "Câu 1. [Mức độ 1] hàm số..." -> "Câu 1. hàm số..."
    """
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    with zipfile.ZipFile(input_path, "r") as zin:
        doc_xml = zin.read("word/document.xml")

    from lxml import etree
    import copy

    doc_xml_bytes = doc_xml.encode("utf-8") if isinstance(doc_xml, str) else doc_xml
    doc_tree = etree.fromstring(doc_xml_bytes)

    namespaces = {"w": W}
    paragraphs = doc_tree.xpath("//w:p", namespaces=namespaces)
    modified = 0

    # Exact strings to remove (with optional spaces)
    patterns_to_remove = [
        "[Mức độ 1]",
        "[ Mức độ 1]",
        "[Mức độ 1 ]",
        "[ Mức độ 1 ]",
        "[Mức độ 2]",
        "[ Mức độ 2]",
        "[Mức độ 2 ]",
        "[ Mức độ 2 ]",
        "[Mức độ 3]",
        "[ Mức độ 3]",
        "[Mức độ 3 ]",
        "[ Mức độ 3 ]",
        "[Mức độ 4]",
        "[ Mức độ 4]",
        "[Mức độ 4 ]",
        "[ Mức độ 4 ]",
    ]

    for p in paragraphs:
        texts = p.xpath(".//w:t", namespaces=namespaces)
        full_text = "".join(t.text or "" for t in texts)

        # Check if any pattern exists in this paragraph
        found_pattern = False
        for pattern in patterns_to_remove:
            if pattern in full_text:
                found_pattern = True
                break
        
        if not found_pattern:
            continue

        # Build index set of characters to remove
        remove_indices = set()
        for pattern in patterns_to_remove:
            idx = 0
            while idx < len(full_text):
                pos = full_text.find(pattern, idx)
                if pos == -1:
                    break
                # Mark all characters in this pattern for removal
                for i in range(pos, pos + len(pattern)):
                    remove_indices.add(i)
                idx = pos + 1

        if not remove_indices:
            continue

        children = list(p)
        p_pr = None
        for child in children:
            if child.tag == f"{{{W}}}pPr":
                p_pr = child
                break

        flat_items = []
        global_idx = 0
        for child in children:
            if child.tag == f"{{{W}}}pPr":
                continue
            if child.tag != f"{{{W}}}r":
                flat_items.append({"type": "generic", "elem": child})
                continue

            r = child
            r_pr = r.find(f"{{{W}}}rPr")
            run_elements = list(r)

            for elem in run_elements:
                if elem.tag == f"{{{W}}}rPr":
                    continue
                if elem.tag == f"{{{W}}}br":
                    flat_items.append({"type": "br", "elem": elem, "r_pr": r_pr})
                elif elem.tag == f"{{{W}}}t":
                    txt = elem.text or ""
                    start = global_idx
                    end = global_idx + len(txt)
                    flat_items.append({
                        "type": "t",
                        "text": txt,
                        "r_pr": r_pr,
                        "start_idx": start,
                        "end_idx": end
                    })
                    global_idx += len(txt)
                else:
                    flat_items.append({"type": "obj", "elem": elem, "r_pr": r_pr})

        new_children = []
        if p_pr is not None:
            new_children.append(p_pr)

        for item in flat_items:
            itype = item["type"]
            if itype == "generic":
                new_children.append(item["elem"])
            elif itype == "br" or itype == "obj":
                new_run = etree.Element(f"{{{W}}}r")
                if item["r_pr"] is not None:
                    new_run.append(copy.deepcopy(item["r_pr"]))
                new_run.append(item["elem"])
                new_children.append(new_run)
            elif itype == "t":
                txt = item["text"]
                r_pr = item["r_pr"]
                start_idx = item["start_idx"]
                end_idx = item["end_idx"]

                new_txt_chars = []
                for idx in range(start_idx, end_idx):
                    if idx not in remove_indices:
                        new_txt_chars.append(txt[idx - start_idx])
                new_txt = "".join(new_txt_chars)

                if new_txt:
                    new_r = etree.Element(f"{{{W}}}r")
                    if r_pr is not None:
                        new_r.append(copy.deepcopy(r_pr))
                    t_elem = etree.SubElement(new_r, f"{{{W}}}t")
                    t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    t_elem.text = new_txt
                    new_children.append(new_r)

        p.clear()
        for child in new_children:
            p.append(child)
        modified += 1

    print(f"Removed [Mức độ 1,2,3,4] from {modified} paragraph(s)")

    doc_xml_bytes = etree.tostring(
        doc_tree, encoding="utf-8", xml_declaration=True, standalone=True
    )
    tmp_out = output_path + ".tmp"
    with zipfile.ZipFile(input_path, "r") as zin:
        with zipfile.ZipFile(tmp_out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                if item == "word/document.xml":
                    zout.writestr(item, doc_xml_bytes)
                else:
                    zout.writestr(item, zin.read(item))

    try:
        if os.path.exists(output_path):
            os.remove(output_path)
        shutil.move(tmp_out, output_path)
        print(f"Saved: {output_path}")
    except Exception as e:
        print(f"\n[ERROR] Cannot write to '{output_path}'.")
        print(f"Error: {e}\n")
        try:
            if os.path.exists(tmp_out):
                os.remove(tmp_out)
        except:
            pass
        raise



def convert_tables_to_text(input_path, output_path):
    """
    Convert all tables in a .docx file to plain text paragraphs.
    Each cell's paragraphs are extracted in reading order (leftâ†’right, topâ†’bottom)
    and inserted at the position where the table was. The table element is then removed.
    """
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    with zipfile.ZipFile(input_path, "r") as zin:
        doc_xml = zin.read("word/document.xml")

    from lxml import etree
    import copy

    doc_xml_bytes = doc_xml.encode("utf-8") if isinstance(doc_xml, str) else doc_xml
    doc_tree = etree.fromstring(doc_xml_bytes)

    namespaces = {"w": W}

    # Find all tables; process in reverse so index shifting doesn't matter
    tables = doc_tree.xpath("//w:tbl", namespaces=namespaces)
    converted = 0

    for tbl in tables:
        parent = tbl.getparent()
        if parent is None:
            continue

        insert_pos = list(parent).index(tbl)

        # Collect all paragraphs from every cell, in order
        extracted_paragraphs = []
        rows = tbl.xpath(".//w:tr", namespaces=namespaces)
        for row in rows:
            cells = row.xpath(".//w:tc", namespaces=namespaces)
            for cell in cells:
                cell_paragraphs = cell.xpath(".//w:p", namespaces=namespaces)
                for p in cell_paragraphs:
                    extracted_paragraphs.append(copy.deepcopy(p))

        # Remove the table
        parent.remove(tbl)

        # Insert extracted paragraphs at the table's former position
        for i, p in enumerate(extracted_paragraphs):
            parent.insert(insert_pos + i, p)

        converted += 1

    print(f"Converted {converted} table(s) to text")

    doc_xml_bytes = etree.tostring(
        doc_tree, encoding="utf-8", xml_declaration=True, standalone=True
    )
    tmp_out = output_path + ".tmp"
    with zipfile.ZipFile(input_path, "r") as zin:
        with zipfile.ZipFile(tmp_out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                if item == "word/document.xml":
                    zout.writestr(item, doc_xml_bytes)
                else:
                    zout.writestr(item, zin.read(item))

    try:
        if os.path.exists(output_path):
            os.remove(output_path)
        shutil.move(tmp_out, output_path)
        print(f"Saved: {output_path}")
    except Exception as e:
        print(f"\n[ERROR] Cannot write to '{output_path}'.")
        print(f"Error: {e}\n")
        try:
            if os.path.exists(tmp_out):
                os.remove(tmp_out)
        except:
            pass
        raise


def format_all_text_times_new_roman_12(input_path, output_path):
    """
    Format all text in the document with Times New Roman font, size 12pt.
    This applies to all runs (w:r) in the document.
    """
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    with zipfile.ZipFile(input_path, "r") as zin:
        doc_xml = zin.read("word/document.xml")

    from lxml import etree
    import copy

    doc_xml_bytes = doc_xml.encode("utf-8") if isinstance(doc_xml, str) else doc_xml
    doc_tree = etree.fromstring(doc_xml_bytes)

    namespaces = {"w": W}

    # Find all runs (w:r) in the document
    runs = doc_tree.xpath("//w:r", namespaces=namespaces)
    modified = 0

    for run in runs:
        # Get or create w:rPr (run properties)
        rpr = run.find(f"{{{W}}}rPr")
        if rpr is None:
            rpr = etree.Element(f"{{{W}}}rPr")
            # Insert at beginning of run
            run.insert(0, rpr)

        # Set font to Times New Roman
        rfonts = rpr.find(f"{{{W}}}rFonts")
        if rfonts is None:
            rfonts = etree.Element(f"{{{W}}}rFonts")
            rpr.append(rfonts)
        rfonts.set(f"{{{W}}}ascii", "Times New Roman")
        rfonts.set(f"{{{W}}}hAnsi", "Times New Roman")
        rfonts.set(f"{{{W}}}cs", "Times New Roman")
        rfonts.set(f"{{{W}}}eastAsia", "Times New Roman")

        # Set font size to 12pt (24 half-points)
        sz = rpr.find(f"{{{W}}}sz")
        if sz is None:
            sz = etree.Element(f"{{{W}}}sz")
            rpr.append(sz)
        sz.set(f"{{{W}}}val", "24")

        # Set font size for complex scripts to 12pt (24 half-points)
        sz_cs = rpr.find(f"{{{W}}}szCs")
        if sz_cs is None:
            sz_cs = etree.Element(f"{{{W}}}szCs")
            rpr.append(sz_cs)
        sz_cs.set(f"{{{W}}}val", "24")

        modified += 1

    print(f"Formatted {modified} text run(s) with Times New Roman 12pt")

    doc_xml_bytes = etree.tostring(
        doc_tree, encoding="utf-8", xml_declaration=True, standalone=True
    )
    tmp_out = output_path + ".tmp"
    with zipfile.ZipFile(input_path, "r") as zin:
        with zipfile.ZipFile(tmp_out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                if item == "word/document.xml":
                    zout.writestr(item, doc_xml_bytes)
                else:
                    zout.writestr(item, zin.read(item))

    try:
        if os.path.exists(output_path):
            os.remove(output_path)
        shutil.move(tmp_out, output_path)
        print(f"Saved: {output_path}")
    except Exception as e:
        print(f"\n[ERROR] Cannot write to '{output_path}'.")
        print(f"Error: {e}\n")
        try:
            if os.path.exists(tmp_out):
                os.remove(tmp_out)
        except:
            pass
        raise


def extract_underlined_answer(docx_path):
    """
    Extract the correct answer from underlined text in Word document.
    Looks for underlined single letters: A, B, C, D (MCQ) or a, b, c, d (MSQ)
    
    Returns:
        dict: Maps question number to correct answer (e.g., {"1": "A", "2": "B,C"})
    """
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    
    with zipfile.ZipFile(docx_path, "r") as zin:
        doc_xml = zin.read("word/document.xml")
    
    from lxml import etree
    
    doc_xml_bytes = doc_xml.encode("utf-8") if isinstance(doc_xml, str) else doc_xml
    doc_tree = etree.fromstring(doc_xml_bytes)
    
    namespaces = {"w": W}
    paragraphs = doc_tree.xpath("//w:p", namespaces=namespaces)
    
    answer_map = {}
    current_question_num = None
    
    for p in paragraphs:
        # Get full paragraph text to identify question numbers
        texts = p.xpath(".//w:t", namespaces=namespaces)
        full_text = "".join(t.text or "" for t in texts)
        
        # Check if this paragraph starts with "Câu <number>"
        question_match = re.match(r"^\s*[Cc]âu\s+(\d+)", full_text)
        if question_match:
            current_question_num = question_match.group(1)
        
        # Find all runs (w:r) in this paragraph
        runs = p.xpath(".//w:r", namespaces=namespaces)
        
        for run in runs:
            # Check if this run has underline formatting
            rpr = run.find(f"{{{W}}}rPr")
            if rpr is None:
                continue
            
            underline = rpr.find(f"{{{W}}}u")
            if underline is None:
                continue
            
            # This run is underlined - extract text
            t_elems = run.xpath(".//w:t", namespaces=namespaces)
            underlined_text = "".join(t.text or "" for t in t_elems).strip()
            
            if not underlined_text:
                continue
            
            # Match single letter answers: A, B, C, D (uppercase) or a, b, c, d (lowercase)
            # Allow for variations like "A.", "A)", or just "A"
            # Pattern: letter followed by optional delimiter (. or )) or whitespace or end
            answer_pattern = re.findall(r'\b([A-Da-d])(?=[\.\),\s]|$)', underlined_text)
            
            if answer_pattern and current_question_num:
                # Normalize to uppercase
                answers = [a.upper() for a in answer_pattern]
                
                # For MCQ: single answer (A, B, C, or D)
                # For MSQ: multiple answers (a,b or a,c,d) - combine with comma
                if current_question_num not in answer_map:
                    answer_map[current_question_num] = []
                
                answer_map[current_question_num].extend(answers)
    
    # Deduplicate and join multiple answers with comma
    result = {}
    for qnum, answers in answer_map.items():
        # Remove duplicates while preserving order
        seen = set()
        unique_answers = []
        for a in answers:
            if a not in seen:
                seen.add(a)
                unique_answers.append(a)
        result[qnum] = ','.join(unique_answers)
    
    return result


# Import upload images function
def upload_images_to_cloudinary(input_path, output_path):
    """
    Wrapper function to upload images to Cloudinary
    """
    try:
        from upload_images_to_cloudinary import upload_images_in_docx
        upload_images_in_docx(input_path, output_path)
    except ImportError as e:
        print(f"ERROR: Could not import upload module: {e}")
        raise
    except Exception as e:
        print(f"ERROR uploading images: {e}")
        raise


if __name__ == "__main__":

    input_file = r"D:\code-latex\mau-doc.docx"
    output_file = r"D:\code-latex\mau-doc-latex.docx"

    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    if len(sys.argv) > 2:
        output_file = sys.argv[2]

    print(f"Input:  {input_file}")
    print(f"Output: {output_file}")
    print("=" * 60)
    process_docx(input_file, output_file)
