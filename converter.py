"""
Tamil PDF to Word (.docx) High-Fidelity Converter
==================================================
Converts legacy TAM / TAB / Elango encoded Tamil PDFs into clean,
Unicode-standard, editable Word documents (.docx) without losing
any Tamil words, vowel signs, conjuncts, stanzas, or alignments.
"""

import os
import sys
import re
import pymupdf
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Open-Tamil base mapping
import tamil.txt2unicode as t2u

# 1. Base TAM to Unicode Dictionary with All Necessary Fixes
base_tam = dict(t2u.tam2unicode.__globals__['tam2utf8'])

# Crucial fixes for missing/mismapped characters in standard tables
base_tam['ƒ'] = 'ங்'    # Omitted in standard open-tamil (mapped to quote)
base_tam['ª÷£'] = 'ளொ'  # Mismapped to ள in standard open-tamil
base_tam['ªë÷'] = 'ஞௌ'
base_tam['ú§'] = 'ஸு'
base_tam['ú¨'] = 'ஸூ'
base_tam['û§'] = 'ஷு'
base_tam['û¨'] = 'ஷூ'
base_tam['ý§'] = 'ஹு'
base_tam['ý¨'] = 'ஹூ'
base_tam['ñ¢'] = 'ம்'
base_tam['¢'] = '்'     # Standalone virama / pulli
base_tam['£'] = 'ா'     # Standalone kaal
base_tam['¬'] = 'ை'     # Standalone ai-kombu

# Ornamental bullet characters (Wingdings / typesetting fleurons)
ORNAMENT_MAP = {
    '\uf0b2': '❖',
    '\uf077': '❖',
    '': '❖',
    '': '❖',
    '': '❖',
}

# Rules sorted by key length descending so longest ligature matches first
SORTED_TAM_RULES = sorted(base_tam.items(), key=lambda x: -len(x[0]))

LATIN_FONTS = {'times', 'minion', 'myriad', 'arial', 'calibri', 'helvetica'}

def is_latin_font(font_name):
    fn = font_name.lower()
    return any(lf in fn for lf in LATIN_FONTS)

def is_symbol_font(font_name):
    fn = font_name.lower()
    return 'wingdings' in fn or 'symbol' in fn

def sanitize_xml(text):
    """Ensure text contains only valid XML 1.0 characters."""
    # Convert typesetting bullet glyph code to clean bullet
    text = text.replace('\x01', ' • ')
    # Strip any null bytes or invalid control characters
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', text)
    return text

def preprocess_spaced_tam(text):
    """Collapse accidental spacing introduced during justification."""
    # Kombu followed by space then consonant
    text = re.sub(r'([ª«¬])\s+([^\s])', r'\1\2', text)
    # Consonant followed by space then kaal or pulli
    text = re.sub(r'([^\s])\s+([£¢÷])', r'\1\2', text)
    return text

def convert_tam_to_unicode(text):
    """Converts a TAM-encoded string into pure Tamil Unicode."""
    text = preprocess_spaced_tam(text)
    for k, v in SORTED_TAM_RULES:
        text = text.replace(k, v)
    for k, v in ORNAMENT_MAP.items():
        text = text.replace(k, v)
    return sanitize_xml(text)

def set_font_run(run, font_name="Nirmala UI", size_pt=12, bold=False, italic=False, color_rgb=None):
    """Sets font attributes for Word, including complex script (w:cs) for Tamil Unicode."""
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    if color_rgb:
        run.font.color.rgb = RGBColor(*color_rgb)
    
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:cs'), font_name)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)

def is_running_header_footer(line_text, y0, y1, page_height, book_title="", author=""):
    """Detects and filters repeated running headers or footers from body text."""
    clean = line_text.strip()
    clean_no_punct = re.sub(r'[❖\[\]\s\d\.\-]', '', clean)
    
    # Bottom margin footer check (in bottom 12% of page)
    if y0 > page_height * 0.88 or y1 > page_height * 0.88:
        if re.match(r'^[\d\[\]\s❖\.\-]+$', clean):
            return True
        if book_title and book_title in clean:
            return True
        if author and author in clean:
            return True
            
    # Top margin header check (in top 8% of page)
    if y1 < page_height * 0.08:
        if re.match(r'^[\d\[\]\s❖\.\-]+$', clean):
            return True
        if book_title and book_title in clean:
            return True
        if author and author in clean:
            return True
            
    return False

def extract_page_lines(page, clip_rect=None):
    """Extract ordered lines with font info, coordinates, and converted text."""
    blocks = page.get_text("dict", clip=clip_rect)["blocks"]
    extracted_lines = []
    
    for b in blocks:
        if "lines" not in b:
            continue
        for l in b["lines"]:
            spans = l["spans"]
            if not spans:
                continue
            line_text_parts = []
            max_size = 0.0
            is_bold = False
            is_italic = False
            
            for s in spans:
                font = s["font"]
                raw_txt = s["text"]
                if is_symbol_font(font):
                    conv_txt = "".join(ORNAMENT_MAP.get(c, c) for c in raw_txt)
                elif is_latin_font(font):
                    conv_txt = raw_txt
                else:
                    conv_txt = convert_tam_to_unicode(raw_txt)
                conv_txt = sanitize_xml(conv_txt)
                line_text_parts.append(conv_txt)
                if s["size"] > max_size:
                    max_size = s["size"]
                if "bold" in font.lower() or (s.get("flags", 0) & 2):
                    is_bold = True
                if "italic" in font.lower() or (s.get("flags", 0) & 1):
                    is_italic = True
                    
            full_line_text = "".join(line_text_parts).rstrip("\r\n")
            if full_line_text.strip():
                extracted_lines.append({
                    "text": full_line_text,
                    "bbox": l["bbox"], # (x0, y0, x1, y1)
                    "size": max_size,
                    "bold": is_bold,
                    "italic": is_italic,
                })
    extracted_lines.sort(key=lambda item: item["bbox"][1])
    return extracted_lines

def convert_pdf_to_docx(pdf_path, output_docx_path, book_title=None, author=""):
    """
    Main conversion function. Converts any Tamil PDF into a beautifully styled .docx document.
    Automatically handles 2-page landscape spreads and single-page portraits.
    """
    doc_pdf = pymupdf.open(pdf_path)
    
    # Auto-detect spread: if width significantly exceeds height
    p0 = doc_pdf[0]
    is_spread = (p0.rect.width > p0.rect.height * 1.15)
    
    # Auto-detect title if not given
    if not book_title:
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        book_title = base_name
    
    doc_word = docx.Document()
    
    # Set page margins: 1 inch top/bottom, 1.2 inch left/right
    for s in doc_word.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.2)
        s.right_margin = Inches(1.2)
        
        # Word Header: Book title and author
        header = s.header
        p_hdr = header.paragraphs[0]
        p_hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hdr_text = f"{book_title} — {author}" if author else book_title
        r_hdr = p_hdr.add_run(hdr_text)
        set_font_run(r_hdr, font_name="Nirmala UI", size_pt=9, italic=True, color_rgb=(120, 120, 120))
        
        # Word Footer: Page number
        footer = s.footer
        p_ftr = footer.paragraphs[0]
        p_ftr.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fldSimple = OxmlElement('w:fldSimple')
        fldSimple.set(qn('w:instr'), 'PAGE')
        p_ftr._p.append(fldSimple)
        
    # Book Cover / Front Title
    p_title = doc_word.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(72)
    p_title.paragraph_format.space_after = Pt(16)
    r_title = p_title.add_run(book_title)
    set_font_run(r_title, font_name="Nirmala UI", size_pt=26, bold=True, color_rgb=(30, 30, 30))
    
    p_orn = doc_word.add_paragraph()
    p_orn.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_orn.paragraph_format.space_after = Pt(16)
    r_orn = p_orn.add_run("❖ ❖ ❖")
    set_font_run(r_orn, font_name="Nirmala UI", size_pt=14, color_rgb=(100, 100, 100))
    
    if author:
        p_auth = doc_word.add_paragraph()
        p_auth.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_auth.paragraph_format.space_after = Pt(72)
        r_auth = p_auth.add_run(author)
        set_font_run(r_auth, font_name="Nirmala UI", size_pt=18, bold=True, italic=True, color_rgb=(60, 60, 60))
    
    doc_word.add_page_break()
    
    # Extract all book pages
    total_pages = len(doc_pdf)
    book_pages_content = []
    
    for pidx in range(total_pages):
        pdf_page = doc_pdf[pidx]
        if is_spread:
            mid_x = pdf_page.rect.width / 2.0
            rect_left = pymupdf.Rect(0, 0, mid_x, pdf_page.rect.height)
            rect_right = pymupdf.Rect(mid_x, 0, pdf_page.rect.width, pdf_page.rect.height)
            
            lines_l = extract_page_lines(pdf_page, clip_rect=rect_left)
            lines_r = extract_page_lines(pdf_page, clip_rect=rect_right)
            
            book_pages_content.append((lines_l, mid_x, pdf_page.rect.height))
            book_pages_content.append((lines_r, mid_x, pdf_page.rect.height))
        else:
            lines = extract_page_lines(pdf_page)
            book_pages_content.append((lines, pdf_page.rect.width, pdf_page.rect.height))
            
    is_first_page = True
    for page_idx, (lines, pw, ph) in enumerate(book_pages_content):
        # Filter headers and footers
        clean_lines = []
        for l in lines:
            if not is_running_header_footer(l["text"], l["bbox"][1], l["bbox"][3], ph, book_title=book_title, author=author):
                clean_lines.append(l)
                
        if not clean_lines:
            continue
            
        if not is_first_page:
            doc_word.add_page_break()
        is_first_page = False
        
        pc = pw / 2.0
        prev_bottom_y = None
        
        for i, line in enumerate(clean_lines):
            txt = line["text"].strip()
            if not txt:
                continue
                
            x0, y0, x1, y1 = line["bbox"]
            cw = x1 - x0
            cx = (x0 + x1) / 2.0
            
            # Alignment check
            if abs(cx - pc) < 32.0 and cw < pw * 0.75:
                alignment = WD_ALIGN_PARAGRAPH.CENTER
                left_indent = Inches(0)
            elif x0 > pw * 0.28:
                alignment = WD_ALIGN_PARAGRAPH.LEFT
                indent_inch = max(0.0, (x0 - pw * 0.18) / 72.0)
                left_indent = Inches(min(indent_inch, 1.5))
            else:
                alignment = WD_ALIGN_PARAGRAPH.LEFT
                left_indent = Inches(0)
                
            # Stanza break calculation
            space_before = 0
            if prev_bottom_y is not None:
                gap = y0 - prev_bottom_y
                if gap > 14.0:
                    space_before = min(gap * 0.6, 14.0)
            prev_bottom_y = y1
            
            # Title vs verse styling
            is_heading = (line["size"] >= 12.5) or (line["bold"] and line["size"] >= 11.0)
            
            p = doc_word.add_paragraph()
            p.alignment = alignment
            p.paragraph_format.left_indent = left_indent
            p.paragraph_format.line_spacing = 1.25
            
            if is_heading:
                p.paragraph_format.space_before = Pt(max(space_before, 12))
                p.paragraph_format.space_after = Pt(8)
                run = p.add_run(txt)
                set_font_run(run, font_name="Nirmala UI", size_pt=15, bold=True, color_rgb=(20, 20, 20))
            else:
                p.paragraph_format.space_before = Pt(space_before)
                p.paragraph_format.space_after = Pt(2)
                run = p.add_run(txt)
                set_font_run(run, font_name="Nirmala UI", size_pt=11.5, bold=line["bold"], italic=line["italic"])
                
    os.makedirs(os.path.dirname(os.path.abspath(output_docx_path)), exist_ok=True)
    doc_word.save(output_docx_path)
    return output_docx_path

if __name__ == "__main__":
    if len(sys.argv) > 1:
        in_pdf = sys.argv[1]
        out_doc = sys.argv[2] if len(sys.argv) > 2 else os.path.splitext(in_pdf)[0] + ".docx"
        convert_pdf_to_docx(in_pdf, out_doc)
        print(f"Converted {in_pdf} -> {out_doc}")
    else:
        print("Usage: python converter.py <input.pdf> [output.docx]")
