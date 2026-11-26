import json, glob, os, re, html, subprocess
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

BASE = r"D:\parmatmapanchvinshika\parmatmapanch"
SRC = os.path.join(BASE, "src")

# Load front & back matter
with open(os.path.join(SRC, "front_and_back_matter.json"), encoding="utf-8") as f:
    fm = json.load(f)

# Load all 25 shlokas in order
shlokas = []
for fp in sorted(glob.glob(os.path.join(SRC, "shlokas_*.json"))):
    with open(fp, encoding="utf-8") as f:
        shlokas.extend(json.load(f))
shlokas.sort(key=lambda x: x["num"])
assert len(shlokas) == 25, f"Expected 25 shlokas, got {len(shlokas)}"

# Helper: parse **bold** segments into (text, is_bold)
def parse_bold(text):
    parts = re.split(r"(\*\*.*?\*\*)", text, flags=re.DOTALL)
    out = []
    for p in parts:
        if p.startswith("**") and p.endswith("**") and len(p) >= 4:
            out.append((p[2:-2], True))
        elif p:
            out.append((p, False))
    return out

# Helper: add formatted paragraph to DOCX
def add_rich_para(doc, text, font_size=11, color=(30, 30, 35), align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=6, space_before=0, bold_all=False, italic_all=False, left_indent=0):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.35
    if left_indent:
        p.paragraph_format.left_indent = Inches(left_indent)
    for seg, is_b in parse_bold(text):
        r = p.add_run(seg)
        r.font.name = "Nirmala UI"
        r.font.size = Pt(font_size)
        r.font.color.rgb = RGBColor(*color)
        r.bold = bold_all or is_b
        r.italic = italic_all
    return p

def shade_para(p, hex_color):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    p._p.get_or_add_pPr().append(shading)

def set_borders(p, color="B8860B", sz="8"):
    pPr = p._p.get_or_add_pPr()
    pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:left w:val="single" w:sz="{sz}" w:space="8" w:color="{color}"/></w:pBdr>')
    pPr.append(pBdr)

# ================= BUILD DOCX =================
doc = Document()
for sec in doc.sections:
    sec.top_margin = Inches(0.85)
    sec.bottom_margin = Inches(0.85)
    sec.left_margin = Inches(0.95)
    sec.right_margin = Inches(0.95)

# Page 1: Title Page
add_rich_para(doc, fm["title_page"]["invocation"], 14, (139, 69, 19), WD_ALIGN_PARAGRAPH.CENTER, 24, 60, bold_all=True)
add_rich_para(doc, fm["title_page"]["sanskrit_title"], 26, (128, 0, 32), WD_ALIGN_PARAGRAPH.CENTER, 12, 20, bold_all=True)
add_rich_para(doc, fm["title_page"]["gujarati_title"], 20, (128, 0, 32), WD_ALIGN_PARAGRAPH.CENTER, 28, 0, bold_all=True)
add_rich_para(doc, fm["title_page"]["author"], 13, (40, 40, 80), WD_ALIGN_PARAGRAPH.CENTER, 24, 10, bold_all=True)
add_rich_para(doc, fm["title_page"]["subtitle"], 11.5, (60, 60, 60), WD_ALIGN_PARAGRAPH.CENTER, 12, 20)
doc.add_page_break()

# Page 2: Single Mangala Shloka Page
add_rich_para(doc, fm["page2_shloka"]["heading"], 16, (128, 0, 32), WD_ALIGN_PARAGRAPH.CENTER, 24, 120, bold_all=True)
p_m = add_rich_para(doc, fm["page2_shloka"]["sanskrit"], 15, (25, 25, 90), WD_ALIGN_PARAGRAPH.CENTER, 16, 10, bold_all=True)
shade_para(p_m, "FDF6E3")
add_rich_para(doc, fm["page2_shloka"]["source"], 10.5, (100, 60, 20), WD_ALIGN_PARAGRAPH.CENTER, 24, 4, italic_all=True)
add_rich_para(doc, fm["page2_shloka"]["gujarati"], 12, (40, 40, 45), WD_ALIGN_PARAGRAPH.CENTER, 12, 10)
doc.add_page_break()

# Page 3: Prastavana
add_rich_para(doc, fm["prastavana"]["title"], 16, (128, 0, 32), WD_ALIGN_PARAGRAPH.LEFT, 14, 0, bold_all=True)
for para in fm["prastavana"]["paragraphs"]:
    add_rich_para(doc, para, 11, (30, 30, 35), WD_ALIGN_PARAGRAPH.JUSTIFY, 10)
doc.add_page_break()

# Granth Parichaya
add_rich_para(doc, fm["parichaya"]["title"], 16, (128, 0, 32), WD_ALIGN_PARAGRAPH.LEFT, 14, 0, bold_all=True)
for para in fm["parichaya"]["paragraphs"]:
    add_rich_para(doc, para, 11, (30, 30, 35), WD_ALIGN_PARAGRAPH.JUSTIFY, 10)
doc.add_page_break()

# Vishayanukramanika (TOC)
add_rich_para(doc, "विषयानुक्रमणिका (વિષયાનુક્રમણિકા)", 16, (128, 0, 32), WD_ALIGN_PARAGRAPH.CENTER, 16, 0, bold_all=True)
toc_items = [
    "१. मुखपृष्ठ (ગ્રંથનું નામ)",
    "२. मङ्गलश्लोकः",
    "३. ગ્રંથ પ્રસ્તાવના : ભક્તિરસ અને સ્વસંવેદનના મહાસાગરમાં પ્રવેશ",
    "૪. ગ્રંથ પરિચય : શૈલી, શાસ્ત્રીય બંધારણ અને ભાવાર્થનું વિશ્લેષણ",
]
for t in toc_items:
    add_rich_para(doc, t, 11, (30, 30, 70), WD_ALIGN_PARAGRAPH.LEFT, 4, bold_all=True)
for s in shlokas:
    add_rich_para(doc, f"શ્લોક {s['num']} : {s['title']}", 10.5, (30, 30, 35), WD_ALIGN_PARAGRAPH.LEFT, 3, left_indent=0.2)
add_rich_para(doc, "પરિશિષ્ટ ૧ : મૂળ પરમાત્મપંચવિંશતિકા (સંપૂર્ણ ૨૫ શ્લોકો અને ૪–૫ પંક્તિનો ભાવાનુવાદ)", 11, (128, 0, 32), WD_ALIGN_PARAGRAPH.LEFT, 4, 6, bold_all=True)
add_rich_para(doc, "પરિશિષ્ટ ૨ : પરમાત્મપંચવિંશતિકા — સંપૂર્ણ ૨૫ ગુર્જર કાવ્ય (દોહરા)", 11, (128, 0, 32), WD_ALIGN_PARAGRAPH.LEFT, 4, 2, bold_all=True)
doc.add_page_break()

# Shlokas 1 to 25
for idx, s in enumerate(shlokas):
    add_rich_para(doc, f"શ્લોક {s['num']}", 11, (139, 69, 19), WD_ALIGN_PARAGRAPH.CENTER, 2, 0, bold_all=True)
    add_rich_para(doc, s["title"], 14.5, (128, 0, 32), WD_ALIGN_PARAGRAPH.CENTER, 12, 0, bold_all=True)

    add_rich_para(doc, "संस्कृत अवतरणिका", 11.5, (128, 0, 32), WD_ALIGN_PARAGRAPH.LEFT, 4, 6, bold_all=True)
    add_rich_para(doc, s["avataranika_sa"], 11, (30, 30, 35), WD_ALIGN_PARAGRAPH.JUSTIFY, 8)

    add_rich_para(doc, "અવતરણિકાનો ગુજરાતી અનુવાદ", 11.5, (0, 90, 90), WD_ALIGN_PARAGRAPH.LEFT, 4, 4, bold_all=True)
    add_rich_para(doc, s["avataranika_gu"], 11, (30, 30, 35), WD_ALIGN_PARAGRAPH.JUSTIFY, 10)

    add_rich_para(doc, "मूल संस्कृत श्लोकः", 11.5, (128, 0, 32), WD_ALIGN_PARAGRAPH.CENTER, 4, 6, bold_all=True)
    ps = add_rich_para(doc, s["shloka"], 13, (20, 20, 90), WD_ALIGN_PARAGRAPH.CENTER, 10, 2, bold_all=True)
    shade_para(ps, "FDF6E3")

    add_rich_para(doc, "ગુર્જર કાવ્ય (દોહરો)", 11.5, (139, 69, 19), WD_ALIGN_PARAGRAPH.CENTER, 4, 4, bold_all=True)
    pk = add_rich_para(doc, s["gurjar_kavya"], 11, (60, 30, 10), WD_ALIGN_PARAGRAPH.CENTER, 10, 2, italic_all=True)
    shade_para(pk, "F5F9F5")

    add_rich_para(doc, "સરળ ભાવાનુવાદ", 11.5, (0, 90, 90), WD_ALIGN_PARAGRAPH.LEFT, 4, 4, bold_all=True)
    pb = add_rich_para(doc, s["bhavanuvada"], 11, (30, 30, 35), WD_ALIGN_PARAGRAPH.LEFT, 10, left_indent=0.15)
    set_borders(pb, "006666", "6")

    add_rich_para(doc, "संस्कृत वृत्तिः (अन्वयार्थ-क्रमेण)", 12, (128, 0, 32), WD_ALIGN_PARAGRAPH.LEFT, 6, 8, bold_all=True)
    for blk in s["vritti_sa"].split("\n\n"):
        if blk.strip().startswith(">"):
            q = "\n".join([line.lstrip("> ").strip() for line in blk.splitlines()])
            pq = add_rich_para(doc, q, 11, (25, 25, 100), WD_ALIGN_PARAGRAPH.LEFT, 8, left_indent=0.25)
            shade_para(pq, "F8F4EC")
            set_borders(pq, "B8860B", "8")
        else:
            add_rich_para(doc, blk.strip(), 11, (30, 30, 35), WD_ALIGN_PARAGRAPH.JUSTIFY, 7)

    add_rich_para(doc, "સંસ્કૃત વૃત્તિનો ગુજરાતી અનુવાદ", 12, (0, 90, 90), WD_ALIGN_PARAGRAPH.LEFT, 6, 8, bold_all=True)
    for blk in s["vritti_gu"].split("\n\n"):
        if blk.strip().startswith(">"):
            q = "\n".join([line.lstrip("> ").strip() for line in blk.splitlines()])
            pq = add_rich_para(doc, q, 10.5, (25, 70, 70), WD_ALIGN_PARAGRAPH.LEFT, 8, left_indent=0.25)
            shade_para(pq, "F2F8F8")
            set_borders(pq, "008080", "8")
        else:
            add_rich_para(doc, blk.strip(), 11, (30, 30, 35), WD_ALIGN_PARAGRAPH.JUSTIFY, 7)

    add_rich_para(doc, "વ્યાવહારિક સાધના ખંડ", 12, (128, 0, 32), WD_ALIGN_PARAGRAPH.LEFT, 6, 8, bold_all=True)
    add_rich_para(doc, "૩ સાધના પ્રયોગો :", 11, (139, 69, 19), WD_ALIGN_PARAGRAPH.LEFT, 3, 2, bold_all=True)
    for i, pr in enumerate(s["sadhana_prayog"], 1):
        add_rich_para(doc, f"{i}. {pr}", 10.5, (30, 30, 35), WD_ALIGN_PARAGRAPH.LEFT, 3, left_indent=0.2)

    add_rich_para(doc, "૫ આત્મપ્રશ્નો :", 11, (139, 69, 19), WD_ALIGN_PARAGRAPH.LEFT, 3, 6, bold_all=True)
    for i, pq in enumerate(s["atmaprashna"], 1):
        add_rich_para(doc, f"{i}. {pq}", 10.5, (30, 30, 35), WD_ALIGN_PARAGRAPH.LEFT, 3, left_indent=0.2)

    add_rich_para(doc, "૫ સાધના સૂત્રો (પંચલાઇન) :", 11, (139, 69, 19), WD_ALIGN_PARAGRAPH.LEFT, 3, 6, bold_all=True)
    for i, psut in enumerate(s["sadhana_sutra"], 1):
        add_rich_para(doc, f"{i}. **{psut}**", 10.5, (80, 20, 20), WD_ALIGN_PARAGRAPH.LEFT, 3, left_indent=0.2)

    doc.add_page_break()

# Parishishta 1: All 25 Shlokas + Bhavanuvada
add_rich_para(doc, "પરિશિષ્ટ ૧ : મૂળ પરમાત્મપંચવિંશતિકા (સંપૂર્ણ ૨૫ શ્લોકો અને ભાવાનુવાદ)", 15, (128, 0, 32), WD_ALIGN_PARAGRAPH.CENTER, 16, 0, bold_all=True)
for s in shlokas:
    ps = add_rich_para(doc, s["shloka"], 12, (20, 20, 90), WD_ALIGN_PARAGRAPH.CENTER, 4, 6, bold_all=True)
    shade_para(ps, "FDF6E3")
    add_rich_para(doc, s["bhavanuvada"], 10.5, (30, 30, 35), WD_ALIGN_PARAGRAPH.LEFT, 10, left_indent=0.2)
doc.add_page_break()

# Parishishta 2: All 25 Gurjar Kavyas
add_rich_para(doc, "પરિશિષ્ટ ૨ : પરમાત્મપંચવિંશતિકા — સંપૂર્ણ ૨૫ ગુર્જર કાવ્ય (દોહરા)", 15, (128, 0, 32), WD_ALIGN_PARAGRAPH.CENTER, 16, 0, bold_all=True)
for s in shlokas:
    add_rich_para(doc, f"દોહરો {s['num']}", 10.5, (139, 69, 19), WD_ALIGN_PARAGRAPH.CENTER, 2, 6, bold_all=True)
    pk = add_rich_para(doc, s["gurjar_kavya"], 11, (40, 20, 10), WD_ALIGN_PARAGRAPH.CENTER, 8, 0, italic_all=True)
    shade_para(pk, "F5F9F5")

docx_path = os.path.join(BASE, "Paramatma_Panchavimshatika_Complete.docx")
doc.save(docx_path)
print("Saved DOCX:", docx_path)

# ================= BUILD HTML & PDF =================
def md_to_html(text):
    esc = html.escape(text)
    esc = re.sub(r"\*\*(.*?)\*\*", r"<strong>\1</strong>", esc, flags=re.DOTALL)
    return esc.replace("\n", "<br>")

def vritti_to_html(text):
    blocks = []
    for blk in text.split("\n\n"):
        b = blk.strip()
        if b.startswith(">"):
            q = "\n".join([line.lstrip("> ").strip() for line in b.splitlines()])
            blocks.append(f'<blockquote class="ref-quote">{md_to_html(q)}</blockquote>')
        else:
            blocks.append(f'<p>{md_to_html(b)}</p>')
    return "\n".join(blocks)

html_parts = [f"""<!DOCTYPE html>
<html lang="sa">
<head>
<meta charset="UTF-8">
<title>श्रीपरमात्मपञ्चविंशतिका</title>
<style>
  @page {{
    size: A4;
    margin: 20mm 18mm 20mm 18mm;
    @bottom-center {{
      content: "— " counter(page) " —";
      font-family: "Nirmala UI", sans-serif;
      font-size: 9.5pt;
      color: #666;
    }}
  }}
  body {{
    font-family: "Nirmala UI", "Noto Sans Devanagari", "Noto Sans Gujarati", sans-serif;
    font-size: 11pt;
    line-height: 1.65;
    color: #222;
    margin: 0;
  }}
  .page-break {{ page-break-after: always; }}
  .title-page {{
    text-align: center;
    padding-top: 45mm;
  }}
  .invocation {{ font-size: 14pt; color: #8B4513; font-weight: bold; margin-bottom: 18mm; }}
  .main-title-sa {{ font-size: 28pt; color: #800020; font-weight: bold; margin-bottom: 6mm; }}
  .main-title-gu {{ font-size: 21pt; color: #800020; font-weight: bold; margin-bottom: 14mm; }}
  .author {{ font-size: 13pt; color: #282850; font-weight: bold; margin-bottom: 12mm; }}
  .subtitle {{ font-size: 11pt; color: #444; max-width: 85%; margin: 0 auto; }}
  .page2 {{ text-align: center; padding-top: 55mm; }}
  .shloka-box {{
    background: #FDF6E3;
    border: 1.5px solid #D4AF37;
    border-radius: 6px;
    padding: 12px 18px;
    text-align: center;
    font-size: 13pt;
    font-weight: bold;
    color: #14145A;
    margin: 10px 0 14px 0;
  }}
  .kavya-box {{
    background: #F4F9F4;
    border: 1px solid #8FBC8F;
    border-radius: 6px;
    padding: 10px 16px;
    text-align: center;
    font-size: 11pt;
    color: #2E472E;
    margin: 8px 0 12px 0;
  }}
  .bhav-box {{
    border-left: 4px solid #006666;
    padding: 6px 14px;
    background: #F7FAFA;
    margin: 8px 0 14px 0;
  }}
  .ref-quote {{
    background: #F9F4EA;
    border-left: 4px solid #B8860B;
    margin: 10px 14px;
    padding: 8px 14px;
    color: #1E1E5A;
  }}
  h1 {{ color: #800020; font-size: 17pt; border-bottom: 2px solid #D4AF37; padding-bottom: 4px; }}
  h2.shloka-title {{ color: #800020; font-size: 14.5pt; text-align: center; margin-top: 2px; margin-bottom: 14px; }}
  .shloka-num {{ text-align: center; color: #8B4513; font-weight: bold; font-size: 11pt; }}
  h3.sec-sa {{ color: #800020; font-size: 12pt; margin-top: 14px; margin-bottom: 4px; }}
  h3.sec-gu {{ color: #005A5A; font-size: 12pt; margin-top: 14px; margin-bottom: 4px; }}
  .sadhana-box {{
    background: #FFFDF8;
    border: 1px solid #E5D4B3;
    border-radius: 6px;
    padding: 10px 16px;
    margin-top: 12px;
  }}
  .sadhana-box h4 {{ color: #8B4513; margin: 10px 0 4px 0; font-size: 11pt; }}
  ol {{ margin: 4px 0 8px 0; padding-left: 20px; }}
  li {{ margin-bottom: 3px; }}
  .toc-list {{ list-style: none; padding-left: 0; }}
  .toc-list li {{ margin-bottom: 4px; }}
  .toc-sub {{ padding-left: 18px; font-size: 10.5pt; }}
  p {{ text-align: justify; margin: 0 0 8px 0; }}
</style>
</head>
<body>
"""]

# Page 1
html_parts.append(f"""
<div class="title-page page-break">
  <div class="invocation">{md_to_html(fm["title_page"]["invocation"])}</div>
  <div class="main-title-sa">{md_to_html(fm["title_page"]["sanskrit_title"])}</div>
  <div class="main-title-gu">{md_to_html(fm["title_page"]["gujarati_title"])}</div>
  <div class="author">{md_to_html(fm["title_page"]["author"])}</div>
  <div class="subtitle">{md_to_html(fm["title_page"]["subtitle"])}</div>
</div>
""")

# Page 2
html_parts.append(f"""
<div class="page2 page-break">
  <h2 style="color:#800020;">{md_to_html(fm["page2_shloka"]["heading"])}</h2>
  <div class="shloka-box">{md_to_html(fm["page2_shloka"]["sanskrit"])}</div>
  <div style="font-style:italic; color:#643C14; margin-bottom:16px;">{md_to_html(fm["page2_shloka"]["source"])}</div>
  <div style="font-size:12pt; color:#333;">{md_to_html(fm["page2_shloka"]["gujarati"])}</div>
</div>
""")

# Page 3: Prastavana
html_parts.append('<div class="page-break"><h1>' + md_to_html(fm["prastavana"]["title"]) + '</h1>')
for p in fm["prastavana"]["paragraphs"]:
    html_parts.append(f"<p>{md_to_html(p)}</p>")
html_parts.append('</div>')

# Parichaya
html_parts.append('<div class="page-break"><h1>' + md_to_html(fm["parichaya"]["title"]) + '</h1>')
for p in fm["parichaya"]["paragraphs"]:
    html_parts.append(f"<p>{md_to_html(p)}</p>")
html_parts.append('</div>')

# TOC
html_parts.append('<div class="page-break"><h1 style="text-align:center;">विषयानुक्रमणिका (વિષયાનુક્રમણિકા)</h1><ul class="toc-list">')
for t in toc_items:
    html_parts.append(f"<li><strong>{md_to_html(t)}</strong></li>")
for s in shlokas:
    html_parts.append(f'<li class="toc-sub">શ્લોક {s["num"]} : {md_to_html(s["title"])}</li>')
html_parts.append('<li><strong>પરિશિષ્ટ ૧ : મૂળ પરમાત્મપંચવિંશતિકા (સંપૂર્ણ ૨૫ શ્લોકો અને ભાવાનુવાદ)</strong></li>')
html_parts.append('<li><strong>પરિશિષ્ટ ૨ : પરમાત્મપંચવિંશતિકા — સંપૂર્ણ ૨૫ ગુર્જર કાવ્ય (દોહરા)</strong></li>')
html_parts.append('</ul></div>')

# Shlokas 1-25
for s in shlokas:
    html_parts.append('<div class="page-break">')
    html_parts.append(f'<div class="shloka-num">શ્લોક {s["num"]}</div>')
    html_parts.append(f'<h2 class="shloka-title">{md_to_html(s["title"])}</h2>')
    html_parts.append(f'<h3 class="sec-sa">संस्कृत अवतरणिका</h3><p>{md_to_html(s["avataranika_sa"])}</p>')
    html_parts.append(f'<h3 class="sec-gu">અવતરણિકાનો ગુજરાતી અનુવાદ</h3><p>{md_to_html(s["avataranika_gu"])}</p>')
    html_parts.append(f'<h3 class="sec-sa" style="text-align:center;">मूल संस्कृत श्लोकः</h3><div class="shloka-box">{md_to_html(s["shloka"])}</div>')
    html_parts.append(f'<h3 class="sec-sa" style="text-align:center; color:#8B4513;">ગુર્જર કાવ્ય (દોહરો)</h3><div class="kavya-box">{md_to_html(s["gurjar_kavya"])}</div>')
    html_parts.append(f'<h3 class="sec-gu">સરળ ભાવાનુવાદ</h3><div class="bhav-box">{md_to_html(s["bhavanuvada"])}</div>')
    html_parts.append(f'<h3 class="sec-sa">संस्कृत वृत्तिः (अन्वयार्थ-क्रमेण)</h3>{vritti_to_html(s["vritti_sa"])}')
    html_parts.append(f'<h3 class="sec-gu">સંસ્કૃત વૃત્તિનો ગુજરાતી અનુવાદ</h3>{vritti_to_html(s["vritti_gu"])}')
    html_parts.append('<div class="sadhana-box"><h3 class="sec-sa" style="margin-top:0;">વ્યાવહારિક સાધના ખંડ</h3>')
    html_parts.append('<h4>૩ સાધના પ્રયોગો :</h4><ol>' + "".join([f"<li>{md_to_html(x)}</li>" for x in s["sadhana_prayog"]]) + '</ol>')
    html_parts.append('<h4>૫ આત્મપ્રશ્નો :</h4><ol>' + "".join([f"<li>{md_to_html(x)}</li>" for x in s["atmaprashna"]]) + '</ol>')
    html_parts.append('<h4>૫ સાધના સૂત્રો (પંચલાઇન) :</h4><ol>' + "".join([f"<li><strong>{md_to_html(x)}</strong></li>" for x in s["sadhana_sutra"]]) + '</ol>')
    html_parts.append('</div></div>')

# Parishishta 1
html_parts.append('<div class="page-break"><h1 style="text-align:center;">પરિશિષ્ટ ૧ : મૂળ પરમાત્મપંચવિંશતિકા (સંપૂર્ણ ૨૫ શ્લોકો અને ભાવાનુવાદ)</h1>')
for s in shlokas:
    html_parts.append(f'<div class="shloka-box">{md_to_html(s["shloka"])}</div><div class="bhav-box">{md_to_html(s["bhavanuvada"])}</div>')
html_parts.append('</div>')

# Parishishta 2
html_parts.append('<div><h1 style="text-align:center;">પરિશિષ્ટ ૨ : પરમાત્મપંચવિંશતિકા — સંપૂર્ણ ૨૫ ગુર્જર કાવ્ય (દોહરા)</h1>')
for s in shlokas:
    html_parts.append(f'<div class="shloka-num" style="margin-top:10px;">દોહરો {s["num"]}</div><div class="kavya-box">{md_to_html(s["gurjar_kavya"])}</div>')
html_parts.append('</div></body></html>')

html_path = os.path.join(BASE, "Paramatma_Panchavimshatika_Complete.html")
with open(html_path, "w", encoding="utf-8") as f:
    f.write("\n".join(html_parts))
print("Saved HTML:", html_path)

pdf_path = os.path.join(BASE, "Paramatma_Panchavimshatika_Complete.pdf")
chrome_exe = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
subprocess.run([chrome_exe, "--headless", "--disable-gpu", f"--print-to-pdf={pdf_path}", "--no-pdf-header-footer", html_path], check=True)
print("Saved PDF:", pdf_path)
