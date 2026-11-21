# 20.-Automatinis-pristatymo-skaidri-generavimas

# 20.-Automatinis-pristatymo-skaidri-generavimas

## Text / Word to PowerPoint generator

Turn pasted text or a Word `.docx` document into a downloadable PowerPoint presentation.
The app runs in Google Colab and uses Gradio to provide a web interface.

### Run the app in Google Colab

1. Create a new notebook in [Google Colab](https://colab.research.google.com/).
2. In the first cell, install the required packages and run the cell:

   ```python
   !pip install -q gradio python-pptx python-docx requests
   ```

3. Copy the Python code below into a second cell and run it.
4. When it starts, open the public `https://....gradio.live` link printed in the cell output.
   Keep the Colab cell running while you use the app; the shared link is temporary and stops
   working when the app or its Colab runtime stops.
5. In the app, paste text or upload a `.docx`, choose a theme and structuring mode, then select
   **Generate presentation**. Download the resulting `.pptx` file from the page.

Local structuring works without an API key. OpenAI and Google Gemini modes require an API key
from the selected provider.

### App code

```python
# =====================================================================
#  Text / Word  ->  PowerPoint generator   (Google Colab + Gradio)
#
#  CELL 1 (run first, then paste this code into CELL 2):
#       !pip install -q gradio python-pptx python-docx requests
# =====================================================================

import os
import re
import json
import math
import tempfile
import datetime
from collections import Counter

import requests
import gradio as gr
from lxml import etree
from docx import Document
from docx.oxml.ns import qn as docx_qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn as pptx_qn

# ---------------------------------------------------------------------
# Themes
# ---------------------------------------------------------------------
THEMES = {
    "Modern Blue": dict(
        bg="FFFFFF", dark_bg="0B2545", heading="0B2545", text="1F2937",
        muted="64748B", accent="1E88E5", card="EEF4FB", on_dark="FFFFFF",
        on_dark_muted="B6C7E0", badge_fg="FFFFFF",
        font_head="Calibri", font_body="Calibri"),
    "Minimalist Dark": dict(
        bg="14171F", dark_bg="0A0C11", heading="FFFFFF", text="E5E7EB",
        muted="8B93A3", accent="2DD4BF", card="1D222D", on_dark="FFFFFF",
        on_dark_muted="9AA3B5", badge_fg="0A0C11",
        font_head="Calibri", font_body="Calibri"),
    "Warm Corporate": dict(
        bg="FFF9F2", dark_bg="6B2A10", heading="6B2A10", text="3B3B3B",
        muted="8A7F76", accent="E8590C", card="FBEBD9", on_dark="FFF7ED",
        on_dark_muted="F3C9A8", badge_fg="FFFFFF",
        font_head="Georgia", font_body="Calibri"),
}

# ---------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------
STOP = set("""a about above after again against all also am an and any are aren't as at be because been before
being below between both but by can could did do does doing down during each few for from further had has have
having he her here hers him his how i if in into is it its itself just let may me might more most must my no nor
not of off on once one only or other our ours out over own same she should so some such than that the their theirs
them then there these they this those through to too under until up us use used using very was we were what when
where which while who whom why will with would you your yours however thus therefore among within without across
per via etc many much like get gets got make makes made cannot don't doesn't isn't it's can't""".split())

ABBREV_END_RE = re.compile(r'\b(?:Dr|Mr|Mrs|Ms|Prof|Sr|Jr|vs|etc|e\.g|i\.e|Inc|Ltd|No|Fig)\.$')
LIST_MARK_RE = re.compile(r'^\s*(?:[-*\u2022\u25AA\u25CF\u25E6\u2023\u2013\u2014]|\d{1,2}[.)]|[a-zA-Z][.)])\s+')
BULLET_RE = re.compile(r'^\s*(?:[-*\u2022\u25AA\u25CF\u25E6\u2023\u2013\u2014]|\d{1,2}[.)]|[a-zA-Z][.)])\s+(.*\S)\s*$')
NUM_PREFIX_RE = re.compile(r'^\s*(?:\d+(?:\.\d+)+[.)]?|\d+[.)]|[IVXLC]+[.)])\s+')


def clean_text(s):
    s = re.sub(r'\[\d+\]', '', s or '')
    return re.sub(r'\s+', ' ', s).strip()


def clean_title(t):
    t = clean_text(t)
    t = NUM_PREFIX_RE.sub('', t).strip().rstrip(':').strip()
    if len(t) > 3 and t.isupper():
        t = t.title()
    return t or "Untitled"


def shorten_title(t, limit=70):
    t = clean_text(t)
    if len(t) <= limit:
        return t
    cut = t[:limit]
    sp = cut.rfind(" ")
    if sp > limit * 0.5:
        cut = cut[:sp]
    return cut.rstrip(" ,;:-\u2013\u2014") + "\u2026"


def shorten(text, limit=150):
    t = clean_text(text)
    t = re.sub(r'\s*\([^)]{25,}\)', '', t)
    if t.endswith('.') and not t.endswith('..'):
        t = t[:-1]
    t = t.strip()
    if len(t) > limit:
        cut = t[:limit]
        idx = max(cut.rfind(sep) for sep in ["; ", ", ", " \u2013 ", " \u2014 ", ": ", " - "])
        if idx >= int(limit * 0.55):
            t = cut[:idx].rstrip(" ,;:-\u2013\u2014") + "\u2026"
        else:
            sp = cut.rfind(" ")
            t = (cut[:sp] if sp > 0 else cut).rstrip(" ,;:-\u2013\u2014") + "\u2026"
    return t[:1].upper() + t[1:] if t else t


def tokenize(text):
    return re.findall(r"[^\W\d_][\w'-]*", text.lower())


def split_sentences(text):
    text = clean_text(text)
    parts = re.split(r'(?<=[.!?])\s+(?=[A-Z0-9"\u201C(])', text)
    merged = []
    for p in parts:
        if merged and ABBREV_END_RE.search(merged[-1]):
            merged[-1] += " " + p
        else:
            merged.append(p)
    return [m.strip() for m in merged if m.strip()]


def keywords(text, n=3, avoid=()):
    freq = Counter(w for w in tokenize(text) if w not in STOP and len(w) >= 3)
    ranked = [w for w, _ in sorted(freq.items(), key=lambda kv: -kv[1])]
    fresh = [w for w in ranked if w.capitalize() not in avoid]
    chosen = []
    for w in fresh + [w for w in ranked if w not in fresh]:
        if any(w[:5] == c[:5] for c in chosen):
            continue
        chosen.append(w)
        if len(chosen) == n:
            break
    return [c.capitalize() for c in chosen] or ["Overview"]


def title_from_keywords(text, n=2, avoid=()):
    kws = keywords(text, n, avoid)
    if len(kws) == 1:
        return kws[0] + " Overview"
    return ", ".join(kws[:-1]) + " & " + kws[-1]


# ---------------------------------------------------------------------
# Input parsing  ->  blocks = [(kind, text, level)]   kind: h | p | b
# ---------------------------------------------------------------------
def parse_plain_text(text):
    lines = [l.rstrip() for l in (text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    n = len(lines)

    def next_nonempty(i):
        for j in range(i + 1, n):
            if lines[j].strip():
                return lines[j].strip()
        return None

    blocks = []
    prev_blank = True
    for i, raw in enumerate(lines):
        line = raw.strip()
        if not line:
            prev_blank = True
            continue
        nxt = next_nonempty(i)

        m = re.match(r'^(#{1,6})\s+(.*\S)$', line)
        if m:
            blocks.append(("h", clean_title(m.group(2)), len(m.group(1))))
            prev_blank = False
            continue

        bm = BULLET_RE.match(line)
        if bm:
            body = bm.group(1).strip()
            numbered = re.match(r'^\d{1,2}[.)]\s', line) is not None
            if (numbered and len(body) <= 60 and not re.search(r'[.!?,;:]$', body)
                    and nxt is not None and not LIST_MARK_RE.match(nxt)):
                blocks.append(("h", clean_title(line), 2))
            else:
                blocks.append(("b", body, 0))
            prev_blank = False
            continue

        words = line.split()
        stripped = line.rstrip(':').strip()
        is_caps = (stripped.isupper() and len(stripped) <= 80 and len(words) <= 12
                   and any(c.isalpha() for c in stripped))
        short_title = (prev_blank and len(line) <= 70 and len(words) <= 10
                       and not re.search(r'[.!?,;]$', line)
                       and nxt is not None and (len(nxt) > len(line) or not blocks))
        if is_caps or short_title:
            blocks.append(("h", clean_title(line), 1 if not blocks else 2))
            prev_blank = False
            continue

        if (blocks and blocks[-1][0] == "p" and not prev_blank
                and not re.search(r'[.!?:"\u201D\u2019)]$', blocks[-1][1])):
            blocks[-1] = ("p", blocks[-1][1] + " " + line, 0)
        else:
            blocks.append(("p", line, 0))
        prev_blank = False
    return blocks


def parse_docx(path):
    doc = Document(path)
    blocks = []
    for child in doc.element.body.iterchildren():
        if child.tag == docx_qn('w:p'):
            p = Paragraph(child, doc)
            text = clean_text(p.text)
            if not text:
                continue
            try:
                style = (p.style.name or "").lower()
            except Exception:
                style = ""
            if style == "title":
                blocks.append(("h", clean_title(text), 0))
            elif style.startswith("heading"):
                m = re.search(r'\d+', style)
                blocks.append(("h", clean_title(text), int(m.group()) if m else 1))
            elif "list" in style:
                blocks.append(("b", text, 0))
            else:
                has_num = False
                try:
                    has_num = child.pPr is not None and child.pPr.numPr is not None
                except Exception:
                    pass
                runs = [r for r in p.runs if r.text.strip()]
                all_bold = bool(runs) and all(r.bold for r in runs)
                if has_num:
                    blocks.append(("b", text, 0))
                elif all_bold and len(text) <= 70 and not re.search(r'[.!?,;]$', text):
                    blocks.append(("h", clean_title(text), 2))
                else:
                    blocks.append(("p", text, 0))
        elif child.tag == docx_qn('w:tbl'):
            table = Table(child, doc)
            for row in table.rows:
                cells = []
                for c in row.cells:
                    ct = clean_text(c.text)
                    if ct and (not cells or cells[-1] != ct):
                        cells.append(ct)
                if cells:
                    blocks.append(("b", " | ".join(cells), 0))
    return blocks


def blocks_to_text(blocks):
    out = []
    for kind, text, lvl in blocks:
        if kind == "h":
            out.append("\n## " + text)
        elif kind == "b":
            out.append("- " + text)
        else:
            out.append(text)
    return "\n".join(out).strip()


# ---------------------------------------------------------------------
# Local (rule-based) structuring
# ---------------------------------------------------------------------
def build_sections(blocks):
    heads = [b for b in blocks if b[0] == "h"]
    doc_title = None
    if heads and blocks[0][0] == "h":
        lvl = blocks[0][2]
        if all(h[2] > lvl for h in heads[1:]):
            doc_title = blocks[0][1]
            blocks = blocks[1:]
    sections = []
    cur_title, cur_items = None, []
    for kind, text, lvl in blocks:
        if kind == "h":
            if cur_items:
                sections.append((cur_title, cur_items))
            cur_title, cur_items = text, []
        else:
            cur_items.append((kind, text))
    if cur_items:
        sections.append((cur_title, cur_items))
    return doc_title, sections


def chunk_items(items):
    expanded = []
    for kind, text in items:
        if kind == "p" and len(text.split()) > 110:
            group, words = [], 0
            for sent in split_sentences(text):
                group.append(sent)
                words += len(sent.split())
                if words >= 70:
                    expanded.append(("p", " ".join(group)))
                    group, words = [], 0
            if group:
                expanded.append(("p", " ".join(group)))
        else:
            expanded.append((kind, text))
    items = expanded
    total = sum(len(t.split()) for _, t in items)
    target = max(150, total // 8)
    chunks, cur, words = [], [], 0
    for kind, text in items:
        cur.append((kind, text))
        words += len(text.split())
        if words >= target:
            chunks.append((None, cur))
            cur, words = [], 0
    if cur:
        if chunks and words < target * 0.4:
            chunks[-1] = (None, chunks[-1][1] + cur)
        else:
            chunks.append((None, cur))
    return chunks


def section_to_bullets(items, max_bullets):
    cap = max_bullets * 3
    full_text = " ".join(t for _, t in items)
    freq = Counter(w for w in tokenize(full_text) if w not in STOP and len(w) > 2)
    has_bullets = any(k == "b" for k, _ in items)

    cands = []  # (idx, text, score, mandatory)
    idx = 0
    n_sent = 0
    for kind, text in items:
        if kind == "b":
            cands.append((idx, text, 0.0, True))
            idx += 1
        else:
            sents = split_sentences(text)
            for si, s in enumerate(sents):
                words = s.split()
                if len(words) < 3:
                    continue
                toks = [w for w in tokenize(s) if w not in STOP and len(w) > 2]
                score = sum(freq[w] for w in toks) / (len(toks) ** 0.6 + 1) if toks else 0.0
                if si == 0:
                    score *= 1.25
                if len(words) > 45:
                    score *= 0.7
                cands.append((idx, s, score, False))
                idx += 1
                n_sent += 1

    n_mand = sum(1 for c in cands if c[3])
    if has_bullets:
        k = min(2, n_sent, max(0, cap - n_mand))
    elif n_sent <= 4:
        k = n_sent
    else:
        k = min(12, max(4, math.ceil(n_sent * 0.45)))
    k = min(k, cap)

    non_mand = sorted([c for c in cands if not c[3]], key=lambda c: -c[2])[:k]
    chosen = {c[0] for c in non_mand}
    final = [c for c in cands if c[3] or c[0] in chosen]
    final.sort(key=lambda c: c[0])

    bullets, seen = [], set()
    for c in final[:cap]:
        b = shorten(c[1], 150)
        if b and b.lower() not in seen:
            seen.add(b.lower())
            bullets.append(b)
    if not bullets:
        for _, t in items[:max_bullets]:
            b = shorten(t, 150)
            if b:
                bullets.append(b)
    return bullets


def local_structure(blocks, topic, max_bullets):
    doc_title, sections = build_sections(blocks)
    all_text = " ".join(t for _, t, _ in blocks)
    total_words = sum(len(t.split()) for _, items in sections for _, t in items)
    if len(sections) == 1 and total_words >= 170:
        only_title, items = sections[0]
        if only_title and not doc_title:
            doc_title = only_title
        sections = chunk_items(items)
    has_titles = any(t for t, _ in sections)

    out = []
    used_kw = set()
    for i, (t, items) in enumerate(sections):
        bullets = section_to_bullets(items, max_bullets)
        if not bullets:
            continue
        if not t:
            if has_titles:
                t = "Introduction"
            else:
                t = title_from_keywords(" ".join(x for _, x in items), 2, used_kw)
                used_kw.update(re.findall(r"[^\W\d_][\w'-]*", t))
        out.append({"title": t, "bullets": bullets})
    if not out:
        raise ValueError("Could not find enough usable content in the text.")

    title = (topic or "").strip() or doc_title
    if not title:
        title = title_from_keywords(all_text, 3) + ": An Overview"
    return {"title": title, "subtitle": "", "sections": out, "closing": ""}


# ---------------------------------------------------------------------
# Optional LLM structuring (OpenAI / Gemini through plain REST calls)
# ---------------------------------------------------------------------
LLM_SYSTEM = "You are an expert presentation designer. You always answer with valid JSON only."


def llm_prompt(text, topic):
    topic_line = f'The presentation title MUST be: "{topic}".' if topic else \
        "Invent a short, specific presentation title."
    return f"""Turn the CONTENT below into a slide deck outline.
{topic_line}
Return ONLY a JSON object with exactly this shape:
{{"title": "...", "subtitle": "...", "sections": [{{"title": "...", "bullets": ["...", "..."]}}], "closing_message": "..."}}

Rules:
- 3 to 12 sections, in a logical order.
- Every section title must be UNIQUE, specific and descriptive (max 60 characters). Never reuse a title.
- 3 to 8 bullets per section; each bullet is a concise phrase or sentence of at most 20 words.
- Use only facts present in the content. Do not invent numbers or claims.
- Write in the same language as the content.
- closing_message is one short sentence summarising the main takeaway.

CONTENT:
\"\"\"
{text}
\"\"\""""


def parse_json_loose(s):
    s = (s or "").strip()
    s = re.sub(r"^```(?:json)?\s*|\s*```$", "", s).strip()
    try:
        return json.loads(s)
    except json.JSONDecodeError:
        a, b = s.find("{"), s.rfind("}")
        if a >= 0 and b > a:
            return json.loads(s[a:b + 1])
        raise


def call_llm(provider, api_key, model, prompt):
    if provider == "openai":
        r = requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            json={"model": model, "temperature": 0.3,
                  "response_format": {"type": "json_object"},
                  "messages": [{"role": "system", "content": LLM_SYSTEM},
                               {"role": "user", "content": prompt}]},
            timeout=120)
        if r.status_code != 200:
            raise RuntimeError(f"OpenAI error {r.status_code}: {r.text[:250]}")
        return r.json()["choices"][0]["message"]["content"]
    r = requests.post(
        f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
        headers={"x-goog-api-key": api_key, "Content-Type": "application/json"},
        json={"contents": [{"parts": [{"text": LLM_SYSTEM + "\n\n" + prompt}]}],
              "generationConfig": {"temperature": 0.3, "responseMimeType": "application/json"}},
        timeout=120)
    if r.status_code != 200:
        raise RuntimeError(f"Gemini error {r.status_code}: {r.text[:250]}")
    return r.json()["candidates"][0]["content"]["parts"][0]["text"]


def llm_structure(blocks, topic, provider, api_key, model):
    text = blocks_to_text(blocks)
    truncated = len(text) > 14000
    if truncated:
        text = text[:14000]
    default_model = "gpt-4o-mini" if provider == "openai" else "gemini-2.5-flash"
    raw = call_llm(provider, api_key, (model or "").strip() or default_model, llm_prompt(text, topic))
    data = parse_json_loose(raw)
    if not isinstance(data, dict) or not isinstance(data.get("sections"), list):
        raise ValueError("the model returned an unexpected JSON structure")
    secs = []
    for s in data["sections"]:
        if not isinstance(s, dict):
            continue
        t = clean_text(str(s.get("title", "")))
        bl = s.get("bullets") or s.get("points") or []
        if isinstance(bl, str):
            bl = [bl]
        bl = [shorten(str(b), 170) for b in bl if str(b).strip()]
        if t and bl:
            secs.append({"title": t, "bullets": bl})
    if not secs:
        raise ValueError("the model returned no usable sections")
    title = (topic or "").strip() or clean_text(str(data.get("title") or "")) or "Presentation"
    deck = {"title": title,
            "subtitle": clean_text(str(data.get("subtitle") or "")),
            "sections": secs,
            "closing": clean_text(str(data.get("closing_message") or ""))}
    return deck, truncated


# ---------------------------------------------------------------------
# Slide planning (pagination + guaranteed-unique titles)
# ---------------------------------------------------------------------
AGENDA_PER_SLIDE = 8


def paginate(bullets, max_n, max_chars=650):
    bullets = [b for b in bullets if b]
    if not bullets:
        return []
    pages, cur_n, cur_c = 0, 0, 0
    for b in bullets:
        if cur_n and (cur_n >= max_n or cur_c + len(b) > max_chars):
            pages += 1
            cur_n, cur_c = 0, 0
        cur_n += 1
        cur_c += len(b)
    pages += 1
    per = math.ceil(len(bullets) / pages)
    return [bullets[i:i + per] for i in range(0, len(bullets), per)]


def plan_slides(deck, max_bullets):
    sections = [s for s in deck["sections"] if s["bullets"]]
    used = set()

    def register(title):
        base = title
        cand, k = base, 2
        while cand.lower() in used:
            cand = f"{base} ({k})"
            k += 1
        used.add(cand.lower())
        return cand

    deck_title = shorten_title(deck["title"], 95)
    used.add(deck_title.lower())
    used.add("thank you")

    n_agenda = math.ceil(len(sections) / AGENDA_PER_SLIDE)
    agenda_titles = []
    for i in range(n_agenda):
        agenda_titles.append(register("Agenda" if n_agenda == 1 else f"Agenda (Part {i + 1})"))

    content, bases = [], []
    total_secs = len(sections)
    for si, sec in enumerate(sections, 1):
        base = register(shorten_title(sec["title"], 60))
        bases.append(base)
        pages = paginate(sec["bullets"], max_bullets)
        for pi, page in enumerate(pages, 1):
            title = base if len(pages) == 1 else register(f"{base} (Part {pi})")
            content.append({"kind": "content", "title": title, "bullets": page,
                            "kicker": f"TOPIC {si} OF {total_secs}"})

    plan = [{"kind": "title", "title": deck_title,
             "subtitle": deck.get("subtitle") or f"An overview of {total_secs} key topic{'s' if total_secs != 1 else ''}"}]
    for i in range(n_agenda):
        chunk = bases[i * AGENDA_PER_SLIDE:(i + 1) * AGENDA_PER_SLIDE]
        items = [(i * AGENDA_PER_SLIDE + j + 1, shorten_title(b, 60)) for j, b in enumerate(chunk)]
        plan.append({"kind": "agenda", "title": agenda_titles[i], "items": items})
    plan.extend(content)

    msg = deck.get("closing") or ""
    if not msg:
        msg = "Topics covered: " + " \u2022 ".join(shorten_title(b, 40) for b in bases[:6])
        if len(bases) > 6:
            msg += " \u2022 \u2026"
    plan.append({"kind": "closing", "title": "Thank You", "message": msg})
    return plan, deck_title


# ---------------------------------------------------------------------
# PPTX rendering
# ---------------------------------------------------------------------
def rgb(h):
    return RGBColor.from_string(h)


def set_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb(color)


def rect(slide, x, y, w, h, color, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = rgb(color)
    s.line.fill.background()
    s.shadow.inherit = False
    return s


def add_text(slide, x, y, w, h, text, size, color, font, bold=False, italic=False,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    f = r.font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.name = font
    f.color.rgb = rgb(color)
    return tb


def fit_single(text, width_in, height_in, max_pt, min_pt):
    for size in range(max_pt, min_pt - 1, -2):
        cpl = max(8, int(width_in * 72 / (size * 0.55)))
        if math.ceil(len(text) / cpl) * size * 1.2 <= height_in * 72:
            return size
    return min_pt


def fit_bullets(bullets, width_in, height_in, max_pt=26, min_pt=15):
    for size in range(max_pt, min_pt - 1, -1):
        cpl = max(10, int((width_in - 0.4) * 72 / (size * 0.52)))
        lines = sum(max(1, math.ceil(len(b) / cpl)) for b in bullets)
        need = lines * size * 1.25 + len(bullets) * size * 0.6
        if need <= height_in * 72:
            return size
    return min_pt


def add_bullets(slide, x, y, w, h, bullets, size, color, accent, font):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.TOP
    indent = int(Inches(0.4))
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(size * 0.6)
        p.line_spacing = 1.05
        r = p.add_run()
        r.text = b
        r.font.size = Pt(size)
        r.font.name = font
        r.font.color.rgb = rgb(color)
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", str(indent))
        pPr.set("indent", str(-indent))
        buClr = etree.SubElement(pPr, pptx_qn("a:buClr"))
        etree.SubElement(buClr, pptx_qn("a:srgbClr")).set("val", accent)
        etree.SubElement(pPr, pptx_qn("a:buFont")).set("typeface", "Arial")
        etree.SubElement(pPr, pptx_qn("a:buChar")).set("char", "\u2022")
    return tb


def add_badge(slide, x, y, d, num, th):
    s = rect(slide, x, y, d, d, th["accent"], MSO_SHAPE.OVAL)
    tf = s.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = str(num)
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.name = th["font_head"]
    r.font.color.rgb = rgb(th["badge_fg"])


def footer(slide, th, deck_title, n):
    add_text(slide, 0.9, 6.98, 8.0, 0.3, shorten_title(deck_title, 60), 11, th["muted"], th["font_body"])
    add_text(slide, 11.6, 6.98, 0.8, 0.3, str(n), 11, th["muted"], th["font_body"], align=PP_ALIGN.RIGHT)


def render(plan, deck_title, theme_name, show_date, out_path):
    th = THEMES[theme_name]
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    for n, sd in enumerate(plan, 1):
        s = prs.slides.add_slide(blank)
        kind = sd["kind"]

        if kind == "title":
            set_bg(s, th["dark_bg"])
            rect(s, 0, 0, 0.35, 7.5, th["accent"])
            size = fit_single(sd["title"], 11.1, 2.6, 54, 30)
            add_text(s, 1.1, 1.6, 11.1, 2.7, sd["title"], size, th["on_dark"], th["font_head"],
                     bold=True, anchor=MSO_ANCHOR.BOTTOM)
            rect(s, 1.1, 4.55, 1.6, 0.08, th["accent"])
            add_text(s, 1.1, 4.85, 11.1, 1.0, sd["subtitle"], 22, th["on_dark_muted"], th["font_body"])
            if show_date:
                d = datetime.date.today()
                add_text(s, 1.1, 6.3, 6.0, 0.4, f"{d:%B} {d.day}, {d.year}", 14,
                         th["on_dark_muted"], th["font_body"])

        elif kind == "agenda":
            set_bg(s, th["bg"])
            rect(s, 0, 0, 0.25, 7.5, th["accent"])
            add_text(s, 0.9, 0.45, 8, 0.3, "WHAT WE WILL COVER", 12, th["accent"], th["font_body"], bold=True)
            add_text(s, 0.9, 0.75, 11.5, 1.0, sd["title"], 36, th["heading"], th["font_head"],
                     bold=True, anchor=MSO_ANCHOR.BOTTOM)
            rect(s, 0.9, 1.85, 1.2, 0.06, th["accent"])
            items = sd["items"]
            two_cols = len(items) > 4
            col_w = 5.4 if two_cols else 10.5
            per_col = 4 if two_cols else len(items)
            for j, (num, text) in enumerate(items):
                col, row = divmod(j, per_col)
                x = 0.9 + col * 6.1
                y = 2.3 + row * 1.15
                add_badge(s, x, y, 0.6, num, th)
                add_text(s, x + 0.85, y - 0.1, col_w - 0.9, 0.8, text, 20 if two_cols else 22,
                         th["text"], th["font_body"], anchor=MSO_ANCHOR.MIDDLE)
            footer(s, th, deck_title, n)

        elif kind == "content":
            set_bg(s, th["bg"])
            rect(s, 0, 0, 0.25, 7.5, th["accent"])
            add_text(s, 0.9, 0.42, 8, 0.3, sd["kicker"], 12, th["accent"], th["font_body"], bold=True)
            tsize = fit_single(sd["title"], 11.5, 1.0, 34, 22)
            add_text(s, 0.9, 0.75, 11.5, 1.0, sd["title"], tsize, th["heading"], th["font_head"],
                     bold=True, anchor=MSO_ANCHOR.BOTTOM)
            rect(s, 0.9, 1.85, 1.2, 0.06, th["accent"])
            card = rect(s, 0.9, 2.1, 11.5, 4.65, th["card"], MSO_SHAPE.ROUNDED_RECTANGLE)
            card.adjustments[0] = 0.03
            bsize = fit_bullets(sd["bullets"], 10.8, 4.1)
            add_bullets(s, 1.25, 2.4, 10.8, 4.1, sd["bullets"], bsize, th["text"], th["accent"], th["font_body"])
            footer(s, th, deck_title, n)

        else:  # closing
            set_bg(s, th["dark_bg"])
            rect(s, 0, 0, 0.35, 7.5, th["accent"])
            add_text(s, 1.1, 1.5, 11.1, 1.6, sd["title"], 60, th["on_dark"], th["font_head"],
                     bold=True, anchor=MSO_ANCHOR.BOTTOM)
            rect(s, 1.1, 3.4, 1.6, 0.08, th["accent"])
            add_text(s, 1.1, 3.75, 11.1, 0.6, "Questions & Discussion", 26, th["on_dark"], th["font_body"])
            add_text(s, 1.1, 4.7, 11.1, 1.6, sd["message"], 16, th["on_dark_muted"], th["font_body"])

    prs.save(out_path)


# ---------------------------------------------------------------------
# Gradio callback
# ---------------------------------------------------------------------
MODE_LOCAL = "Local (no API key needed)"
MODE_OPENAI = "OpenAI (API key)"
MODE_GEMINI = "Google Gemini (API key)"


def generate(topic, pasted, file, theme_name, mode, api_key, model, max_bullets, show_date):
    try:
        blocks = []
        if file:
            path = getattr(file, "name", file)
            if not str(path).lower().endswith(".docx"):
                return "\u274c Please upload a **.docx** Word file.", "", None
            blocks += parse_docx(path)
        if pasted and pasted.strip():
            blocks += parse_plain_text(pasted)
        if sum(len(t.split()) for _, t, _ in blocks) < 8:
            return "\u274c Please paste some text (or upload a .docx) with at least a few sentences.", "", None

        topic = (topic or "").strip()
        max_bullets = int(max_bullets)
        notes, deck, used_mode = [], None, "Local"

        if mode in (MODE_OPENAI, MODE_GEMINI):
            key = (api_key or "").strip()
            if not key:
                notes.append("\u26a0\ufe0f No API key entered \u2013 used local mode instead.")
            else:
                try:
                    provider = "openai" if mode == MODE_OPENAI else "gemini"
                    deck, truncated = llm_structure(blocks, topic, provider, key, model)
                    used_mode = "OpenAI" if provider == "openai" else "Gemini"
                    if truncated:
                        notes.append("\u2139\ufe0f Input was long; the AI saw the first ~14,000 characters.")
                except Exception as e:
                    notes.append(f"\u26a0\ufe0f AI mode failed ({str(e)[:200]}) \u2013 fell back to local mode.")
        if deck is None:
            deck = local_structure(blocks, topic, max_bullets)

        plan, deck_title = plan_slides(deck, max_bullets)
        titles = [p["title"].lower() for p in plan]
        assert len(titles) == len(set(titles)), "internal error: duplicate slide titles"

        slug = re.sub(r"[^\w\-]+", "_", deck_title).strip("_")[:40] or "presentation"
        out_path = os.path.join(tempfile.mkdtemp(prefix="deck_"), slug + ".pptx")
        render(plan, deck_title, theme_name, bool(show_date), out_path)

        outline = "\n".join(f"{i:02d}  [{p['kind']:<7}] {p['title']}" for i, p in enumerate(plan, 1))
        status = (f"\u2705 Created **{len(plan)} slides** ({used_mode} mode, theme: {theme_name}). "
                  "Download the file below.")
        if notes:
            status += "\n\n" + "\n\n".join(notes)
        return status, outline, out_path
    except Exception as e:
        return f"\u274c Something went wrong: {e}", "", None


# ---------------------------------------------------------------------
# Web UI
# ---------------------------------------------------------------------
with gr.Blocks(title="Text to PowerPoint") as demo:
    gr.Markdown(
        "# \U0001F4CA Text / Word \u2192 PowerPoint\n"
        "Paste text or upload a `.docx`. You get a title slide, an agenda, topic slides with "
        "unique titles, and a closing slide.")
    with gr.Row():
        with gr.Column(scale=3):
            topic_in = gr.Textbox(label="Presentation title / topic (optional \u2013 auto-detected if empty)")
            text_in = gr.Textbox(label="Paste your text here", lines=14,
                                 placeholder="Paste an article, notes, a report\u2026 Headings, bullet lists "
                                             "and markdown (#, ##) are detected automatically.")
            file_in = gr.File(label="\u2026or upload a Word file (.docx)", file_types=[".docx"], type="filepath")
        with gr.Column(scale=2):
            theme_in = gr.Dropdown(list(THEMES.keys()), value="Modern Blue", label="Color theme")
            mode_in = gr.Radio([MODE_LOCAL, MODE_OPENAI, MODE_GEMINI], value=MODE_LOCAL,
                               label="Structuring mode")
            key_in = gr.Textbox(label="API key (only for the AI modes)", type="password")
            model_in = gr.Textbox(label="Model (optional)", placeholder="gpt-4o-mini  /  gemini-2.5-flash")
            bullets_in = gr.Slider(3, 7, value=5, step=1, label="Max bullets per slide")
            date_in = gr.Checkbox(value=True, label="Show today's date on the title slide")
            go = gr.Button("Generate presentation", variant="primary")
    status_out = gr.Markdown()
    outline_out = gr.Textbox(label="Slide outline (all titles are unique)", lines=12, interactive=False)
    file_out = gr.File(label="Download your .pptx")
    go.click(generate,
             inputs=[topic_in, text_in, file_in, theme_in, mode_in, key_in, model_in, bullets_in, date_in],
             outputs=[status_out, outline_out, file_out])

demo.launch(share=True, debug=True)   # prints a public https://xxxx.gradio.live link
```
