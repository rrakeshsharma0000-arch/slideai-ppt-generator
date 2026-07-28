import os
import io
import gc
import json
import re
import time
import traceback

import requests as http_requests
from flask import Flask, render_template, request, send_file, jsonify
from dotenv import load_dotenv
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

load_dotenv()

app = Flask(__name__)

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

# ─── Palette ──────────────────────────────────────────────────────────────────

C = {
    "primary":   RGBColor(0x2B, 0x45, 0x90),
    "secondary": RGBColor(0x4A, 0x90, 0xD9),
    "accent":    RGBColor(0xF5, 0xA6, 0x23),
    "dark":      RGBColor(0x2C, 0x2C, 0x2C),
    "white":     RGBColor(0xFF, 0xFF, 0xFF),
    "light":     RGBColor(0xF5, 0xF7, 0xFA),
    "success":   RGBColor(0x27, 0xAE, 0x60),
    "purple":    RGBColor(0x6C, 0x3D, 0xC4),
    "panel":     RGBColor(0xE8, 0xEF, 0xFB),
    "darkblue":  RGBColor(0x1A, 0x2F, 0x6B),
    "midblue":   RGBColor(0x3B, 0x5A, 0xAA),
    "lightblue": RGBColor(0xB8, 0xCC, 0xF0),
}

W = Inches(13.33)
H = Inches(7.5)

# ─── AI Content Generation ────────────────────────────────────────────────────

GEMINI_REST_URL = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}"


def _call_gemini_rest(prompt, model):
    url = GEMINI_REST_URL.format(model=model, key=GEMINI_API_KEY)
    payload = {"contents": [{"parts": [{"text": prompt}]}]}
    resp = http_requests.post(url, json=payload, timeout=80)
    if resp.status_code == 429:
        raise Exception(f"429 RESOURCE_EXHAUSTED {resp.text}")
    resp.raise_for_status()
    return resp.json()["candidates"][0]["content"]["parts"][0]["text"]


def generate_ppt_content(topic, audience, num_slides, additional_info):
    prompt = f"""Create a comprehensive, engaging PowerPoint presentation for:

Topic: {topic}
Audience: {audience}
Content slides requested: {num_slides}
Extra context: {additional_info or "None"}

Return ONLY a JSON object (no markdown, no extra text) with this exact structure:
{{
  "title": "Catchy Presentation Title",
  "subtitle": "Descriptive subtitle in 8-12 words",
  "agenda": ["Topic area 1", "Topic area 2", "Topic area 3", "Topic area 4", "Topic area 5"],
  "slides": [
    {{
      "title": "Slide Title",
      "type": "content",
      "emoji": "📌",
      "key_points": [
        "Concise point, max 15 words",
        "Another clear point",
        "Third supporting point",
        "Fourth point if needed"
      ],
      "example": "One concrete real-world example that makes this tangible",
      "visual_description": "Chart or diagram idea for this slide",
      "speaker_notes": "2-3 sentence presenter guide for this slide"
    }}
  ],
  "key_takeaways": [
    "Most important thing to remember",
    "Second key learning",
    "Third key learning",
    "Fourth key learning"
  ],
  "discussion_questions": [
    "Thought-provoking question 1?",
    "Question 2 that encourages reflection?",
    "Question 3?"
  ]
}}

Slide type rules:
- First 2 slides: type = "content" (foundational concepts)
- Every 3rd slide: type = "example" (real-world case study)
- One slide: type = "activity" (interactive exercise — also populate the "activity" field with a group exercise)
- Remaining slides: type = "content"
- Total: exactly {num_slides} slides in the "slides" array

Make it educational, clear, and suitable for {audience}.
Emoji should match the slide topic visually."""

    models = ["gemini-2.0-flash", "gemini-2.5-flash", "gemini-2.0-flash-lite", "gemini-2.5-flash-lite"]
    last_err = None
    for model in models:
        for attempt in range(2):
            try:
                text = _call_gemini_rest(prompt, model)
                text = text.strip()
                text = re.sub(r"^```(?:json)?\s*", "", text)
                text = re.sub(r"\s*```$", "", text)
                return json.loads(text)
            except Exception as e:
                last_err = e
                err_str = str(e)
                if "404" in err_str or "Not Found" in err_str:
                    break  # model not available — skip to next
                if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                    if "limit: 0" in err_str or "limit_value: 0" in err_str:
                        break  # quota disabled for this model entirely — skip it immediately
                    delay = 20
                    m = re.search(r"retry.*?(\d+)s", err_str)
                    if m:
                        delay = min(int(m.group(1)) + 2, 30)
                    if attempt == 0:
                        time.sleep(delay)
                        continue  # retry same model
                    break  # try next model
                raise  # non-rate-limit errors bubble up immediately
    err_str = str(last_err)
    if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
        raise Exception("API quota exhausted. Please try again later or contact the administrator.")
    raise last_err

# ─── Slide Helpers ────────────────────────────────────────────────────────────

def _rect(slide, x, y, w, h, fill_color, line_color=None):
    shape = slide.shapes.add_shape(1, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
    else:
        shape.line.fill.background()
    return shape



def _textbox(slide, x, y, w, h, text, size, color, bold=False, italic=False,
             align=PP_ALIGN.LEFT, wrap=True):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.italic = italic
    p.alignment = align
    return tb


def _header(slide, title, bg=None, fg=None, height=Inches(1.4), font_size=26):
    bg = bg or C["primary"]
    fg = fg or C["white"]
    _rect(slide, 0, 0, W, height, bg)
    _textbox(slide, Inches(0.45), Inches(0.2), Inches(12.4), height - Inches(0.2),
             title, font_size, fg, bold=True)


def _notes(slide, text):
    if text:
        slide.notes_slide.notes_text_frame.text = text

# ─── Efficient multi-paragraph helper ─────────────────────────────────────────

def _bullets_textbox(slide, x, y, w, h, items, size, color, prefix="  •  ", space_after=Pt(6)):
    """Single textbox containing all bullet items — far fewer shapes than one-per-bullet."""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"{prefix}{item}"
        p.font.size = size
        p.font.color.rgb = color
        p.space_after = space_after
    return tb

# ─── Individual Slide Builders ────────────────────────────────────────────────

def _title_slide(prs, data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # 4 shapes total
    _rect(slide, 0, 0, W, H, C["primary"])
    _rect(slide, 0, Inches(5.4), W, Inches(2.1), C["darkblue"])
    _rect(slide, Inches(1.0), Inches(5.25), Inches(11.33), Inches(0.07), C["accent"])
    _textbox(slide, Inches(1.0), Inches(1.4), Inches(10.5), Inches(2.8),
             data["title"], 44, C["white"], bold=True, wrap=True)
    _textbox(slide, Inches(1.0), Inches(4.3), Inches(10.5), Inches(0.9),
             data["subtitle"], 20, C["lightblue"], wrap=True)
    _textbox(slide, Inches(1.0), Inches(6.5), Inches(5.0), Inches(0.6),
             "✦ Interactive Presentation", 13, C["accent"], bold=True)


def _agenda_slide(prs, data):
    # 4 shapes total (was 2 + 2×N)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(slide, 0, 0, W, H, C["white"])
    _header(slide, "📋  What We'll Cover Today")
    _rect(slide, Inches(0.5), Inches(1.6), Inches(0.07), Inches(5.5), C["accent"])
    items = [f"{i+1}.  {item}" for i, item in enumerate(data.get("agenda", [])[:7])]
    _bullets_textbox(slide, Inches(1.0), Inches(1.7), Inches(11.5), Inches(5.4),
                     items, Pt(19), C["dark"], prefix="", space_after=Pt(10))
    _notes(slide, "Walk through the agenda. Let students know what they'll learn and why it matters.")


def _content_slide(prs, slide_data):
    # 5 shapes total (was 2 + 2×N + 2)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(slide, 0, 0, W, H, C["white"])
    _header(slide, f"{slide_data.get('emoji','📌')}  {slide_data['title']}")
    points = slide_data.get("key_points", [])
    if points:
        _bullets_textbox(slide, Inches(0.6), Inches(1.6), Inches(12.1), Inches(5.0),
                         points[:5], Pt(19), C["dark"])
    example = slide_data.get("example", "")
    if example:
        _textbox(slide, Inches(0.5), Inches(6.5), Inches(12.33), Inches(0.65),
                 f"💡  {example}", 13, C["secondary"], italic=True, wrap=True)
    _rect(slide, 0, Inches(7.35), W, Inches(0.15), C["secondary"])
    _notes(slide, slide_data.get("speaker_notes", ""))


def _example_slide(prs, slide_data):
    # 5 shapes total (was 5 + N)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(slide, 0, 0, W, H, C["light"])
    _header(slide, f"💡  {slide_data['title']} — Example", bg=C["secondary"])
    example_text = slide_data.get("example", "")
    _rect(slide, Inches(0.5), Inches(1.6), Inches(12.33), Inches(3.0), C["panel"],
          line_color=C["secondary"])
    _textbox(slide, Inches(0.8), Inches(1.75), Inches(11.73), Inches(2.75),
             example_text, 17, C["dark"], wrap=True)
    points = slide_data.get("key_points", [])
    if points:
        _bullets_textbox(slide, Inches(0.6), Inches(4.8), Inches(12.1), Inches(2.4),
                         points[:3], Pt(15), C["dark"], prefix="✅  ")
    _notes(slide, slide_data.get("speaker_notes", ""))


def _activity_slide(prs, slide_data):
    # 5 shapes total (was 7)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(slide, 0, 0, W, H, C["primary"])
    _textbox(slide, Inches(0.5), Inches(0.3), Inches(4.0), Inches(0.6),
             "🎯  ACTIVITY", 16, C["accent"], bold=True)
    _textbox(slide, Inches(0.5), Inches(1.0), Inches(12.0), Inches(1.4),
             slide_data["title"], 34, C["white"], bold=True, wrap=True)
    activity_text = slide_data.get("activity") or slide_data.get("example", "Activity coming soon.")
    _rect(slide, Inches(0.5), Inches(2.55), Inches(12.33), Inches(3.5), C["darkblue"])
    _textbox(slide, Inches(0.85), Inches(2.75), Inches(11.63), Inches(3.1),
             activity_text, 19, C["white"], wrap=True)
    _textbox(slide, Inches(0.5), Inches(6.3), Inches(12.0), Inches(0.7),
             "⏱  Take 3 minutes — then share with the group", 14, C["accent"])
    _notes(slide, slide_data.get("speaker_notes", ""))


def _summary_slide(prs, data):
    # 4 shapes total (was 2 + 2×N)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(slide, 0, 0, W, H, C["white"])
    _header(slide, "🎯  Key Takeaways", bg=C["success"])
    takeaways = [f"{i+1}.  {tk}" for i, tk in enumerate(data.get("key_takeaways", [])[:5])]
    _bullets_textbox(slide, Inches(0.6), Inches(1.6), Inches(12.1), Inches(4.8),
                     takeaways, Pt(19), C["dark"], prefix="", space_after=Pt(12))
    questions = data.get("discussion_questions", [])
    if questions:
        _textbox(slide, Inches(0.5), Inches(6.65), Inches(12.33), Inches(0.55),
                 f"💬  Discuss: {questions[0]}", 14, C["secondary"], italic=True, wrap=True)
    _notes(slide, "Summarize the key points. Ask students to reflect on what surprised them most.")


def _qa_slide(prs, data):
    # 4 shapes total (was 4 + N)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(slide, 0, 0, W, H, C["primary"])
    _textbox(slide, Inches(1.5), Inches(1.5), Inches(9.0), Inches(2.5),
             "Questions?", 62, C["white"], bold=True, align=PP_ALIGN.CENTER)
    _textbox(slide, Inches(1.5), Inches(4.0), Inches(9.0), Inches(0.9),
             "Let's discuss! 🙋", 28, C["accent"], align=PP_ALIGN.CENTER)
    questions = data.get("discussion_questions", [])
    if questions:
        _bullets_textbox(slide, Inches(1.0), Inches(5.1), Inches(11.33), Inches(1.8),
                         questions[:3], Pt(14), C["lightblue"],
                         prefix="  •  ", space_after=Pt(4))
    _notes(slide, "Open the floor for questions. Use discussion prompts above to spark conversation if needed.")

# ─── Assemble Presentation ────────────────────────────────────────────────────

def create_presentation(data):
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H

    _title_slide(prs, data)
    _agenda_slide(prs, data)

    for s in data.get("slides", []):
        t = s.get("type", "content")
        if t == "example":
            _example_slide(prs, s)
        elif t == "activity":
            _activity_slide(prs, s)
        else:
            _content_slide(prs, s)

    _summary_slide(prs, data)
    _qa_slide(prs, data)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf

# ─── Routes ───────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    try:
        if not GEMINI_API_KEY:
            return jsonify({"error": "GEMINI_API_KEY is not set. Add it to your .env file."}), 500

        topic = (request.form.get("topic") or "").strip()[:200]
        if not topic:
            return jsonify({"error": "Please enter a topic."}), 400

        audience = request.form.get("audience", "college students").strip()
        num_slides = max(4, min(int(request.form.get("num_slides", 10)), 30))
        # Cap additional_info to avoid inflating the prompt and causing timeouts
        additional_info = (request.form.get("additional_info") or "").strip()[:300]

        ppt_data = generate_ppt_content(topic, audience, num_slides, additional_info)
        safe = re.sub(r"[^a-zA-Z0-9 ]", "", ppt_data.get("title", topic))[:40]
        filename = safe.strip().replace(" ", "_") + ".pptx"

        buf = create_presentation(ppt_data)
        del ppt_data
        gc.collect()

        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name=filename,
        )

    except json.JSONDecodeError:
        return jsonify({"error": "AI returned an unexpected response. Please try again."}), 500
    except Exception as e:
        traceback.print_exc()
        msg = str(e)
        if "429" in msg or "RESOURCE_EXHAUSTED" in msg:
            return jsonify({"error": "API rate limit reached. Please wait a minute and try again."}), 429
        if "quota" in msg.lower():
            return jsonify({"error": "Daily API quota exceeded. Try again tomorrow or use a different API key."}), 429
        if "API_KEY" in msg or "api key" in msg.lower():
            return jsonify({"error": "Invalid or missing Gemini API key."}), 500
        return jsonify({"error": "Something went wrong generating your presentation. Please try again."}), 500


@app.errorhandler(500)
def internal_error(e):
    return jsonify({"error": f"Internal server error: {str(e)}"}), 500


@app.errorhandler(Exception)
def unhandled_exception(e):
    traceback.print_exc()
    return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=False, port=port)
