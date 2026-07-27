import os
import io
import json
import re
import traceback

from flask import Flask, render_template, request, send_file, jsonify
from dotenv import load_dotenv
from google import genai
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

def generate_ppt_content(topic, audience, num_slides, additional_info):
    client = genai.Client(api_key=GEMINI_API_KEY)

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

    resp = client.models.generate_content(model="gemini-flash-latest", contents=prompt)
    text = resp.text.strip()
    # Strip markdown fences if present
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    return json.loads(text)

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


def _oval(slide, x, y, w, h, fill_color):
    shape = slide.shapes.add_shape(9, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
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

# ─── Individual Slide Builders ────────────────────────────────────────────────

def _title_slide(prs, data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    _rect(slide, 0, 0, W, H, C["primary"])
    _rect(slide, 0, Inches(5.4), W, Inches(2.1), C["darkblue"])
    _rect(slide, Inches(1.0), Inches(5.25), Inches(11.33), Inches(0.07), C["accent"])

    # Decorative circles (background art)
    _oval(slide, Inches(10.5), Inches(-0.5), Inches(3.5), Inches(3.5), C["midblue"])
    _oval(slide, Inches(11.5), Inches(5.5), Inches(2.5), Inches(2.5), C["midblue"])

    _textbox(slide, Inches(1.0), Inches(1.4), Inches(10.5), Inches(2.8),
             data["title"], 44, C["white"], bold=True, wrap=True)
    _textbox(slide, Inches(1.0), Inches(4.3), Inches(10.5), Inches(0.9),
             data["subtitle"], 20, C["lightblue"], wrap=True)
    _textbox(slide, Inches(1.0), Inches(6.5), Inches(5.0), Inches(0.6),
             "✦ Interactive Presentation", 13, C["accent"], bold=True)


def _agenda_slide(prs, data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(slide, 0, 0, W, H, C["white"])
    _header(slide, "📋  What We'll Cover Today")
    _rect(slide, Inches(0.5), Inches(1.6), Inches(0.07), Inches(5.5), C["accent"])

    for i, item in enumerate(data.get("agenda", [])[:7]):
        y = Inches(1.7 + i * 0.72)
        num = _oval(slide, Inches(0.75), y, Inches(0.44), Inches(0.44), C["primary"])
        num_tf = num.text_frame
        p = num_tf.paragraphs[0]
        p.text = str(i + 1)
        p.font.color.rgb = C["white"]
        p.font.size = Pt(12)
        p.font.bold = True
        p.alignment = PP_ALIGN.CENTER
        _textbox(slide, Inches(1.4), y + Inches(0.05), Inches(11.0), Inches(0.5),
                 item, 18, C["dark"])

    _notes(slide, "Walk through the agenda. Let students know what they'll learn and why it matters.")


def _content_slide(prs, slide_data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(slide, 0, 0, W, H, C["white"])
    _header(slide, slide_data["title"])

    emoji = slide_data.get("emoji", "📌")
    _textbox(slide, Inches(11.3), Inches(0.05), Inches(1.8), Inches(1.3),
             emoji, 40, C["white"], align=PP_ALIGN.CENTER)

    points = slide_data.get("key_points", [])
    for i, point in enumerate(points[:5]):
        y = Inches(1.65 + i * 0.88)
        dot = _oval(slide, Inches(0.4), y + Inches(0.12), Inches(0.2), Inches(0.2), C["accent"])
        _textbox(slide, Inches(0.8), y, Inches(12.1), Inches(0.75),
                 point, 19, C["dark"], wrap=True)

    example = slide_data.get("example", "")
    if example:
        box_y = Inches(1.65 + len(points[:5]) * 0.88 + 0.1)
        if box_y < Inches(6.5):
            _rect(slide, Inches(0.4), box_y, Inches(12.5), Inches(0.06), C["secondary"])
            _textbox(slide, Inches(0.4), box_y + Inches(0.12), Inches(12.5), Inches(0.7),
                     f"💡 {example}", 14, C["secondary"], italic=True, wrap=True)

    _rect(slide, 0, Inches(7.35), W, Inches(0.15), C["secondary"])
    _notes(slide, slide_data.get("speaker_notes", ""))


def _example_slide(prs, slide_data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(slide, 0, 0, W, H, C["light"])
    _header(slide, f"💡  {slide_data['title']}", bg=C["secondary"])

    # EXAMPLE tag
    tag = _oval(slide, Inches(11.1), Inches(0.22), Inches(1.9), Inches(0.42), C["accent"])
    tag_tf = tag.text_frame
    p = tag_tf.paragraphs[0]
    p.text = "EXAMPLE"
    p.font.color.rgb = C["white"]
    p.font.size = Pt(12)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER

    example_text = slide_data.get("example", "Example coming soon.")
    _rect(slide, Inches(0.5), Inches(1.6), Inches(12.33), Inches(3.0), C["panel"],
          line_color=C["secondary"])
    _textbox(slide, Inches(0.8), Inches(1.75), Inches(11.73), Inches(2.75),
             example_text, 17, C["dark"], wrap=True)

    points = slide_data.get("key_points", [])
    if points:
        _textbox(slide, Inches(0.5), Inches(4.8), Inches(4.0), Inches(0.5),
                 "✅  Key Points:", 16, C["primary"], bold=True)
        for i, pt in enumerate(points[:3]):
            _textbox(slide, Inches(0.6), Inches(5.4 + i * 0.5), Inches(12.1), Inches(0.45),
                     f"  •  {pt}", 15, C["dark"])

    _notes(slide, slide_data.get("speaker_notes", ""))


def _activity_slide(prs, slide_data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(slide, 0, 0, W, H, C["primary"])
    _oval(slide, Inches(10.0), Inches(0.3), Inches(3.0), Inches(3.0), C["midblue"])
    _oval(slide, Inches(-0.5), Inches(5.5), Inches(2.5), Inches(2.5), C["midblue"])

    _textbox(slide, Inches(0.5), Inches(0.3), Inches(4.0), Inches(0.6),
             "🎯  ACTIVITY", 16, C["accent"], bold=True)
    _textbox(slide, Inches(0.5), Inches(1.0), Inches(12.0), Inches(1.4),
             slide_data["title"], 34, C["white"], bold=True, wrap=True)

    activity_text = slide_data.get("activity") or slide_data.get("example", "Activity coming soon.")
    _rect(slide, Inches(0.5), Inches(2.55), Inches(12.33), Inches(3.3), C["darkblue"])
    _textbox(slide, Inches(0.85), Inches(2.75), Inches(11.63), Inches(2.9),
             activity_text, 19, C["white"], wrap=True)

    _textbox(slide, Inches(0.5), Inches(6.3), Inches(12.0), Inches(0.7),
             "⏱  Take 3 minutes — then share with the group", 14, C["accent"])

    _notes(slide, slide_data.get("speaker_notes", ""))


def _summary_slide(prs, data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(slide, 0, 0, W, H, C["white"])
    _header(slide, "🎯  Key Takeaways", bg=C["success"])

    for i, tk in enumerate(data.get("key_takeaways", [])[:5]):
        y = Inches(1.6 + i * 0.92)
        num = _oval(slide, Inches(0.4), y, Inches(0.48), Inches(0.48), C["success"])
        tf = num.text_frame
        p = tf.paragraphs[0]
        p.text = str(i + 1)
        p.font.color.rgb = C["white"]
        p.font.size = Pt(13)
        p.font.bold = True
        p.alignment = PP_ALIGN.CENTER
        _textbox(slide, Inches(1.1), y + Inches(0.06), Inches(11.73), Inches(0.72),
                 tk, 18, C["dark"], wrap=True)

    questions = data.get("discussion_questions", [])
    if questions:
        _textbox(slide, Inches(0.5), Inches(6.65), Inches(12.33), Inches(0.55),
                 f"💬  Discuss: {questions[0]}", 14, C["secondary"], italic=True, wrap=True)

    _notes(slide, "Summarize the key points. Ask students to reflect on what surprised them most.")


def _qa_slide(prs, data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(slide, 0, 0, W, H, C["primary"])
    _oval(slide, Inches(8.8), Inches(0.8), Inches(5.5), Inches(5.5), C["midblue"])
    _oval(slide, Inches(-1.0), Inches(4.5), Inches(3.0), Inches(3.5), C["midblue"])

    _textbox(slide, Inches(1.5), Inches(1.8), Inches(9.0), Inches(2.2),
             "Questions?", 62, C["white"], bold=True, align=PP_ALIGN.CENTER)
    _textbox(slide, Inches(1.5), Inches(4.0), Inches(9.0), Inches(0.9),
             "Let's discuss! 🙋", 28, C["accent"], align=PP_ALIGN.CENTER)

    questions = data.get("discussion_questions", [])
    for i, q in enumerate(questions[:3]):
        _textbox(slide, Inches(1.0), Inches(5.1 + i * 0.5), Inches(11.33), Inches(0.45),
                 f"  •  {q}", 14, C["lightblue"], align=PP_ALIGN.CENTER, wrap=True)

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

        topic = (request.form.get("topic") or "").strip()
        if not topic:
            return jsonify({"error": "Please enter a topic."}), 400

        audience = request.form.get("audience", "college students").strip()
        num_slides = max(4, min(int(request.form.get("num_slides", 15)), 50))
        additional_info = (request.form.get("additional_info") or "").strip()

        ppt_data = generate_ppt_content(topic, audience, num_slides, additional_info)
        safe = re.sub(r"[^a-zA-Z0-9 ]", "", ppt_data.get("title", topic))[:40]
        filename = safe.strip().replace(" ", "_") + ".pptx"

        buf = create_presentation(ppt_data)
        del ppt_data  # free memory before streaming

        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name=filename,
        )

    except json.JSONDecodeError as e:
        return jsonify({"error": f"AI response could not be parsed. Try again. ({e})"}), 500
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=True, port=port)
