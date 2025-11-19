from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import List, Dict, Optional
import google.generativeai as genai
from pptx import Presentation
import io
from datetime import datetime
import json
from pathlib import Path
import os

app = FastAPI(title="AI PPT Generator API - New Flow")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

GEMINI_API_KEY = "AIzaSyCoFbUiVBek9who2wXOC4tkE4mJ0JeV0_o"
genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel("gemini-2.0-flash")

TEMPLATE_DIR = Path("templates")
TEMPLATE_DIR.mkdir(exist_ok=True)

class ContentGenRequest(BaseModel):
    topic: str
    slide_count: int = 3
    language: str = "English"
    tone: str = "Professional"


class OutlineResponse(BaseModel):
    topic: str
    slides: List[Dict]
    message: str


class GeneratePPTRequest(BaseModel):
    topic: str
    outline: List[Dict]  # Generated outline from step 1
    template: str

# -----------------------------------------------------------------------------
# Helper Functions
# -----------------------------------------------------------------------------

TEMPLATE_INFO = {
    "template-1.pptx": {
        "template_id": "template_1",
        "subject": "Blue and White Minimalist Mid-Year Work Report Template",
        "cover": "templates/previews/template-1.png"
    },
    "template-2.pptx": {
        "template_id": "template_2",
        "subject": "Modern Dark Professional Pitch Deck",
        "cover": "templates/previews/template-2.png"
    },
    "template-3.pptx": {
        "template_id": "template_3",
        "subject": "Clean Light Corporate Presentation Template",
        "cover": "templates/previews/template-3.png"
    }
}


TEMPLATE_FOLDER = "templates"
@app.get("/templates")
def list_templates():
    templates = []

    for filename, info in TEMPLATE_INFO.items():

        file_path = os.path.join(TEMPLATE_FOLDER, filename)

        templates.append({
            "file": file_path,              # ← now this is templates/template-1.pptx
            "template_id": info["template_id"],
            "subject": info["subject"],
            "thumbnail_url": info["cover"]
        })

    return {"templates": templates}

def get_paragraph_font_size(paragraph):
    """Return average font size of a paragraph."""
    sizes = []
    for run in paragraph.runs:
        if run.font.size:
            sizes.append(run.font.size.pt)
    return sum(sizes) / len(sizes) if sizes else 0


def replace_text_preserve_style(shape, new_text):
    """
    Replace text while preserving styling safely
    (works even when template uses theme/scheme colors).
    """

    if not shape.has_text_frame:
        return

    text_frame = shape.text_frame
    p = text_frame.paragraphs[0]

    # If no runs, create one
    if len(p.runs) == 0:
        run = p.add_run()
        run.text = new_text
        return

    runs = p.runs

    # Keep the first run for style
    style_run = runs[0]

    # SAFELY extract formatting
    font_style = {
        "bold": style_run.font.bold,
        "italic": style_run.font.italic,
        "underline": style_run.font.underline,
        "size": style_run.font.size,
        "name": style_run.font.name,
        "color_rgb": None
    }

    # COLOR FIX → only copy rgb if available
    if style_run.font.color and hasattr(style_run.font.color, "rgb") and style_run.font.color.rgb:
        font_style["color_rgb"] = style_run.font.color.rgb

    # Remove extra runs
    for r in list(runs)[1:]:
        p._p.remove(r._r)

    # Replace text
    style_run.text = new_text

    # Reapply style safely
    style_run.font.bold = font_style["bold"]
    style_run.font.italic = font_style["italic"]
    style_run.font.underline = font_style["underline"]
    style_run.font.size = font_style["size"]
    style_run.font.name = font_style["name"]

    # Apply color only if real RGB exists
    if font_style["color_rgb"]:
        style_run.font.color.rgb = font_style["color_rgb"]


def get_shape_font_score(shape):
    """Calculate average font size of all paragraphs in a shape."""
    if not shape.has_text_frame:
        return 0

    sizes = []
    for p in shape.text_frame.paragraphs:
        avg_font = get_paragraph_font_size(p)
        if avg_font > 0:
            sizes.append(avg_font)

    return sum(sizes) / len(sizes) if sizes else 0


def generate_presentation_outline(topic: str, slide_count: int) -> List[Dict]:
    """
    Generates N-slide outline.
    Each slide:
    - Title = 2–3 words
    - Content = 1 short sentence
    """

    prompt = f"""
You are a professional presentation writer.

Create a clean {slide_count}-slide outline for the topic: "{topic}".

STRICT RULES:
1. EXACTLY {slide_count} slides.
2. Title = 2–3 words only.
3. Content = 1 short sentence (max 12–15 words).
4. Follow this structure:
   - Slide 1: Introduction
   - Slide {slide_count}: Conclusion
   - Slides 2 to {slide_count-1}: Key points

RETURN ONLY VALID JSON (NO markdown):
[
  {{
    "slide": 1,
    "title": "Two Words",
    "contents": ["Short sentence"]
  }}
]
"""

    try:
        response = model.generate_content(prompt)
        json_text = response.text.strip()

        # Clean possible code blocks
        if "```json" in json_text:
            json_text = json_text.split("```json")[1].split("```")[0].strip()
        elif "```" in json_text:
            json_text = json_text.split("```")[1].split("```")[0].strip()

        outline = json.loads(json_text)

        if not isinstance(outline, list) or len(outline) != slide_count:
            raise ValueError("AI did not generate correct number of slides")

        for i, slide in enumerate(outline, 1):

            slide["slide"] = i

            if "title" not in slide or "contents" not in slide:
                raise ValueError(f"Slide {i} missing title/content")

            if len(slide["title"].split()) > 3:
                raise ValueError(f"Slide {i} title too long (max 3 words)")

            if len(slide["contents"]) != 1:
                raise ValueError(
                    f"Slide {i} must contain exactly 1 content item")

            if len(slide["contents"][0].split()) > 15:
                raise ValueError(f"Slide {i} content too long (max 15 words)")

        return outline

    except json.JSONDecodeError:
        raise ValueError("Failed to parse JSON")
    except Exception as e:
        raise ValueError(f"Failed to generate outline: {str(e)}")


def trim_slides(prs, required_count):
    """
    Remove extra slides beyond required_count.
    """
    total_slides = len(prs.slides)

    if total_slides <= required_count:
        return prs  # no change needed

    # Remove from last to first (safe deletion)
    for index in range(total_slides - 1, required_count - 1, -1):
        rId = prs.slides._sldIdLst[index].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[index]

    return prs


def replace_text_in_ppt(template_path: str, ai_slides: List[Dict]):
    prs = Presentation(template_path)
    logs = []

    required = len(ai_slides)

    # STEP 1: Remove extra slides
    original_slides = len(prs.slides)
    trim_slides(prs, required)

    logs.append(f"🗑 Removed {original_slides - required} extra slides")

    # STEP 2: Replace text in remaining slides
    for idx, slide in enumerate(prs.slides):
        target = ai_slides[idx]
        new_title = target["title"]
        new_contents = target["contents"]

        shapes_list = []

        for shp in slide.shapes:
            if shp.has_text_frame:
                text = shp.text.strip()
                if not text:
                    continue
                font = get_shape_font_score(shp)
                shapes_list.append((shp, text, font))

        if not shapes_list:
            logs.append(f"⚠️ Slide {idx+1}: No text shapes found")
            continue

        # Sort shapes by font size (title = biggest)
        shapes_list.sort(key=lambda x: x[2], reverse=True)

        # Replace title
        title_shape = shapes_list[0][0]
        # title_shape.text = new_title
        replace_text_preserve_style(title_shape, new_title)

        logs.append(f"✅ Slide {idx+1} TITLE replaced")

        # Replace content
        for i, (shape, old_text, _) in enumerate(shapes_list[1:]):
            if i < len(new_contents):
                # shape.text = new_contents[i]
                replace_text_preserve_style(shape, new_contents[i])

                logs.append(f"   CONTENT {i+1} replaced")
            else:
                shape.text = ""
                logs.append(f"   CONTENT {i+1} cleared (no more AI content)")

    stream = io.BytesIO()
    prs.save(stream)
    stream.seek(0)
    return stream, logs

# -----------------------------------------------------------------------------
# API Endpoints
# -----------------------------------------------------------------------------


@app.get("/")
def root():
    return {
        "message": "AI PPT Generator API - New Flow",
        "version": "3.0",
        "status": "running",
        "flow": [
            "1. POST /analyze-topic - Generate outline from topic",
            "2. POST /generate-ppt - Create PPT with outline + template"
        ],
        "templates_directory": str(TEMPLATE_DIR),
        "valid_templates": VALID_TEMPLATES
    }


@app.get("/health")
def health_check():
    """Check system health and templates"""
    templates_status = {}
    for template in VALID_TEMPLATES:
        path = TEMPLATE_DIR / template
        templates_status[template] = "found" if path.exists() else "missing"

    return {
        "status": "healthy",
        "gemini_configured": bool(GEMINI_API_KEY),
        "templates": templates_status
    }



@app.post("/generate-content")
async def generate_content(request: ContentGenRequest):

    if not request.topic.strip():
        raise HTTPException(status_code=400, detail="Topic cannot be empty")
     
    prompt = f"""
You are an expert presentation generator.

TASK:
Generate a hierarchical children JSON structure for the topic: "{request.topic}".

LANGUAGE: {request.language}
TONE: {request.tone}

STRICT OUTPUT RULES:
1. Output MUST be ONLY valid JSON.
2. DO NOT use Markdown.
3. DO NOT use ``` code blocks.
4. DO NOT write any explanation.
5. DO NOT include the word json.
6. DO NOT wrap inside backticks.
7. Output must start with {{
8. Output must end with }}

HIERARCHY FORMAT (FOLLOW EXACT SHAPE):
{{
  "children": [
    {{
      "level": 1,
      "name": "Main Title",
      "children": [
        {{
          "level": 2,
          "name": "Section Heading",
          "children": [
            {{
              "level": 3,
              "name": "Subheading",
              "children": [
                {{
                  "level": 4,
                  "name": "Sub-subheading",
                  "children": [
                    {{
                      "level": 0,
                      "name": "Paragraph text here"
                    }}
                  ]
                }}
              ]
            }}
          ]
        }}
      ]
    }}
  ]
}}

CONTENT RULES:
- Write meaningful titles based on topic.
- Every sub-level must have 2–4 children with level: 0 paragraphs.
- Paragraphs should be 1–2 lines each.
- Maintain logical flow (intro → background → pillars → deep dive → relationships → conclusion).

YOUR RESPONSE:
Return ONLY the final JSON.
"""


    def stream_response():
        try:
            stream = model.generate_content(prompt, stream=True)
            for chunk in stream:
                if chunk.text:
                    clean = chunk.text.strip()
                    final_data = json.dumps({
                        "status": 3,
                          "text": clean
                          })
                    yield f"data: {json.dumps({'status': 4, 'result': final_data})}\n\n"


        except Exception as e:
            yield f"\nERROR: {str(e)}"

    return StreamingResponse(stream_response(), media_type="text/event-stream")


@app.post("/generate-ppt")
async def generate_ppt(request: GeneratePPTRequest):
    """
    STEP 2: Generate final PPT with selected template
    Takes the outline from step 1 and applies it to chosen template
    """
    try:
        print(f"\n{'='*60}")
        print(f"📊 STEP 2: Generating PowerPoint")
        print(f"{'='*60}")
        print(f"Topic: {request.topic}")
        print(f"Template: {request.template}")
        print(f"Slides: {len(request.outline)}")

        # Validate inputs
        if not request.topic.strip():
            raise HTTPException(
                status_code=400, detail="Topic cannot be empty")

        if not request.outline:
            raise HTTPException(
                status_code=400, detail="Outline cannot be empty")

        # Get template path
        print("🔍 Loading template...")
        template_path = get_template_path(request.template)

        # Apply outline to template
        print("📝 Applying content to template...")
        ppt_stream, logs = replace_text_in_ppt(template_path, request.outline)

        print("\n✨ Replacement Log:")
        for log in logs:
            print(f"   {log}")

        print(f"\n✅ PPT created successfully!")
        print(f"   Size: {ppt_stream.getbuffer().nbytes:,} bytes")

        # Generate filename
        safe_topic = "".join(
            c for c in request.topic if c.isalnum() or c in (' ', '-', '_')).strip()
        safe_topic = safe_topic.replace(" ", "_")[:30]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"AI_{safe_topic}_{timestamp}.pptx"

        print(f"   Filename: {filename}")
        print(f"{'='*60}\n")

        # Return file
        headers = {
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Access-Control-Expose-Headers": "Content-Disposition",
        }

        return StreamingResponse(
            ppt_stream,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers=headers
        )

    except FileNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(
            status_code=500, detail=f"Failed to generate PPT: {str(e)}")

# -----------------------------------------------------------------------------
# Run Server
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    import uvicorn

    print("=" * 60)
    print("AI PPT Generator API Server - New Flow")
    print("=" * 60)
    print(f"Templates directory: {TEMPLATE_DIR.absolute()}")
    print(f"Required templates: {', '.join(VALID_TEMPLATES)}")
    print("=" * 60)

    # Check templates
    missing = [t for t in VALID_TEMPLATES if not (TEMPLATE_DIR / t).exists()]
    if missing:
        print(f"⚠️  WARNING: Missing templates: {', '.join(missing)}")
        print(f"   Please add these files to: {TEMPLATE_DIR.absolute()}")
    else:
        print("✅ All templates found!")

    print("=" * 60)
    print("\nNEW FLOW:")
    print("1. POST /analyze-topic - User provides topic, get outline")
    print("2. POST /generate-ppt - User selects template, get final PPT")
    print("=" * 60)
    print("\nStarting server on http://localhost:8000")
    print("API Docs: http://localhost:8000/docs")
    print("=" * 60)

    uvicorn.run(app, host="0.0.0.0", port=8000, reload=True)
