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
from typing import List, Dict, Optional, Any
import os

app = FastAPI(title="AI PPT Generator API")

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
    slide_count: int = 5
    language: str = "English"
    tone: str = "Professional"

class GeneratePPTRequest(BaseModel):
    content: Any  # JSON string from AI
    template_id: str  # template_id like "template_1"

TEMPLATE_INFO = {
    "template-1.pptx": {
        "id": "template_1",
        "name": "Blue Minimalist Report",
        "category": "Business",
        "slides_count": 5,
        "thumbnail_url": "/templates/previews/template-1.png"
    },
    "template-2.pptx": {
        "id": "template_2",
        "name": "Dark Professional Pitch",
        "category": "Pitch Deck",
        "slides_count": 5,
        "thumbnail_url": "/templates/previews/template-2.png"
    },
    "template-3.pptx": {
        "id": "template_3",
        "name": "Light Corporate",
        "category": "Corporate",
        "slides_count": 5,
        "thumbnail_url": "/templates/previews/template-3.png"
    }
}

# -----------------------------------------------------------------------------
# API Endpoints
# -----------------------------------------------------------------------------

@app.get("/")
def root():
    return {
        "message": "AI PPT Generator API",
        "version": "4.0",
        "status": "running",
        "endpoints": {
            "generate_content": "POST /generate-content",
            "generate_ppt": "POST /generate-ppt",
            "templates": "GET /templates",
            "health": "GET /health"
        }
    }


@app.get("/health")
def health_check():
    """Check system health"""
    templates_status = {}
    for filename in TEMPLATE_INFO.keys():
        path = TEMPLATE_DIR / filename
        templates_status[filename] = "found" if path.exists() else "missing"

    return {
        "status": "healthy",
        "gemini_configured": bool(GEMINI_API_KEY),
        "templates": templates_status
    }


@app.get("/templates")
def list_templates():
    """Get available templates"""
    templates = []
    for filename, info in TEMPLATE_INFO.items():
        templates.append({
            "id": info["id"],
            "name": info["name"],
            "category": info["category"],
            "slides_count": info["slides_count"],
            "thumbnail_url": info["thumbnail_url"]
        })

    return {"templates": templates}


@app.post("/generate-content")
async def generate_content(request: ContentGenRequest):
    """Generate hierarchical content using Gemini AI"""
    if not request.topic.strip():
        raise HTTPException(status_code=400, detail="Topic cannot be empty")

    prompt = f"""
You are an expert presentation generator.

TASK:
Generate a hierarchical JSON structure for the topic: "{request.topic}".

LANGUAGE: {request.language}
TONE: {request.tone}
SLIDES: {request.slide_count}

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
      "name": "Slide 1 Title",
      "children": [
        {{
          "level": 2,
          "name": "Main Point 1",
          "children": [
            {{
              "level": 0,
              "name": "Supporting detail or explanation text"
            }}
          ]
        }},
        {{
          "level": 2,
          "name": "Main Point 2",
          "children": [
            {{
              "level": 0,
              "name": "Supporting detail or explanation text"
            }}
          ]
        }}
      ]
    }},
    {{
      "level": 1,
      "name": "Slide 2 Title",
      "children": [
        {{
          "level": 2,
          "name": "Main Point",
          "children": [
            {{
              "level": 0,
              "name": "Details here"
            }}
          ]
        }}
      ]
    }}
  ]
}}

CONTENT RULES:
- Create exactly {request.slide_count} slides (level 1 nodes).
- Each slide should have 2-4 main points (level 2).
- Each point can have supporting details (level 0).
- Keep text concise and presentation-friendly.
- Follow logical flow: introduction → main content → conclusion.

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
            yield f"data: {json.dumps({'error': str(e)})}\n\n"

    return StreamingResponse(stream_response(), media_type="text/event-stream")













def get_template_path(template_id: str) -> str:
    """Get template file path from template_id"""
    for filename, info in TEMPLATE_INFO.items():
        if info["id"] == template_id:
            path = TEMPLATE_DIR / filename
            if not path.exists():
                raise FileNotFoundError(f"Template file not found: {filename}")
            return str(path)
    raise ValueError(f"Invalid template_id: {template_id}")

def parse_content_to_slides(content_dict: dict) -> List[Dict]:
    """
    Parse hierarchical JSON structure into slide format.
    Input: {"children": [{"level": 1, "name": "...", "children": [...]}, ...]}
    Output: [{"title": "...", "contents": ["...", "..."]}, ...]
    """
    slides = []
    current_slide = None

    def process_node(node):
        nonlocal current_slide
        
        level = node.get("level", 0)
        name = node.get("name", "").strip()
        
        if not name:
            # Process children if name is empty
            for child in node.get("children", []):
                process_node(child)
            return
        
        # Level 1 = New Slide (Title)
        if level == 1:
            if current_slide:
                slides.append(current_slide)
            current_slide = {"title": name, "contents": []}
        
        # Level 2 = Main bullet point
        elif level == 2 and current_slide:
            current_slide["contents"].append(f"• {name}")
        
        # Level 3 = Sub bullet point (indented)
        elif level == 3 and current_slide:
            current_slide["contents"].append(f"  ◦ {name}")
        
        # Level 4 = Sub-sub bullet point (more indented)
        elif level == 4 and current_slide:
            current_slide["contents"].append(f"    ▪ {name}")
        
        # Level 0 = Paragraph/description
        elif level == 0 and current_slide:
            current_slide["contents"].append(name)
        
        # Process children recursively
        for child in node.get("children", []):
            process_node(child)
    
    # Start processing
    if "children" in content_dict:
        for child in content_dict["children"]:
            process_node(child)
    else:
        process_node(content_dict)
    
    # Add last slide
    if current_slide:
        slides.append(current_slide)
    
    return slides

def get_shape_font_size(shape):
    """Get average font size of shape (larger = likely title)"""
    if not shape.has_text_frame:
        return 0
    
    sizes = []
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            if run.font.size:
                sizes.append(run.font.size.pt)
    
    return sum(sizes) / len(sizes) if sizes else 0


def replace_text_preserve_style(shape, new_text):
    """Replace text in shape while preserving formatting"""
    if not shape.has_text_frame:
        return
    
    text_frame = shape.text_frame
    
    # Clear existing paragraphs except first
    while len(text_frame.paragraphs) > 1:
        p = text_frame.paragraphs[1]
        p._element.getparent().remove(p._element)
    
    paragraph = text_frame.paragraphs[0]
    
    # Save original formatting from first run
    original_format = None
    if paragraph.runs:
        first_run = paragraph.runs[0]
        original_format = {
            "bold": first_run.font.bold,
            "italic": first_run.font.italic,
            "size": first_run.font.size,
            "name": first_run.font.name,
        }
        
        # Get color if available
        try:
            if first_run.font.color.rgb:
                original_format["color"] = first_run.font.color.rgb
        except:
            pass
    
    # Clear all runs
    for run in list(paragraph.runs):
        run._element.getparent().remove(run._element)
    
    # Add new text with original formatting
    new_run = paragraph.add_run()
    new_run.text = new_text
    
    if original_format:
        new_run.font.bold = original_format.get("bold")
        new_run.font.italic = original_format.get("italic")
        new_run.font.size = original_format.get("size")
        new_run.font.name = original_format.get("name")
        
        if "color" in original_format:
            try:
                new_run.font.color.rgb = original_format["color"]
            except:
                pass

def trim_slides(prs, required_count):
    """Remove extra slides if template has more than needed"""
    slide_count = len(prs.slides)
    
    if slide_count <= required_count:
        return prs
    
    # Remove slides from end
    for i in range(slide_count - 1, required_count - 1, -1):
        rId = prs.slides._sldIdLst[i].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[i]
    
    return prs

def duplicate_last_slide(prs, times):
    """Duplicate the last slide to match required count"""
    if len(prs.slides) == 0:
        return prs
    
    source_slide = prs.slides[-1]
    
    for _ in range(times):
        # Get blank slide layout (usually index 6, but can vary)
        blank_layout = prs.slide_layouts[6]
        new_slide = prs.slides.add_slide(blank_layout)
        
        # Copy all shapes from source slide
        for shape in source_slide.shapes:
            el = shape.element
            newel = copy.deepcopy(el)
            new_slide.shapes._spTree.insert_element_before(newel, 'p:extLst')
    
    return prs

def apply_content_to_slides(prs, ai_slides: List[Dict]):
    """Apply AI-generated content to presentation slides"""
    logs = []
    required_slides = len(ai_slides)
    current_slides = len(prs.slides)
    
    # Adjust slide count
    if current_slides > required_slides:
        trim_slides(prs, required_slides)
        logs.append(f"🗑 Trimmed from {current_slides} to {required_slides} slides")
    elif current_slides < required_slides:
        # Note: Duplicating slides can be complex, keeping simple for now
        logs.append(f"⚠️ Template has {current_slides} slides, need {required_slides}")
    
    # Apply content to each slide
    for idx, slide in enumerate(prs.slides):
        if idx >= len(ai_slides):
            break
        
        ai_slide = ai_slides[idx]
        title_text = ai_slide["title"]
        content_texts = ai_slide["contents"]
        
        # Find all text shapes and sort by font size
        text_shapes = []
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text.strip():
                font_size = get_shape_font_size(shape)
                text_shapes.append((shape, font_size))
        
        if not text_shapes:
            logs.append(f"⚠️ Slide {idx+1}: No text shapes found")
            continue
        
        # Sort by font size (largest first = title)
        text_shapes.sort(key=lambda x: x[1], reverse=True)
        
        # Replace title (largest text shape)
        title_shape = text_shapes[0][0]
        replace_text_preserve_style(title_shape, title_text)
        logs.append(f"✅ Slide {idx+1}: Title updated - '{title_text}'")
        
        # Replace content in remaining shapes
        content_shapes = [shape for shape, _ in text_shapes[1:]]
        
        for i, content_text in enumerate(content_texts):
            if i < len(content_shapes):
                replace_text_preserve_style(content_shapes[i], content_text)
                logs.append(f"   ✓ Content {i+1}: {content_text[:50]}...")
            else:
                logs.append(f"   ⚠️ Not enough content shapes for item {i+1}")
        
        # Clear unused content shapes
        for i in range(len(content_texts), len(content_shapes)):
            content_shapes[i].text = ""
    
    return prs, logs


@app.post("/create-presentation")
async def create_presentation(request: GeneratePPTRequest):
    """
    Generate PowerPoint from template and AI content
    
    Body:
    {
        "template_id": "template_1",
        "content": {
            "children": [
                {"level": 1, "name": "Slide Title", "children": [...]},
                ...
            ]
        }
    }
    """
    try:
        print(f"\n{'='*60}")
        print(f"📊 Creating PowerPoint Presentation")
        print(f"{'='*60}")
        print(f"Template ID: {request.template_id}")
        
        # Validate template
        template_path = get_template_path(request.template_id)
        print(f"✓ Template found: {template_path}")
        
        # Parse content structure
        print(f"📝 Parsing content structure...")
        ai_slides = parse_content_to_slides(request.content)
        print(f"✓ Parsed {len(ai_slides)} slides")
        
        # Debug: Print parsed slides
        for i, slide in enumerate(ai_slides):
            print(f"\nSlide {i+1}: {slide['title']}")
            print(f"  Contents: {len(slide['contents'])} items")
        
        # Load template
        print(f"\n📂 Loading template...")
        prs = Presentation(template_path)
        print(f"✓ Template loaded: {len(prs.slides)} slides")
        
        # Apply content
        print(f"\n🔄 Applying content to slides...")
        prs, logs = apply_content_to_slides(prs, ai_slides)
        
        # Print logs
        for log in logs:
            print(log)
        
        # Save to BytesIO
        print(f"\n💾 Saving presentation...")
        ppt_stream = io.BytesIO()
        prs.save(ppt_stream)
        ppt_stream.seek(0)
        
        # Generate filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"Presentation_{timestamp}.pptx"
        
        print(f"✅ Presentation created successfully!")
        print(f"{'='*60}\n")
        
        # Return file
        return StreamingResponse(
            ppt_stream,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={
                "Content-Disposition": f'attachment; filename="{filename}"',
                "Access-Control-Expose-Headers": "Content-Disposition"
            }
        )
    
    except FileNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(
            status_code=500,
            detail=f"Failed to create presentation: {str(e)}"
        )













if __name__ == "__main__":
    import uvicorn

    print("=" * 60)
    print("🚀 AI PPT Generator API Server")
    print("=" * 60)
    print(f"📁 Templates directory: {TEMPLATE_DIR.absolute()}")
    print(f"📋 Available templates: {len(TEMPLATE_INFO)}")
    
    # Check templates
    missing = [f for f in TEMPLATE_INFO.keys() if not (TEMPLATE_DIR / f).exists()]
    if missing:
        print(f"⚠️  WARNING: Missing templates: {', '.join(missing)}")
        print(f"   Add these files to: {TEMPLATE_DIR.absolute()}")
    else:
        print("✅ All templates found!")

    print("=" * 60)
    print("\n📡 API Endpoints:")
    print("   POST /generate-content - Generate AI content")
    print("   POST /generate-ppt     - Create PowerPoint")
    print("   GET  /templates        - List templates")
    print("   GET  /health           - Health check")
    print("=" * 60)
    print("\n🌐 Starting server on http://localhost:8000")
    print("📚 API Docs: http://localhost:8000/docs")
    print("=" * 60 + "\n")

    uvicorn.run(app, host="0.0.0.0", port=8000, reload=True)