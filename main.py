from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import List, Dict
import google.generativeai as genai
from pptx import Presentation
import io
from datetime import datetime
import uuid
import os
from pathlib import Path

# -----------------------------------------------------------------------------
# FastAPI Setup
# -----------------------------------------------------------------------------
app = FastAPI(title="AI PPT Generator API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# -----------------------------------------------------------------------------
# Gemini API Config
# -----------------------------------------------------------------------------
GEMINI_API_KEY = "AIzaSyCoFbUiVBek9who2wXOC4tkE4mJ0JeV0_o"
genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel("gemini-2.0-flash")

# -----------------------------------------------------------------------------
# Template Directory Setup
# -----------------------------------------------------------------------------
TEMPLATE_DIR = Path("templates")
TEMPLATE_DIR.mkdir(exist_ok=True)

# Valid templates
VALID_TEMPLATES = ["template-1.pptx", "template-2.pptx", "template-3.pptx"]

# -----------------------------------------------------------------------------
# Pydantic Models
# -----------------------------------------------------------------------------
class ExtractRequest(BaseModel):
    template: str

class GenerateRequest(BaseModel):
    template: str
    topic: str
    extracted_texts: List[Dict]

class ExtractResponse(BaseModel):
    extracted_texts: List[Dict]
    message: str

# -----------------------------------------------------------------------------
# Helper Functions
# -----------------------------------------------------------------------------
def get_template_path(template_name: str) -> Path:
    """Get full path to template file"""
    if template_name not in VALID_TEMPLATES:
        raise ValueError(f"Invalid template. Must be one of: {VALID_TEMPLATES}")
    
    template_path = TEMPLATE_DIR / template_name
    
    if not template_path.exists():
        raise FileNotFoundError(
            f"Template file not found: {template_path}\n"
            f"Please place your .pptx template files in the '{TEMPLATE_DIR}' directory"
        )
    
    return template_path


def extract_texts_from_ppt(ppt_path: Path) -> List[Dict]:
    """Extract text content from a PPT template"""
    try:
        prs = Presentation(str(ppt_path))
    except Exception as e:
        raise Exception(f"Failed to open PowerPoint file: {str(e)}")
    
    extracted = []

    def recurse_shapes(shape, slide_num):
        try:
            if shape.has_text_frame:
                full_text = "\n".join(
                    "".join(run.text for run in p.runs).strip()
                    for p in shape.text_frame.paragraphs
                ).strip()
                if full_text:
                    extracted.append({
                        "slide": slide_num,
                        "original": full_text,
                        "type": "text",
                        "paragraph_count": len(shape.text_frame.paragraphs),
                    })
            elif shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        text = cell.text.strip()
                        if text:
                            extracted.append({
                                "slide": slide_num,
                                "original": text,
                                "type": "table",
                                "paragraph_count": 1,
                            })
            elif shape.shape_type == 6:  # Group
                for s in shape.shapes:
                    recurse_shapes(s, slide_num)
        except Exception as e:
            print(f"Warning: Error processing shape on slide {slide_num}: {e}")

    for i, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            recurse_shapes(shape, i)

    if not extracted:
        raise Exception("No text content found in template. Template might be empty or have images only.")

    return extracted


def generate_ai_content(topic: str, extracted_texts: List[Dict]) -> List[str]:
    """Generate AI content for extracted template text"""
    template_info = "\n".join(
        f"Text {i+1} (Slide {item['slide']}, Type: {item['type']}): \"{item['original']}\""
        for i, item in enumerate(extracted_texts)
    )

    prompt = f"""You are an expert presentation content creator.

I have a PowerPoint template with {len(extracted_texts)} text placeholders:

{template_info}

Create professional, engaging content for a presentation on: "{topic}"

CRITICAL RULES:
1. Generate EXACTLY {len(extracted_texts)} text replacements
2. For titles (short original text): Keep it 2-7 words, catchy and clear
3. For content (longer original text): Keep it concise, 1-3 lines maximum
4. If original text has multiple lines, you can use \\n for line breaks
5. Match the style and tone of the original placeholder
6. Be specific and relevant to the topic "{topic}"

OUTPUT FORMAT (strict):
Text 1: [your content here]
Text 2: [your content here]
Text 3: [your content here]
...

Generate exactly {len(extracted_texts)} replacements now:"""

    try:
        response = model.generate_content(prompt)
        ai_text = response.text.strip()
    except Exception as e:
        raise Exception(f"AI generation failed: {str(e)}")

    # Parse AI output
    new_texts = []
    for line in ai_text.split("\n"):
        line = line.strip()
        if line.startswith("Text ") and ":" in line:
            # Extract content after "Text N:"
            parts = line.split(":", 1)
            if len(parts) == 2:
                text = parts[1].strip()
                # Handle escaped newlines
                text = text.replace("\\n", "\n")
                # Remove quotes if present
                text = text.strip('"').strip("'")
                if text:
                    new_texts.append(text)

    # Ensure exact count match
    if len(new_texts) < len(extracted_texts):
        print(f"Warning: AI generated {len(new_texts)} texts but needed {len(extracted_texts)}")
        # Fill remaining with generic content
        for i in range(len(new_texts), len(extracted_texts)):
            new_texts.append(f"Content for {topic}")
    
    # Trim excess
    new_texts = new_texts[:len(extracted_texts)]

    return new_texts


def generate_ppt_in_memory(template_path: Path, new_texts: List[str]) -> io.BytesIO:
    """Generate updated PPT in memory"""
    try:
        prs = Presentation(str(template_path))
    except Exception as e:
        raise Exception(f"Failed to open template: {str(e)}")
    
    index = 0

    def replace_text(shape):
        nonlocal index
        try:
            if shape.has_text_frame:
                full_text = "\n".join(
                    "".join(run.text for run in p.runs).strip()
                    for p in shape.text_frame.paragraphs
                ).strip()
                
                if full_text and index < len(new_texts):
                    new_text = new_texts[index]
                    lines = new_text.split("\n")

                    # Clear existing text
                    for p in shape.text_frame.paragraphs:
                        for run in p.runs:
                            run.text = ""

                    # Add new text preserving formatting
                    for i, line in enumerate(lines):
                        if i < len(shape.text_frame.paragraphs):
                            p = shape.text_frame.paragraphs[i]
                            if p.runs:
                                p.runs[0].text = line
                            else:
                                run = p.add_run()
                                run.text = line
                        else:
                            p = shape.text_frame.add_paragraph()
                            p.text = line
                    
                    index += 1

            elif shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        text = cell.text.strip()
                        if text and index < len(new_texts):
                            cell.text = new_texts[index]
                            index += 1

            elif shape.shape_type == 6:  # Group
                for s in shape.shapes:
                    replace_text(s)
        except Exception as e:
            print(f"Warning: Error replacing text in shape: {e}")

    for slide in prs.slides:
        for shape in slide.shapes:
            replace_text(shape)

    # Save to memory
    ppt_stream = io.BytesIO()
    try:
        prs.save(ppt_stream)
        ppt_stream.seek(0)
    except Exception as e:
        raise Exception(f"Failed to save presentation: {str(e)}")

    return ppt_stream

# -----------------------------------------------------------------------------
# API Endpoints
# -----------------------------------------------------------------------------
@app.get("/")
def root():
    return {
        "message": "AI PPT Generator API",
        "version": "2.0",
        "status": "running",
        "templates_directory": str(TEMPLATE_DIR),
        "valid_templates": VALID_TEMPLATES
    }


@app.get("/health")
def health_check():
    """Check if templates exist"""
    templates_status = {}
    for template in VALID_TEMPLATES:
        path = TEMPLATE_DIR / template
        templates_status[template] = "found" if path.exists() else "missing"
    
    return {
        "status": "healthy",
        "templates": templates_status
    }


@app.post("/extract", response_model=ExtractResponse)
async def extract_template(request: ExtractRequest):
    """Extract text elements from the PPT template"""
    try:
        template_path = get_template_path(request.template)
        extracted = extract_texts_from_ppt(template_path)
        
        return ExtractResponse(
            extracted_texts=extracted,
            message=f"Successfully extracted {len(extracted)} text elements from {request.template}"
        )
    
    except FileNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Extraction failed: {str(e)}")


@app.post("/generate")
async def generate_presentation(request: GenerateRequest):
    """Generate a presentation in-memory and return it"""
    try:
        print(f"\n{'='*60}")
        print(f"📝 Generation Request Received")
        print(f"{'='*60}")
        print(f"Template: {request.template}")
        print(f"Topic: {request.topic}")
        print(f"Extracted texts: {len(request.extracted_texts)}")
        
        # Validate inputs
        if not request.topic or request.topic.strip() == "":
            raise HTTPException(status_code=400, detail="Topic cannot be empty")
        
        if not request.extracted_texts:
            raise HTTPException(status_code=400, detail="No extracted texts provided")
        
        # Get template path
        print("🔍 Loading template...")
        template_path = get_template_path(request.template)
        
        # Generate AI content
        print("🤖 Generating AI content...")
        new_texts = generate_ai_content(request.topic, request.extracted_texts)
        print(f"✅ Generated {len(new_texts)} text replacements")
        
        # Generate PPT
        print("📊 Creating PowerPoint presentation...")
        ppt_stream = generate_ppt_in_memory(template_path, new_texts)
        print(f"✅ PPT created, size: {ppt_stream.getbuffer().nbytes} bytes")
        
        # Create filename
        safe_topic = "".join(c for c in request.topic if c.isalnum() or c in (' ', '-', '_')).strip()
        safe_topic = safe_topic.replace(' ', '_')[:30]
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"AI_{safe_topic}_{timestamp}.pptx"
        
        print(f"📥 Sending file: {filename}")
        print(f"{'='*60}\n")
        
        headers = {
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Access-Control-Expose-Headers": "Content-Disposition",
            "Cache-Control": "no-cache"
        }
        
        return StreamingResponse(
            ppt_stream,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers=headers
        )
    
    except FileNotFoundError as e:
        print(f"❌ Error: {str(e)}")
        raise HTTPException(status_code=404, detail=str(e))
    except ValueError as e:
        print(f"❌ Error: {str(e)}")
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Generation failed: {str(e)}")


# -----------------------------------------------------------------------------
# Run Server
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    import uvicorn
    
    print("=" * 60)
    print("AI PPT Generator API Server")
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
    print("Starting server on http://localhost:8000")
    print("API Docs: http://localhost:8000/docs")
    print("=" * 60)
    
    uvicorn.run(app, host="0.0.0.0", port=8000, reload=True)