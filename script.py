import google.generativeai as genai
from pptx import Presentation
import os

# -----------------------------------------------------------------------------
# 🔑 Gemini API Config
# -----------------------------------------------------------------------------
genai.configure(api_key="AIzaSyCoFbUiVBek9who2wXOC4tkE4mJ0JeV0_o")
model = genai.GenerativeModel("gemini-2.0-flash")

# -----------------------------------------------------------------------------
# 🎯 Step 1: Extract Text from Template PPT (TREAT MULTI-LINE AS ONE)
# -----------------------------------------------------------------------------
input_path = "templates/template-1.pptx"
output_path = "templates/auto_ai_updated.pptx"

prs = Presentation(input_path)

extracted_structure = []

def extract_texts(shape, slide_num):
    """Extract text treating entire text frame as one unit"""
    if shape.has_text_frame:
        # Get ALL text from the entire text frame (all paragraphs combined)
        full_text = "\n".join(
            "".join(run.text for run in p.runs).strip() 
            for p in shape.text_frame.paragraphs
        ).strip()
        
        if full_text:
            extracted_structure.append({
                'slide': slide_num,
                'original': full_text,
                'type': 'text',
                'paragraph_count': len(shape.text_frame.paragraphs)
            })
    
    elif shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                text = cell.text.strip()
                if text:
                    extracted_structure.append({
                        'slide': slide_num,
                        'original': text,
                        'type': 'table',
                        'paragraph_count': 1
                    })
    
    elif shape.shape_type == 6:  # Group shape
        for s in shape.shapes:
            extract_texts(s, slide_num)

# Extract from all slides
for i, slide in enumerate(prs.slides, 1):
    for shape in slide.shapes:
        extract_texts(shape, i)

print("📄 Extracted Template Structure:\n")
for idx, item in enumerate(extracted_structure):
    print(f"{idx+1}. [Slide {item['slide']}] {repr(item['original'])}")

print("\n-------------------------------------------\n")

# -----------------------------------------------------------------------------
# 🧠 Step 2: Generate AI Content Based on Template Structure
# -----------------------------------------------------------------------------
topic = "The Future of Cricket"

# Create structured prompt with template info
template_info = "\n".join([
    f"Text Box {i+1} (Slide {item['slide']}): {repr(item['original'])}"
    for i, item in enumerate(extracted_structure)
])

prompt = f"""
You are an expert presentation designer.

I have a PowerPoint template with {len(extracted_structure)} text boxes:
{template_info}

Create professional content on the topic: "{topic}"

Rules:
- Generate EXACTLY {len(extracted_structure)} replacements
- If original has multiple lines (like "IT Software\\nPitch Deck"), keep that format
- Keep titles SHORT (2-5 words)
- Keep content CONCISE (1-2 lines)

Output format (preserve line breaks with \\n):
Text 1: ...
Text 2: ...
Text 3: ...
"""

response = model.generate_content(prompt)
ai_text = response.text.strip()

print("🧠 Gemini Generated Content:\n")
print(ai_text)
print("\n-------------------------------------------\n")

# -----------------------------------------------------------------------------
# 🧩 Step 3: Parse AI Response
# -----------------------------------------------------------------------------
new_texts = []
for line in ai_text.split("\n"):
    if line.strip().startswith("Text "):
        text = line.split(":", 1)[1].strip() if ":" in line else ""
        if text:
            # Replace literal \n with actual newline
            text = text.replace("\\n", "\n")
            new_texts.append(text)

# Safety check
if len(new_texts) != len(extracted_structure):
    print(f"⚠️ Warning: AI generated {len(new_texts)} texts but template needs {len(extracted_structure)}")
    while len(new_texts) < len(extracted_structure):
        new_texts.append("[Content]")
    new_texts = new_texts[:len(extracted_structure)]

print("✅ Parsed new_texts:\n")
for i, text in enumerate(new_texts):
    print(f"{i+1}. {repr(text)}")
print("\n-------------------------------------------\n")

# -----------------------------------------------------------------------------
# 🔄 Step 4: Replace Text While Preserving Style (ENTIRE TEXT FRAME)
# -----------------------------------------------------------------------------
prs = Presentation(input_path)  # Reload fresh template
index = 0

def replace_with_style(shape):
    global index
    
    if shape.has_text_frame:
        # Get full existing text
        full_text = "\n".join(
            "".join(run.text for run in p.runs).strip() 
            for p in shape.text_frame.paragraphs
        ).strip()
        
        if full_text and index < len(new_texts):
            new_text = new_texts[index]
            print(f"🔁 [{index+1}] '{full_text}' → '{new_text}'")
            
            # Split new text by lines
            new_lines = new_text.split("\n")
            
            # Clear all existing paragraphs
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.text = ""
            
            # Add new content line by line
            for i, line in enumerate(new_lines):
                if i < len(shape.text_frame.paragraphs):
                    # Use existing paragraph
                    p = shape.text_frame.paragraphs[i]
                    if p.runs:
                        p.runs[0].text = line
                    else:
                        p.add_run().text = line
                else:
                    # Add new paragraph if needed
                    p = shape.text_frame.add_paragraph()
                    p.text = line
            
            index += 1
    
    elif shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                text = cell.text.strip()
                if text and index < len(new_texts):
                    print(f"🔁 [{index+1}] Table: '{text}' → '{new_texts[index]}'")
                    cell.text = new_texts[index]
                    index += 1
    
    elif shape.shape_type == 6:  # Group
        for s in shape.shapes:
            replace_with_style(s)

# Process all slides
for slide in prs.slides:
    for shape in slide.shapes:
        replace_with_style(shape)

prs.save(output_path)
print(f"\n🎉 Done! AI-generated PPT saved as: {output_path}")