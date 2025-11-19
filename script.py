from pptx import Presentation
import json

# -------------------------------------------------------------------
# HELPER: Convert hierarchical tree to slide-wise content
# -------------------------------------------------------------------
def parse_content_tree(tree_data):
    """
    Converts hierarchical content tree into slide-based structure.
    Each level 1 item = 1 slide
    Level 2 items = bullet points
    Level 0 items = sub-bullets or content
    """
    slides_content = []
    
    for section in tree_data:
        if section.get("level") != 1:
            continue
            
        slide_data = {
            "title": section.get("name", ""),
            "bullets": []
        }
        
        # Process children (level 2 and level 0 items)
        for child in section.get("children", []):
            if child.get("level") == 2:
                # Level 2 = main bullet point
                bullet = {
                    "text": child.get("name", ""),
                    "sub_bullets": []
                }
                
                # Check for sub-bullets (level 0)
                for sub_child in child.get("children", []):
                    if sub_child.get("level") == 0:
                        bullet["sub_bullets"].append(sub_child.get("name", ""))
                
                slide_data["bullets"].append(bullet)
            elif child.get("level") == 0:
                # Direct level 0 under level 1
                slide_data["bullets"].append({
                    "text": child.get("name", ""),
                    "sub_bullets": []
                })
        
        slides_content.append(slide_data)
    
    return slides_content


# -------------------------------------------------------------------
# 1) EXTRACT TEMPLATE STRUCTURE
# -------------------------------------------------------------------
def extract_template(path):
    """Extract placeholder structure from template"""
    prs = Presentation(path)
    slides = []

    for idx, slide in enumerate(prs.slides, start=1):
        placeholders = []

        for shape in slide.shapes:
            if shape.has_text_frame:
                # Detect placeholder type based on position/size
                placeholder_type = "BODY"
                if shape.top < 1000000:  # Top of slide (approx)
                    if len(shape.text.strip()) < 50:  # Short text = likely title
                        placeholder_type = "TITLE"
                
                placeholders.append({
                    "id": shape.shape_id,
                    "text": shape.text.strip(),
                    "type": placeholder_type,
                    "left": shape.left,
                    "top": shape.top
                })

        # Sort by position (top to bottom, left to right)
        placeholders.sort(key=lambda x: (x["top"], x["left"]))
        
        slides.append({
            "slide_number": idx,
            "placeholders": placeholders
        })

    return {"slides": slides}


# -------------------------------------------------------------------
# 2) FILL PRESENTATION WITH CONTENT
# -------------------------------------------------------------------
def fill_presentation(template_path, content_tree, output_path):
    """
    Fill template with hierarchical content
    """
    prs = Presentation(template_path)
    slides_content = parse_content_tree(content_tree)
    
    # Match content slides with template slides
    for idx, (slide, content) in enumerate(zip(prs.slides, slides_content)):
        if idx >= len(slides_content):
            break
            
        # Find title and body placeholders
        title_shape = None
        body_shape = None
        
        for shape in slide.shapes:
            if shape.has_text_frame:
                # First large text box at top = title
                if shape.top < 1500000 and title_shape is None:
                    title_shape = shape
                # Second text box = body
                elif body_shape is None and shape != title_shape:
                    body_shape = shape
        
        # Fill title
        if title_shape and content.get("title"):
            title_shape.text = content["title"]
        
        # Fill body with bullets
        if body_shape and content.get("bullets"):
            text_frame = body_shape.text_frame
            text_frame.clear()  # Clear existing content
            
            for bullet_idx, bullet in enumerate(content["bullets"]):
                # Add main bullet
                if bullet_idx == 0:
                    p = text_frame.paragraphs[0]
                else:
                    p = text_frame.add_paragraph()
                
                p.text = bullet["text"]
                p.level = 0  # Main bullet level
                
                # Add sub-bullets
                for sub_bullet in bullet.get("sub_bullets", []):
                    p = text_frame.add_paragraph()
                    p.text = sub_bullet
                    p.level = 1  # Sub-bullet level
    
    prs.save(output_path)
    return output_path


# -------------------------------------------------------------------
# USAGE
# -------------------------------------------------------------------
if __name__ == "__main__":
    template = "templates/template-1.pptx"
    
    # Load your AI-generated content
    with open("ai_content.json") as f:
        ai_content = json.load(f)
    
    # Optional: Extract template structure for debugging
    template_structure = extract_template(template)
    with open("template_structure.json", "w") as f:
        json.dump(template_structure, f, indent=4)
    
    # Fill presentation with content
    output = fill_presentation(template, ai_content, "output_filled.pptx")
    
    print(f"✅ SUCCESS: {output}")
    print(f"📊 Processed {len(ai_content)} sections")