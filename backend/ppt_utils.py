from pptx.util import Pt
from pptx.dml.color import RGBColor
import copy
import re

def replace_text_in_shape(shape, replacements):
    """
    Replaces text in a shape based on a dictionary of replacements.
    replacements: Dict { placeholder_key: { 'text': str, 'formatting': dict, 'is_bullet': bool } }
    """
    if not shape.has_text_frame and not shape.has_table:
        return

    # Table handling
    if shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                text_frame = cell.text_frame
                process_text_frame(text_frame, replacements)
    
    # TextFrame handling
    if shape.has_text_frame:
        process_text_frame(shape.text_frame, replacements)

def process_text_frame(text_frame, replacements):
    current_text = text_frame.text.strip()
    if not current_text:
        return

    # 1. Exact/Complex Match Check (e.g. Slide 2 specific grouped placeholders)
    # We iterate through keys to see if ALL keys of a "group" are present? 
    # Or simpler: Check if any key is in the text.
    
    # Strategy: Find all unique keys present in the text
    found_keys = []
    for key in replacements:
        if key in current_text:
            found_keys.append(key)
    
    if not found_keys:
        return

    print(f"Found keys {found_keys} in text: '{current_text[:30]}...'")
    
    # If multiple keys found (like Slide 2), we might need careful reconstruction.
    # Logic: If the replacement dict indicates "is_bullet", we assume paragraph reconstruction anyway.
    
    # Detailed replacement logic:
    # We will reconstruct the text_frame.
    # Warning: This destroys existing formatting of non-replaced text if mixed!
    # Assumption (from reqs): "ausschließlich definierte Inhalte ersetzen" usually means the placeholder IS the content or we control the cell.
    
    # For robust partial replacement, we would need run-level splicing. 
    # Given the project scope and "placeholder-principles", complete reconstruction is often safer/cleaner IF the placeholder is the main content.
    # Exception: Slide 2 had mixed placeholders. We will rebuild linearly.
    
    # Sort keys by position in text to process in order? 
    # Actually, simpler: The user wants specific data filled.
    # If we have specific keys for this shape, we clear and fill.
    
    # Special Case: generic "Use Case Title" replacements vs specific keys.
    # If the text is EXACTLY one key, simple replace.
    
    if len(found_keys) == 1 and found_keys[0] == current_text:
        key = found_keys[0]
        data = replacements[key]
        text_frame.clear()
        p = text_frame.paragraphs[0]
        apply_replacement_to_paragraph(p, data)
        return

    # Mixed content or multiple keys (Slide 2)
    # Rebuilding tactic:
    # If we find specific known combinations (like the Slide 2 set), we use a hardcoded structure or a "smart rebuild"?
    # Smart rebuild: Regex split?
    # `{{Key}}` -> Replace. `\n` -> New Paragraph.
    
    # Let's try a regex split approach to preserve structure:
    # Pattern: ({{.*?}}) capture group
    
    pattern = r"(\{\{.*?\}\})"
    parts = re.split(pattern, current_text)
    
    text_frame.clear()
    p = text_frame.paragraphs[0] # Start with first paragraph
    
    for part in parts:
        if not part: continue
        
        # Check if part is a key to replace
        if part in replacements:
            data = replacements[part]
            
            # If "is_bullet", we might need new paragraphs
            if data.get("is_bullet"):
                # Bullets usually start on a new paragraph.
                # If we are currently at start of p (empty), use p.
                # Else add new p.
                
                content_parts = data["text"].split(";")
                for i, content in enumerate(content_parts):
                    if not content.strip(): continue
                    
                    if i == 0 and len(p.runs) == 0:
                        target_p = p
                    else:
                        target_p = text_frame.add_paragraph()
                    
                    target_p.level = 0 # Assume level 0
                    run = target_p.add_run()
                    run.text = "• " + content.strip()
                    apply_formatting(run, data.get("formatting", {}))
                    
                    # Update current p pointer to the last one added
                    p = target_p
            else:
                # Normal text replacement
                run = p.add_run()
                run.text = data["text"]
                apply_formatting(run, data.get("formatting", {}))
        
        else:
            # Static text part (newlines, labels, etc.)
            # Handle newlines if they exist in valid text
            # If part is just whitespace/newlines, we might just add runs.
            # But \n in run.text doesn't always make a new paragraph.
            # If original had \n, split and add paragraphs?
            
            # Simple approach: Add as run text.
            if part.strip() or part == "\n":
                run = p.add_run()
                run.text = part
                # Formatting? Inherit default (no easy way to grab original run font without complex mapping)
                # Assumption: Template has defaults.
                run.font.size = Pt(7) # Safe default from reqs

def apply_replacement_to_paragraph(paragraph, data):
    run = paragraph.add_run()
    run.text = data["text"]
    apply_formatting(run, data.get("formatting", {}))

def apply_formatting(run, formatting):
    if not formatting: return
    font = run.font
    if "bold" in formatting:
        font.bold = formatting["bold"]
    if "font_size" in formatting:
        font.size = Pt(formatting["font_size"])
    if "color" in formatting:
        font.color.rgb = formatting["color"]

def duplicate_slide(prs, source_slide_index):
    """
    Duplicate the slide at source_slide_index and append it to the end of the presentation.
    Returns the new slide.
    """
    source_slide = prs.slides[source_slide_index]
    slide_layout = source_slide.slide_layout
    dest_slide = prs.slides.add_slide(slide_layout)
    
    # Copy shapes
    for shape in source_slide.shapes:
        new_shape = copy_shape(shape, dest_slide)
        
    return dest_slide

def copy_shape(shape, dest_slide):
    # Simple shape copy implementation
    # Note: python-pptx doesn't have a native 'clone_shape'
    # Use element copying for fidelity
    
    new_el = copy.deepcopy(shape.element)
    dest_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
    
    # We verify if we can access it via python-pptx wrapper immediately
    # Typically this works but python-pptx might not 'see' it in the shapes list immediately without reload
    # But for our "replace" logic, we might need to iterate dest_slide.shapes
    
    return new_el
