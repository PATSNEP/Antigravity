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
    # Iterate over paragraphs. 
    # Use index-based loop or copy if we were modifying structure, 
    # but rebuilding IN-PLACE is usually fine if we don't change paragraph count.
    # However, p.clear() clears runs but keeps the paragraph element.
    
    for p in text_frame.paragraphs:
        # Optimization: only touch paragraphs with placeholders
        if "{{" in p.text:
            process_paragraph(p, replacements)

def process_paragraph(p, replacements):
    current_text = p.text
    # Regex split to handle mixed content (like title + subtitle)
    pattern = r"(\{\{.*?\}\})"
    parts = re.split(pattern, current_text)
    
    # Check match again with normalized keys to be sure we have a replacement
    # (Reuse logic from before)
    has_match = False
    for part in parts:
        norm_part = " ".join(part.split())
        if norm_part in replacements:
            has_match = True
            break
            
    if not has_match:
        return

    # Clear and Rebuild Paragraph
    p.clear() 
    # Note: p.clear() removes all runs. Paragraph properties (alignment etc) usually remain.
    
    for part in parts:
        if not part: continue
        
        norm_part = " ".join(part.split())
        
        if norm_part in replacements:
            data = replacements[norm_part]
            # Handle replacement
            run = p.add_run()
            run.text = data["text"]
            apply_formatting(run, data.get("formatting", {}))
        else:
            # Static text part
            # We add it back. 
            if part: # Add even if just whitespace/newlines
                run = p.add_run()
                # Fix for vertical tab (\x0b) rendering as _x000B_
                # Replace with \n for correct line break in PPT
                run.text = part.replace("\x0b", "\n")
                
                # Restore fixed font size for data rows to prevent layout distortion
                # (e.g. alignment tabs becoming too large)
                # Since headers are skipped (no {{), this is safe for data paragraphs.
                run.font.size = Pt(7)


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
    
    # CRITICAL: Generate new ID to prevent corruption
    # Shapes must have unique IDs. 
    # Valid IDs are usually positive integers. 
    # We can try to assume python-pptx will handle it if we modify it, 
    # but safest is to let it generate or pick a random high number?
    # Actually, simpler: removing the cNvPr id attribute might force regeneration on save, 
    # but python-pptx expects it.
    
    # Simple ID generation strategy:
    # Use a random ID or increment global counter? 
    # Easier: Just verify duplication doesn't crash. 
    # The standard way in XML manipulation is typically to create a new ID.
    
    # Let's try to remove the ID and see if PPT repairs it, OR standard practice:
    for sp in new_el.iter():
        if 'id' in sp.attrib:
             # This matches many things. We care about p:cNvPr id="..."
             pass
             
    # Target specifically cNvPr
    check_tags = [
        '{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr',
        '{http://schemas.openxmlformats.org/drawingml/2006/main}cNvPr' 
    ]
    
    # We need to find the NonVisualDrawingProps to set a new ID.
    # But finding the max ID in the slide is hard.
    # Strategy: Just append it. PPT "Repair" often fixes IDs. 
    # User said "without corruption".
    # Better strategy: Do NOT use deepcopy if possible. 
    # But we need exact formatting.
    # Okay, I will try to set a unique ID based on hash of time?
    import time
    import random
    unique_id = int(time.time() * 1000) + random.randint(0, 10000)
    
    # Find cNvPr
    # The tag is usually p:nvSpPr -> p:cNvPr
    # Or p:nvGrpSpPr -> p:cNvPr
    
    found_id = False
    for desc in new_el.iterdescendants():
        if desc.tag.endswith('cNvPr'):
             # Set new ID
             desc.set('id', str(unique_id))
             # Also Name should probably be unique?
             desc.set('name', desc.get('name') + f" {unique_id}")
             found_id = True
             break
             
    dest_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
    
    dest_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
    
    return new_el

def delete_slide(prs, index):
    """
    Delete a slide from the presentation by index.
    """
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    xml_slides.remove(slides[index])
