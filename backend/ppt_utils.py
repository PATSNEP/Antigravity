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
    dest_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
    
    # We verify if we can access it via python-pptx wrapper immediately
    # Typically this works but python-pptx might not 'see' it in the shapes list immediately without reload
    # But for our "replace" logic, we might need to iterate dest_slide.shapes
    
    return new_el
