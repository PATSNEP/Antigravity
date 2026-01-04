from pptx import Presentation
import os

filename = "PPTWITHPLACEHOLDERS.pptx"
cwd = "/Users/patrickschnepf/Desktop/Master WINF/1 Semester/Projekt DT/Antigravity"
path = os.path.join(cwd, filename)

try:
    pr = Presentation(path)
    print(f"Inspecting {filename}...")

    def print_shape_text(shape, indent=0):
        indent_str = " " * indent
        if hasattr(shape, "text") and shape.text.strip():
            txt = shape.text
            if "Marketing" in txt or "{{" in txt:
                 print(f"{indent_str}MATCH: {repr(txt)}")
        
        if shape.shape_type == 6:  # Group
            for child in shape.shapes:
                print_shape_text(child, indent + 2)
                
        if shape.shape_type == 19: # Table
            for row in shape.table.rows:
                for cell in row.cells:
                    if cell.text_frame.text.strip():
                        txt = cell.text_frame.text
                        if "Marketing" in txt or "{{" in txt:
                            print(f"{indent_str}TABLE MATCH: {repr(txt)}")

    for i, slide in enumerate(pr.slides):
        print(f"Slide {i+1}:")
        for shape in slide.shapes:
            print_shape_text(shape, 2)

except Exception as e:
    print(f"Error: {e}")
