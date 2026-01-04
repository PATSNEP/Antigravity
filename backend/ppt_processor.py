import os
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
try:
    from backend.data_loader import load_data
    from backend.ppt_utils import replace_text_in_shape
except ImportError:
    from data_loader import load_data
    from ppt_utils import replace_text_in_shape

# Constants matching user request
TEMPLATE_FILE = "../PPTWITHPLACEHOLDERS.pptx" 
# Note: relative path from backend/ assuming we run from backend root? 
# Best to use absolute or relative to this file.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "..", "PPTWITHPLACEHOLDERS.pptx")

FMT_TITLE = {"bold": True, "font_size": 7, "color": RGBColor(0, 176, 240)}
FMT_DATE = {"bold": True, "font_size": 7, "color": RGBColor(0, 176, 240)}

def process_ppt(csv_path, output_folder):
    """
    Processes the PPT using data from the uploaded CSV.
    Returns: generated filename
    """
    print(f"Processing CSV: {csv_path}")
    
    # 1. Load Data
    # Note: data_loader groups by 'line_of_business'. 
    # UseCases csv has "Marketing" in 'cr4e2_businessunit...' formatted value.
    # Our data_loader maps this to 'business_unit'. 
    # Wait, data_loader maps 'cr4e2_lineofbusiness' to 'line_of_business'.
    # User request says: "rows where column ...FormattedValue contains Marketing".
    # I should check if 'line_of_business' or 'business_unit' is the correct filter.
    # From csv head inspection: "Marketing" is in `cr4e2_businessunit...`.
    # `data_loader` maps `cr4e2_businessunit...` -> `business_unit`.
    
    raw_data = load_data(csv_path) 
    # raw_data is dict: {'Marketing': [...], ...} based on LineOfBusiness column?
    # Let's verify data_loader grouping logic.
    # It groups by `uc.line_of_business`.
    # Does `line_of_business` column match "Marketing"?
    # In the raw csv, `cr4e2_lineofbusiness` column exists.
    # But user specifically mentioned `cr4e2_businessunit...FormattedValue` for filtering.
    
    # Let's Re-filter explicitly to be safe, or check if 'Marketing' key in raw_data covers it.
    # If the user says "Business Unit contains Marketing", but data_loader groups by "Line of Business",
    # these might be different. 
    # I will replicate the user's specific filter logic here to be precise.
    
    # Flatten all cases first
    all_cases = []
    for cases in raw_data.values():
        all_cases.extend(cases)
        
    # Filter for Marketing based on Business Unit (Formatted Value)
    marketing_cases = [
        c for c in all_cases 
        if "Marketing" in getattr(c, "business_unit", "") 
    ]
    
    print(f"Found {len(marketing_cases)} Marketing cases.")
    
    # 2. Open Template
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"Template not found at {TEMPLATE_PATH}")
        
    prs = Presentation(TEMPLATE_PATH)
    
    # 3. Slide 1 Logic
    # Replace {{Marketing USE CASE Title X}}, {{MDX}}, {{MAX}}
    # Assumption: X is 1-based index from our filtered list.
    
    replacements = {}
    for i, uc in enumerate(marketing_cases):
        num = i + 1
        
        # Title: {{Marketing USE CASE Title X}}
        key_title = f"{{{{Marketing USE CASE Title {num}}}}}"
        replacements[key_title] = {"text": uc.title, "formatting": FMT_TITLE}
        
        # Delivery Date: {{MDX}}
        key_del = f"{{{{MD{num}}}}}"
        replacements[key_del] = {"text": uc.delivery_date, "formatting": FMT_DATE}
        
        # Adoption Date: {{MAX}}
        key_adopt = f"{{{{MA{num}}}}}"
        replacements[key_adopt] = {"text": uc.adoption_date, "formatting": FMT_DATE}
        
    # Apply to generic placeholders. 
    # Slide 1 is index 0.
    slide1 = prs.slides[0]
    for shape in slide1.shapes:
        replace_text_in_shape(shape, replacements)
        
    # 4. Save Output
    output_filename = "Final_Report.pptx"
    output_path = os.path.join(output_folder, output_filename)
    prs.save(output_path)
    
    return output_filename
