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
FMT_DATE = {"bold": True, "font_size": 7, "color": RGBColor(0, 0, 0)}

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
    
    # Filter and Generate Replacements
    replacements = {}
    
    # Configuration for different LoBs
    # filter: substring to match in business_unit
    # placeholders: (Title, DeliveryDate, AdoptionDate) patterns
    LOB_CONFIGS = [
        {
            "name": "Marketing",
            "filter": "Marketing",
            "p_title": "Marketing USE CASE Title {}",
            "p_del": "MD{}",
            "p_adopt": "MA{}"
        },
        {
            "name": "Sales",
            "filter": "Sales",
            "p_title": "SALES USE CASE Title {}", # Note CAPS
            "p_del": "SD{}",
            "p_adopt": "SA{}"
        },
        {
            "name": "Compliance",
            "filter": "Compliance",  # Fixed typo from older CSV
            "p_title": "Compliance USE CASE Title {}",
            "p_del": "COD{}",
            "p_adopt": "COA{}"
        }
    ]
    
    all_cases = []
    for cases in raw_data.values():
        all_cases.extend(cases)
        
    for config in LOB_CONFIGS:
        lob_cases = [
            c for c in all_cases 
            if config["filter"] in getattr(c, "business_unit", "")
        ]
        
        print(f"Found {len(lob_cases)} cases for {config['name']} (Filter: {config['filter']})")
        
        for i, uc in enumerate(lob_cases):
            num = i + 1
            
            # Title
            key_title = "{{" + config["p_title"].format(num) + "}}"
            replacements[key_title] = {"text": uc.title, "formatting": FMT_TITLE}
            
            # Delivery Date
            key_del = "{{" + config["p_del"].format(num) + "}}"
            replacements[key_del] = {"text": uc.delivery_date, "formatting": FMT_DATE}
            
            # Adoption Date
            key_adopt = "{{" + config["p_adopt"].format(num) + "}}"
            replacements[key_adopt] = {"text": uc.adoption_date, "formatting": FMT_DATE}
    
    # 2. Open Template
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"Template not found at {TEMPLATE_PATH}")
        
    prs = Presentation(TEMPLATE_PATH)

    # 3. Slide 1 Logic (Restored)
    # Apply to generic placeholders. 
    # Slide 1 is index 0.
    print("Processing Slide 1 (Overview)...")
    slide1 = prs.slides[0]
    for shape in slide1.shapes:
        replace_text_in_shape(shape, replacements)
        
    # 4. Slide 2 Logic (placeholder, if needed later)
    # ...

    FMT_OP_TEXT = {"font_size": 10, "color": RGBColor(0,0,0)}
    
    # Re-gather all cases in LOB order
    ordered_cases = []
    for config in LOB_CONFIGS:
        lob_cases = [c for c in all_cases if config["filter"] in getattr(c, "business_unit", "")]
        ordered_cases.extend(lob_cases)

    # 5. One-Pager Generation (TEST: Fill Last Slide Only)
    # User Request: "nur die placeholder der letzten slide ausgefüllt werden"
    # We pick the FIRST use case from our list to fill into the existing template slide.
    
    if ordered_cases:
        target_uc = ordered_cases[0] # Pick the first one
        print(f"Test Filling Last Slide with: {target_uc.title}")
        
        last_slide_index = len(prs.slides) - 1
        last_slide = prs.slides[last_slide_index]
        
        # Prepare Replacements
        op_replacements = {
            "{{UseCaseOnePagerTitel1}}": {"text": target_uc.title, "formatting": {"bold": True, "color": RGBColor(0, 176, 240)}},
            "{{UseCaseOnePagerPB1}}": {"text": target_uc.problem_statement, "formatting": FMT_OP_TEXT},
            "{{UseCaseOnePagerScope1}}": {"text": target_uc.scope, "formatting": FMT_OP_TEXT},
            "{{UseCaseOnePagerV&KPI1}}": {"text": target_uc.value_kpis, "formatting": FMT_OP_TEXT},
            "{{UseCaseOnePagerBU1}}": {"text": target_uc.line_of_business, "formatting": FMT_OP_TEXT}, 
            "{{UseCaseOnePagerOwner1}}": {"text": target_uc.owner, "formatting": FMT_OP_TEXT},
            "{{UseCaseOnePagerScopeBC}}": {"text": target_uc.business_contacts, "formatting": FMT_OP_TEXT},
            "{{UseCaseOnePagerScopeAFK}}": {"text": target_uc.affected_key_users, "formatting": FMT_OP_TEXT},
            "{{UseCaseOnePagerBSU1}}": {"text": target_uc.business_unit, "formatting": FMT_OP_TEXT}, 
        }
        
        for shape in last_slide.shapes:
            replace_text_in_shape(shape, op_replacements)
    
    # Loop disabled for now to prevent corruption and verify placeholder logic first.

    
    # 6. Save Output
    output_filename = "Final_Report.pptx"
    output_path = os.path.join(output_folder, output_filename)
    prs.save(output_path)
    
    return output_filename
