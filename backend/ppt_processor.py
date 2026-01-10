import os
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
try:
    from backend.data_loader import load_data
    from backend.ppt_utils import replace_text_in_shape, duplicate_slide, delete_slide
except ImportError:
    from data_loader import load_data
    from ppt_utils import replace_text_in_shape, duplicate_slide, delete_slide

# Constants matching user request
TEMPLATE_FILE = "../PPTWITHPLACEHOLDERS.pptx" 
# Note: relative path from backend/ assuming we run from backend root? 
# Best to use absolute or relative to this file.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "..", "PPTWITHPLACEHOLDERS.pptx")

FMT_TITLE = {"bold": True, "font_size": 7, "color": RGBColor(0, 176, 240)}
FMT_DATE = {"bold": True, "font_size": 7, "color": RGBColor(0, 0, 0)}

HEATMAP_CONFIGS = [
    {
        "name": "Marketing",
        "filter": "Marketing",
        "slides": [1, 2], # Slide 2 & 3
        "regex_title": r"\{\{Marketing\s+USE\s+CASE\s+Title\s+(\d+)\}\}",
        "fmt_title": "{{{{Marketing USE CASE Title {idx}}}}}",
        "fmt_status": "{{{{StatusupdateUC{idx}Marketing}}}}",
        "key_owner": "{{UseCaseOwnerMarketing}}",
        "fmt_date_d": "{{{{MD{idx}}}}}",
        "fmt_date_a": "{{{{MA{idx}}}}}"
    },
    {
        "name": "Sales",
        "filter": "Sales",
        "slides": [3, 4], # Slide 4 & 5
        "regex_title": r"\{\{SALES\s+USE\s+CASE\s+Title\s+(\d+)\}\}",
        "fmt_title": "{{{{SALES USE CASE Title {idx}}}}}",
        "fmt_status": "{{{{StatusupdateUC{idx}Sales}}}}",
        "key_owner": "{{UseCaseOwnerSales}}",
        "fmt_date_d": "{{{{SD{idx}}}}}",
        "fmt_date_a": "{{{{SA{idx}}}}}"
    },
    {
        "name": "Compliance",
        "filter": "Compliance",
        "slides": [5], # Slide 6
        "regex_title": r"\{\{Compliance\s+USE\s+CASE\s+Title\s+(\d+)\}\}",
        "fmt_title": "{{{{Compliance USE CASE Title {idx}}}}}",
        "fmt_status": "{{{{StatusupdateUC{idx}Compliance}}}}",
        "key_owner": "{{UseCaseOwnerCompliance}}",
        "fmt_date_d": "{{{{COD{idx}}}}}",
        "fmt_date_a": "{{{{COA{idx}}}}}"
    }
]

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
        },
        {
            "name": "Customer Success",
            "filter": "Customer Success",
            "p_title": "Customer Success USE CASE Title {}",
            "p_del": "CUD{}",
            "p_adopt": "CUA{}"
        }
    ]
    
    all_cases = []
    for cases in raw_data.values():
        all_cases.extend(cases)
        
    for config in LOB_CONFIGS:
        # 1. Broad filter by Business Unit (Marketing, Sales, Compliance, Customer Success)
        lob_cases = [c for c in all_cases if config["filter"] in getattr(c, "business_unit", "")]
        
        # 2. Strict filter for Slide 1 Display (User Request: Only "CDP Business Adoption")
        # Ignore "CDP Foundational Use Case"
        slide1_display_cases = [
            c for c in lob_cases 
            if getattr(c, "use_case_type", "").strip() == "CDP Business Adoption"
        ]
        
        print(f"LoB: {config['name']} | Found: {len(lob_cases)} | Display (Business Adoption): {len(slide1_display_cases)}")
        
        # Sort by adoption date or other criteria? Default is CSV order.
        
        for i, case in enumerate(slide1_display_cases):
            # i+1 because placeholders start at 1
            idx = i + 1
            
            # Map attributes to placeholders
            # Title
            key_title = "{{" + config["p_title"].format(idx) + "}}"
            replacements[key_title] = {
                "text": case.title,
                "formatting": FMT_TITLE
            }
            # Delivery Date
            key_del = "{{" + config["p_del"].format(idx) + "}}"
            replacements[key_del] = {"text": case.delivery_date, "formatting": FMT_DATE}
            
            # Adoption Date
            # Adoption Date
            key_adopt = "{{" + config["p_adopt"].format(idx) + "}}"
            replacements[key_adopt] = {"text": case.adoption_date, "formatting": FMT_DATE}
    
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
        
    # 4. Slide 2-6 Logic (Heatmaps / Status Slides)
    # Generic Logic for Marketing, Sales, Compliance

    for config in HEATMAP_CONFIGS:
        # Filter Cases (LoB + CDP Business Adoption)
        cases = [
            c for c in all_cases 
            if config["filter"] in getattr(c, "business_unit", "") 
            and getattr(c, "use_case_type", "").strip() == "CDP Business Adoption"
        ]
        
        print(f"Processing Heatmaps for {config['name']} ({len(cases)} cases found)...")
        
        for slide_idx in config["slides"]:
            if slide_idx >= len(prs.slides): continue
            
            slide = prs.slides[slide_idx]
            
            # Iterate generic
            for shape in slide.shapes:
                if shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            process_heatmap_cell(cell.text_frame, cases, config)
                
                if shape.has_text_frame:
                    process_heatmap_cell(shape.text_frame, cases, config)
    
    # 5. Slide 7 & 8 Logic (Foundational Use Cases)
    # Filter: Type="CDP Foundational Use Case" (Any BU)
    foundational_cases = [
        c for c in all_cases 
        if getattr(c, "use_case_type", "").strip() == "CDP Foundational Use Case"
    ]
    print(f"Processing Foundational Cases ({len(foundational_cases)} cases found)...")
    
    # Slides 7 (Index 6) and 8 (Index 7)
    foundational_slides = [6, 7]
    
    # Calculate Statistics for Overview Message
    total_foundational = len(foundational_cases)
    positive_count = 0
    for c in foundational_cases:
        status_val = getattr(c, "traffic_light", "").strip().lower()
        # Green OR Grey (or Empty/Blank which implies Neutral/Grey) counts as "on track"
        if "green" in status_val or "grey" in status_val or "gray" in status_val or status_val == "":
            positive_count += 1
            
    overview_msg = ""
    if total_foundational > 0:
        percent_positive = (positive_count / total_foundational) * 100
        if percent_positive >= 80:
            overview_msg = "All CDP Foundational Use Cases are on track and will enable business adoption."
        else:
            # Round to int or 1 decimal? User said "X%".
            overview_msg = f"Only {int(percent_positive)}% CDP Foundational Use Cases are on track and will enable business adoption."
    else:
        overview_msg = "No Foundational Use Cases found."

    # Since placeholders are indexed {{... 1}}, {{... 2}}, we can use a global replacement map 
    # tailored to the available cases.
    f_replacements = {}
    
    # Add Overview Messages (1 & 2)
    # The user mentioned {{AIOverviewMessage1}} and {{AIOverviewMessage2}}
    # Formatting: Font size 11, Not Bold (bold: False)
    f_replacements["{{AIOverviewMessage1}}"] = {"text": overview_msg, "formatting": {"bold": False, "font_size": 11, "color": RGBColor(0,0,0)}}
    f_replacements["{{AIOverviewMessage2}}"] = {"text": overview_msg, "formatting": {"bold": False, "font_size": 11, "color": RGBColor(0,0,0)}}
    
    for i, case in enumerate(foundational_cases):
        idx = i + 1 # 1-based index
        
        # Title
        f_replacements[f"{{{{Foundational Use Case Title {idx}}}}}"] = {
            "text": case.title,
            "formatting": FMT_TITLE
        }
        
        # Owner
        f_replacements[f"{{{{Foundational Use Case Owner {idx}}}}}"] = {
            "text": case.owner,
            "formatting": {"font_size": 7, "color": RGBColor(0,0,0)}
        }
        
        # Overall Status
        f_replacements[f"{{{{Overall Status FUC {idx}}}}}"] = {
            "text": getattr(case, "overall_status", "N/A"),
            "formatting": {"font_size": 7, "color": RGBColor(0,0,0)}
        }
        
        
    for slide_idx in foundational_slides:
        if slide_idx >= len(prs.slides): continue
        slide = prs.slides[slide_idx]
        
        # Iterate Shapes AND Process Coloring (Explicit {{prX}})
        for shape in slide.shapes:
            # 1. Standard Replacement (Title, Owner, Overall Status)
            replace_text_in_shape(shape, f_replacements)
            
            # 2. Traffic Light Coloring (via {{prX}} placeholder)
            # We must scan specifically for {{prX}} patterns.
            # This can be in a text_frame (independent shape) or a table cell.
            
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        process_traffic_light_placeholder(cell, foundational_cases)
            
            if shape.has_text_frame:
                process_traffic_light_placeholder(shape, foundational_cases)

    
    FMT_OP_TEXT = {"font_size": 10, "color": RGBColor(0,0,0)}
    
    # Re-gather all cases in LOB order
    ordered_cases = []
    
    for config in HEATMAP_CONFIGS:
        # lob_cases = [c for c in all_cases if config["filter"] in getattr(c, "business_unit", "")]
        lob_cases = []
        for c in all_cases:
            if config["filter"] in getattr(c, "business_unit", ""):
                lob_cases.append(c)
        ordered_cases.extend(lob_cases)

    
    
    # 5. One-Pager Generation (Fill Pre-Duplicated Slides)
    # The user has manually duplicated Slide 9 multiple times in the template.
    # We just need to iterate through cases and fill the corresponding slides.
    # Start Index for One-Pagers: 8 (Slide 9 is index 8)
    
    start_op_index = 8
    print(f"Generating One-Pagers for {len(ordered_cases)} cases (Starting at Slide {start_op_index+1})...")
    
    cases_processed = 0
    
    for i, target_uc in enumerate(ordered_cases):
        slide_idx = start_op_index + i
        
        # Check if we have enough slides in template
        if slide_idx >= len(prs.slides):
            print(f"WARNING: Not enough One-Pager slides in template! Stopped at Case {i+1}.")
            break
            
        slide = prs.slides[slide_idx]
        
        # Prepare Replacements
        # Note: The placeholders are static in the template (e.g. {{UseCaseOnePagerTitel1}}).
        # We replace them in `slide`.
        
        op_replacements = {
            "{{UseCaseOnePagerTitel1}}": {"text": target_uc.title, "formatting": {"bold": True, "color": RGBColor(0, 176, 240)}},
            "{{UseCaseOnePagerPB1}}": {"text": target_uc.problem_statement, "formatting": FMT_OP_TEXT},
            "{{UseCaseOnePagerScope1}}": {"text": target_uc.scope, "formatting": FMT_OP_TEXT},
            "{{UseCaseOnePagerV&KPI1}}": {"text": target_uc.value_kpis, "formatting": FMT_OP_TEXT},
            "{{UseCaseOnePagerBU1}}": {"text": target_uc.line_of_business, "formatting": FMT_OP_TEXT}, 
            "{{UseCaseOnePagerBSU1}}": {"text": target_uc.business_unit, "formatting": FMT_OP_TEXT}, 
            "{{UseCaseOnePagerOwner1}}": {"text": target_uc.owner, "formatting": FMT_OP_TEXT},
            "{{UseCaseOnePagerScopeBC}}": {"text": target_uc.business_contacts, "formatting": FMT_OP_TEXT},
            "{{UseCaseOnePagerScopeAFK}}": {"text": target_uc.affected_key_users, "formatting": FMT_OP_TEXT},
        }
        
        # Fill Slide
        for shape in slide.shapes:
            replace_text_in_shape(shape, op_replacements)
            
        cases_processed += 1

    # 6. Delete Unused Slides
    # If we have 10 OP slides but only 5 cases, we should remove the remaining 5 empty slides.
    # Start deleting from (start_op_index + cases_processed) to end.
    # Note: Deleting from a list while iterating is tricky. Best to delete from end backwards.
    
    last_filled_index = start_op_index + cases_processed - 1
    total_slides = len(prs.slides)
    
    # We want to keep slides 0..last_filled_index.
    # Delete everything after max(last_filled_index, start_op_index-1).
    # (If 0 cases, we keep 0..7, delete 8..end).
    
    # Range to delete: From (last_filled_index + 1) to (total_slides - 1)
    
    delete_start = last_filled_index + 1
    
    # Check if there are slides to delete
    if delete_start < total_slides:
        print(f"Removing unused slides from index {delete_start} to {total_slides-1}...")
        # Delete backwards to avoid index shifting problems
        for idx in range(total_slides - 1, delete_start - 1, -1):
            delete_slide(prs, idx)
    
    print(f"One-Pager Generation Complete. {cases_processed} slides filled.")

    
    # 6. Save Output
    output_filename = "Final_Report.pptx"
    output_path = os.path.join(output_folder, output_filename)
    prs.save(output_path)
    
    return output_filename

def process_heatmap_cell(text_frame, cases, config):
    """
    Scans a text frame for specific LOB Title placeholders (via regex).
    If found, resolves Index, gets the case, and converts placeholders contextually.
    """
    import re
    try:
        from backend.ppt_utils import process_text_frame
    except ImportError:
        from ppt_utils import process_text_frame
    
    text = text_frame.text
    match = re.search(config["regex_title"], text, re.IGNORECASE)
    
    if match:
        idx = int(match.group(1))
        # Case indices are 1-based in PPT, 0-based in list
        case_idx = idx - 1
        
        if 0 <= case_idx < len(cases):
            case = cases[case_idx]
            
            replacements = {}
            
            # 1. Title
            key_title = config["fmt_title"].format(idx=idx)
            replacements[key_title] = {
                "text": case.title,
                "formatting": FMT_TITLE
            }
            
            # 2. Status Update
            key_status = config["fmt_status"].format(idx=idx)
            replacements[key_status] = {
                "text": getattr(case, "status_update", "N/A"),
                "formatting": {"font_size": 7, "color": RGBColor(0,0,0)}
            }
            
            # 3. Owner (Generic Key)
            replacements[config["key_owner"]] = {
                "text": case.owner,
                "formatting": {"font_size": 7, "color": RGBColor(0,0,0)} 
            }
            
            # 4. Dates
            key_date_d = config["fmt_date_d"].format(idx=idx)
            key_date_a = config["fmt_date_a"].format(idx=idx)
            
            replacements[key_date_d] = {"text": case.delivery_date, "formatting": FMT_DATE}
            replacements[key_date_a] = {"text": case.adoption_date, "formatting": FMT_DATE}
            
            # Apply replacements to this text frame
            process_text_frame(text_frame, replacements)

def process_traffic_light_placeholder(shape_or_cell, cases):
    """
    Checks if shape/cell contains {{prX}}. 
    If yes, colors the background based on case status and REMOVES the text.
    """
    if not hasattr(shape_or_cell, "text_frame"): return
    
    text = shape_or_cell.text_frame.text
    import re
    # Match {{pr1}}, {{pr2}}, etc.
    match = re.search(r"\{\{pr(\d+)\}\}", text, re.IGNORECASE)
    
    if match:
        idx = int(match.group(1))
        c_idx = idx - 1
        
        if 0 <= c_idx < len(cases):
            case = cases[c_idx]
            color_val = getattr(case, "traffic_light", "").strip().lower()
            
            final_color = RGBColor(200, 200, 200) # Default Grey
            
            if "green" in color_val:
                final_color = RGBColor(87, 162, 55)
            elif "red" in color_val:
                final_color = RGBColor(255, 0, 0)
            elif "yellow" in color_val:
                final_color = RGBColor(247, 203, 84)
            elif "grey" in color_val or "gray" in color_val:
                final_color = RGBColor(128, 128, 128)
            
            # Apply Color
            shape_or_cell.fill.solid()
            shape_or_cell.fill.fore_color.rgb = final_color
            
            # Remove the placeholder text
            shape_or_cell.text_frame.text = ""
