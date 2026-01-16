import csv
import collections
from datetime import datetime

# Mapping from CSV Headers to Internal Keys
COLUMN_MAPPING = {
    "cr4e2_businessunit@OData.Community.Display.V1.FormattedValue": "business_unit",
    "cr4e2_businessadoptiondate": "adoption_date",
    "cr4e2_lateststatusupdate": "status_update",
    "cr4e2_usecasetitle": "title",
    "cr4e2_owner": "owner",
    "cr4e2_businesscontacts": "business_contacts",
    "cr4e2_affectedkeyusers": "affected_key_users",
    "cr4e2_deliverydate": "delivery_date",
    "cr4e2_heatmamapping@OData.Community.Display.V1.FormattedValue": "heatmap_status",
    "cr4e2_lineofbusiness": "line_of_business",
    "cr4e2_owneremail": "owner_email",
    "cr4e2_value": "value_kpis",
    "cr4e2_scope": "scope",
    "cr4e2_problemstatement": "problem_statement",
    "cr4e2_usecasetype@OData.Community.Display.V1.FormattedValue": "use_case_type",
    "cr4e2_overallstatus": "overall_status",
    "cr4e2_pr@OData.Community.Display.V1.FormattedValue": "traffic_light",
    "cr4e2_overallcompleteness": "overall_completeness"
}

class UseCase:
    def __init__(self, data_dict):
        self.raw_data = data_dict
        for internal_key, value in data_dict.items():
            setattr(self, internal_key, value)
            
    def __repr__(self):
        return f"<UseCase {self.title} ({self.line_of_business})>"

def load_data(csv_path):
    """
    Reads the CSV and groups data by Line of Business.
    Returns: Dict {'Marketing': [UseCase, ...], 'Sales': [...]}
    """
    grouped_data = collections.defaultdict(list)
    foundational_summaries = [] # For AI placeholders logic later if needed
    
    with open(csv_path, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        
        # Verify columns exist
        headers = reader.fieldnames
        for csv_col in COLUMN_MAPPING.keys():
            if csv_col not in headers:
                print(f"WARNING: Expected column '{csv_col}' not found in CSV.")
        
        for row in reader:
            # Map row to internal keys
            clean_row = {}
            for csv_col, internal_key in COLUMN_MAPPING.items():
                val = row.get(csv_col, "").strip()
                clean_row[internal_key] = val
                
            uc = UseCase(clean_row)
            
            # Grouping Logic
            # Note: "Foundational" might be a Type, not strictly an LoB in the grouping sense for Slide 1?
            # Requirement says: "Grouped by LoB" and mentions "Foundational usecases grouped by LoB".
            # Inspecting mock data: Compliance has Type "Foundational".
            # We will group by 'line_of_business' primarily.
            
            lob = uc.line_of_business
            if lob:
                grouped_data[lob].append(uc)
            else:
                grouped_data["Unknown"].append(uc)
                
    return grouped_data

if __name__ == "__main__":
    # Test run
    data = load_data("mock_data.csv")
    for lob, cases in data.items():
        print(f"LOB: {lob} - {len(cases)} cases")
        for c in cases:
            print(f"  - {c.title} (Type: {c.use_case_type}, Owner: {c.owner})")
