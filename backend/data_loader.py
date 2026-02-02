"""
DATEI: backend/data_loader.py
BESCHREIBUNG:
    Dieses Modul ist für das Einlesen und Bereinigen der Rohdaten verantwortlich.
    
    Hauptaufgaben:
    1.  Definition des Mappings zwischen kryptischen CSV-Spaltennamen (z.B. `cr4e2_businessunit...`) 
        und sprechenden internen Variablennamen (z.B. `business_unit`).
    2.  Einlesen der CSV-Datei.
    3.  Konvertierung jeder Zeile in ein `UseCase`-Objekt.
    4.  Gruppierung der Use Cases nach ihrem Geschäftsbereich (Line of Business).
"""

import csv
import collections
from datetime import datetime

# Mapping von CSV-Headern zu internen Schlüsseln
# Dient der Entkopplung: Wenn sich der CSV-Export ändert, muss nur hier angepasst werden.
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
    """
    Repräsentiert einen einzelnen Anwendungsfall (Use Case) aus der Datenquelle.
    Die Attribute werden dynamisch basierend auf dem `COLUMN_MAPPING` gesetzt.
    """
    def __init__(self, data_dict):
        self.raw_data = data_dict
        for internal_key, value in data_dict.items():
            setattr(self, internal_key, value)
            
    def __repr__(self):
        return f"<UseCase {self.title} ({self.line_of_business})>"

def load_data(csv_path):
    """
    Liest die angegebene CSV-Datei und gruppiert die Daten nach 'Line of Business'.
    
    Argumente:
        csv_path (str): Der absolute Pfad zur CSV-Datei.
        
    Rückgabe:
        dict: Ein Dictionary, wobei der Schlüssel der LoB-Name ist (z.B. 'Marketing')
              und der Wert eine Liste von UseCase-Objekten.
              Beispiel: {'Marketing': [UseCase1, UseCase2], ...}
    """
    grouped_data = collections.defaultdict(list)
    foundational_summaries = [] # Platzhalter für spätere AI-Logik (falls benötigt)
    
    try:
        # UTF-8-SIG wird verwendet, um das BOM (Byte Order Mark) von Excel-Exporten korrekt zu handhaben
        with open(csv_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            
            # Prüfung: Sind alle erwarteten Spalten vorhanden?
            headers = reader.fieldnames
            for csv_col in COLUMN_MAPPING.keys():
                if csv_col not in headers:
                    print(f"WARNUNG: Erwartete Spalte '{csv_col}' wurde in der CSV nicht gefunden.")
            
            for row in reader:
                # Zeile in interne Schlüssel mappen
                clean_row = {}
                for csv_col, internal_key in COLUMN_MAPPING.items():
                    val = row.get(csv_col, "").strip()
                    clean_row[internal_key] = val
                    
                uc = UseCase(clean_row)
                
                # Gruppierungs-Logik
                # Wir gruppieren primär nach der Spalte 'cr4e2_lineofbusiness'.
                # Dies ist entscheidend für die spätere Zuordnung zu den Folien.
                
                lob = uc.line_of_business
                if lob:
                    grouped_data[lob].append(uc)
                else:
                    # Fallback für Zeilen ohne LoB
                    grouped_data["Unknown"].append(uc)
                    
        return grouped_data
        
    except FileNotFoundError:
        print(f"FEHLER: Datei nicht gefunden unter {csv_path}")
        return {}
    except Exception as e:
        print(f"KRITISCHER FEHLER beim Laden der Daten: {e}")
        return {}

if __name__ == "__main__":
    # Testlauf (wird nur ausgeführt, wenn das Skript direkt gestartet wird)
    data = load_data("mock_data.csv")
    for lob, cases in data.items():
        print(f"LOB: {lob} - {len(cases)} Fälle")
        for c in cases:
            print(f"  - {c.title} (Typ: {c.use_case_type}, Owner: {c.owner})")
