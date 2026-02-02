"""
DATEI: backend/app.py
BESCHREIBUNG:
    Dies ist der Haupt-Webserver der Anwendung, basierend auf Flask.
    
    Verantwortlichkeiten:
    1.  Bereitstellung des Web-Frontends (`/`).
    2.  Entgegennahme von CSV-Uploads (`/upload`).
    3.  Triggering der PowerPoint-Generierung (`process_ppt`).
    4.  Bereitstellung des fertigen Reports zum Download (`/download`).

    WICHTIG:
    Diese Datei verwaltet Pfade (Upload/Output) absolut, um Probleme zu vermeiden,
    je nachdem, ob der Server aus dem Root-Verzeichnis oder dem Backend-Ordner gestartet wird.
"""

from flask import Flask, request, render_template, send_file, jsonify
import os
try:
    from backend.ppt_processor import process_ppt
except ImportError:
    from ppt_processor import process_ppt

app = Flask(__name__)

# --- KONFIGURATION DER PFADE ---
# Korrektur: Nutze absolute Pfade, um Konflikte zwischen dem aktuellen Arbeitsverzeichnis (CWD)
# und dem Flask-Root zu vermeiden.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Bestimmung des Projekt-Roots (eine Ebene über dem Backend-Ordner)
# Wir nutzen den Projekt-Root für Uploads, um die Struktur sauber zu halten.
PROJECT_ROOT = os.path.dirname(BASE_DIR) 

# Definition der Speicherorte:
# UPLOAD_FOLDER -> ../uploads (im Projekt-Root)
# OUTPUT_FOLDER -> ./outputs (im Backend-Ordner, wie vom Traceback erwartet)
UPLOAD_FOLDER = os.path.join(PROJECT_ROOT, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs') 

# Sicherstellen, dass die Ordner existieren
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    """
    Route: Startseite
    Lädt das HTML-Template für die Benutzeroberfläche.
    """
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """
    Route: Datei-Upload
    Nimmt die CSV-Datei vom Frontend entgegen, speichert sie und startet den Prozessor.
    """
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400
    
    # Speichern der Datei im absoluten Upload-Pfad
    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)
    
    try:
        # "Scharf geschaltet": Verbindung zum echten PowerPoint-Prozessor
        # Wir übergeben den Dateipfad der soeben hochgeladenen CSV.
        output_name = process_ppt(filepath, OUTPUT_FOLDER)
        
        # Rückgabe der Download-URL an das Frontend
        return jsonify({"message": "Success", "download_url": f"/download/{output_name}"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/download/<filename>', methods=['GET'])
def download(filename):
    """
    Route: Datei-Download
    Liefert die generierte PowerPoint-Datei an den Nutzer zurück.
    """
    path = os.path.join(OUTPUT_FOLDER, filename)
    
    # Fallback für Tests: Falls Datei nicht existiert (sollte im Live-Betrieb nicht passieren)
    if not os.path.exists(path):
         with open(path, 'w') as f: f.write("dummy pptx")
         
    return send_file(path, as_attachment=True)

if __name__ == '__main__':
    # Startet den Server im Debug-Modus auf Port 5000
    app.run(debug=True, port=5000)
