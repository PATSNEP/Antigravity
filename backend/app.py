from flask import Flask, request, render_template, send_file, jsonify
import os
try:
    from backend.ppt_processor import process_ppt
except ImportError:
    from ppt_processor import process_ppt

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400
    
    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)
    
    try:
        # "Scharf geschaltet": Connect real processor
        output_name = process_ppt(filepath, OUTPUT_FOLDER)
        # output_name = "Final_Report.pptx" # Mock return
        
        return jsonify({"message": "Success", "download_url": f"/download/{output_name}"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/download/<filename>', methods=['GET'])
def download(filename):
    path = os.path.join(OUTPUT_FOLDER, filename)
    # Ensure it exists (mocking it for UI test)
    if not os.path.exists(path):
         with open(path, 'w') as f: f.write("dummy pptx")
    return send_file(path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=5000)
