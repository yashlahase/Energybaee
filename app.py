import os
import uuid
from flask import Flask, render_template, request, jsonify, send_file, session
from extractor import process_bill
from excel_handler import fill_excel_template
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "super-secret-key-123")

# For Vercel deployment, use /tmp for writable storage
UPLOAD_FOLDER = '/tmp/uploads'
OUTPUT_FOLDER = '/tmp/outputs'
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'assets/template.xlsx')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No file selected"}), 400

    # Save file
    filename = f"{uuid.uuid4()}_{file.filename}"
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    try:
        # Extract data
        data = process_bill(filepath)
        if "error" in data:
            return jsonify(data), 500
        
        # Store data in session for the next step
        session['extracted_data'] = data
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/generate', methods=['POST'])
def generate_excel():
    # Get potentially edited data from request
    data = request.json
    if not data:
        data = session.get('extracted_data')

    if not data:
        return jsonify({"error": "No data to process"}), 400

    try:
        output_filename = f"Solar_Load_{uuid.uuid4().hex[:8]}.xlsx"
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)
        
        fill_excel_template(data, TEMPLATE_PATH, output_path)
        
        return jsonify({"download_url": f"/download/{output_filename}"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/download/<filename>')
def download_file(filename):
    path = os.path.join(OUTPUT_FOLDER, filename)
    return send_file(path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=5001)
