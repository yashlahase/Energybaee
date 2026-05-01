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

    mime_type = file.content_type or "application/octet-stream"
    print(f"Processing file: {file.filename} with mime_type: {mime_type}")

    try:
        # Extract data
        data = process_bill(filepath, mime_type)
        if "error" in data:
            print(f"Extraction returned error: {data['error']}")
            return jsonify(data), 500
        
        # Store data in session for the next step
        session['extracted_data'] = data
        return jsonify(data)
    except Exception as e:
        import traceback
        print("--- UPLOAD ERROR ---")
        print(traceback.format_exc())
        return jsonify({"error": str(e)}), 500

@app.errorhandler(500)
def handle_500(e):
    import traceback
    print("--- SERVER ERROR 500 ---")
    print(traceback.format_exc())
    return jsonify({"error": "Internal Server Error", "details": str(e)}), 500

@app.route('/generate', methods=['POST'])
def generate_excel():
    # Get potentially edited data from request (list of consumer data)
    data_list = request.json
    
    if not data_list:
        return jsonify({"error": "No data to process"}), 400

    if not isinstance(data_list, list):
        data_list = [data_list]

    try:
        output_filename = f"Solar_Load_Calculator_{uuid.uuid4().hex[:8]}.xlsx"
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)
        
        fill_excel_template(data_list, output_path)
        
        return jsonify({"download_url": f"/download/{output_filename}"})
    except Exception as e:
        import traceback
        print("--- GENERATE ERROR ---")
        print(traceback.format_exc())
        return jsonify({"error": str(e)}), 500

@app.route('/download/<filename>')
def download_file(filename):
    path = os.path.join(OUTPUT_FOLDER, filename)
    return send_file(path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=5001)
