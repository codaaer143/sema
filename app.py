from flask import Flask, request, send_file, render_template
import os
import subprocess

app = Flask(__name__)

INPUT_DIR = "Input.1"
OUTPUT_DIR = "Output.2"

# Ensure folders exist on the server
os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# This loads your website frontend
@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

# This handles the file processing when they click upload
@app.route('/process', methods=['POST'])
def process_file():
    # 1. Clean out old files from previous runs
    for f in os.listdir(INPUT_DIR):
        os.remove(os.path.join(INPUT_DIR, f))
    for f in os.listdir(OUTPUT_DIR):
        os.remove(os.path.join(OUTPUT_DIR, f))

    # 2. Get the uploaded file from the frontend
    if 'file' not in request.files:
        return "No file uploaded", 400
    
    file = request.files['file']
    if file.filename == '':
        return "No file selected", 400

    # 3. Save it to the exact folder your script expects
    file_path = os.path.join(INPUT_DIR, file.filename)
    file.save(file_path)

    # 4. Run your EXACT untouched script
    subprocess.run(["python", "script.py"])

    # 5. Grab the generated Excel file and send it to the frontend
    output_files = os.listdir(OUTPUT_DIR)
    if not output_files:
        return "Processing skipped (wrong columns or missing data)", 400

    output_file_path = os.path.join(OUTPUT_DIR, output_files[0])
    return send_file(output_file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
