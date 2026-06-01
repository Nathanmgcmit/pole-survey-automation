from flask import Flask, render_template, request, send_file, jsonify
import os
import threading

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
TEMPLATE_FILE = 'Template.xlsm'

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Global progress tracker
progress = {"percent": 0, "message": ""}

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/progress')
def get_progress():
    return jsonify(progress)

@app.route('/generate', methods=['POST'])
def generate():
    global progress
    progress = {"percent": 0, "message": "Starting..."}

    project_name = request.form.get('project_name')
    site_name    = request.form.get('site_name')
    address      = request.form.get('address')
    date_taken   = request.form.get('date_taken')
    survey_name  = request.form.get('survey_name')

    csv_file = request.files.get('csv_file')
    csv_path = os.path.join(UPLOAD_FOLDER, csv_file.filename)
    csv_file.save(csv_path)

    output_path = os.path.join(OUTPUT_FOLDER, 'output.xlsm')

    from excel_generator import generate_excel
    generate_excel(
        template_path = TEMPLATE_FILE,
        csv_path      = csv_path,
        output_path   = output_path,
        project_name  = project_name,
        site_name     = site_name,
        address       = address,
        date_taken    = date_taken,
        survey_name   = survey_name,
        progress      = progress
    )

    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    import socket
    hostname = socket.gethostname()
    local_ip = socket.gethostbyname(hostname)
    print("=" * 50)
    print(f"  App is running!")
    print(f"  Your link: http://{local_ip}:5000")
    print(f"  Share this with your colleagues!")
    print("=" * 50)
    app.run(debug=False, host='0.0.0.0', port=5000)