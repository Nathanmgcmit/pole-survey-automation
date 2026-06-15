import sys
sys.stderr.write(f"Python version: {sys.version}\n")
sys.stderr.write(f"Executable: {sys.executable}\n")
sys.stderr.flush()

from flask import Flask, render_template, request, send_file, jsonify
import os
import io
import threading
import subprocess
import time


# Make sure folders exist
os.makedirs('uploads', exist_ok=True)
os.makedirs('output', exist_ok=True)

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
    address      = request.form.get('address') if request.form.get('address_mode') == 'manual' else None
    date_taken   = request.form.get('date_taken') if request.form.get('date_taken_mode') == 'manual' else None
    survey_name  = request.form.get('survey_name') if request.form.get('survey_name_mode') == 'manual' else None

    csv_file = request.files.get('csv_file')
    csv_path = os.path.join(UPLOAD_FOLDER, csv_file.filename)
    csv_file.save(csv_path)

    output_path = os.path.join(OUTPUT_FOLDER, 'output.xlsm')

    # Kill any leftover Excel processes before starting
    subprocess.run(['taskkill', '/F', '/IM', 'excel.exe'], capture_output=True)
    time.sleep(2)

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

    time.sleep(1)
    return send_file(os.path.abspath(output_path), as_attachment=True, download_name='output.xlsm')

@app.route('/cleaner')
def cleaner():
    return render_template('cleaner.html')

# Global address resolution progress
address_progress = {"percent": 0, "message": "Ready", "done": False, "result": None}

@app.route('/cleaner/resolve-addresses', methods=['POST'])
def cleaner_resolve_addresses():
    import threading
    from csv_cleaner import current_data, regroup
    from address_resolver import resolve_addresses

    global address_progress
    address_progress = {"percent": 0, "message": "Starting...", "done": False, "result": None}

    def run_resolution():
        global address_progress
        try:
            poles = current_data.get('clean_poles', current_data.get('poles', []))

            def progress_callback(percent, message):
                global address_progress
                address_progress["percent"] = percent
                address_progress["message"] = message

            updated = resolve_addresses(poles, progress_callback)

            # Update current data
            current_data['clean_poles'] = updated
            current_data['poles']       = updated

            groups = regroup(updated)
            current_data.update(groups)

            address_progress["percent"] = 100
            address_progress["message"] = "Done!"
            address_progress["done"]    = True
            address_progress["result"]  = {
                "status":  "done",
                "updated": len(updated),
                "poles":   updated,
                **groups
            }
        except Exception as e:
            address_progress["percent"] = 0
            address_progress["message"] = f"Error: {str(e)}"
            address_progress["done"]    = True

    thread = threading.Thread(target=run_resolution)
    thread.daemon = True
    thread.start()

    return jsonify({"status": "started"})

@app.route('/cleaner/address-progress')
def address_progress_status():
    global address_progress
    return jsonify(address_progress)

@app.route('/cleaner/restore', methods=['POST'])
def cleaner_restore():
    from csv_cleaner import current_data, regroup
    data = request.get_json()
    if not data:
        return jsonify({"status": "error"})

    poles = data.get('poles', [])
    groups = regroup(poles)

    current_data.update({
        "poles":                  poles,
        "clean_poles":            poles,
        "all_poles":              poles,
        "pending_duplicates":     data.get('duplicates', []),
        "confirmed_duplicates":   [],
        "spelling_issues":        data.get('spelling_issues', []),
        "pole_typos":             data.get('pole_typos', []),
        **groups,
    })

    return jsonify({"status": "ok"})    

@app.route('/cleaner/reset', methods=['POST'])
def cleaner_reset():
    from csv_cleaner import reset_data
    reset_data()
    global address_progress
    address_progress = {"percent": 0, "message": "Ready", "done": False, "result": None}
    return jsonify({"status": "ok"})

@app.route('/cleaner/confirm-duplicates', methods=['POST'])
def cleaner_confirm_duplicates():
    from csv_cleaner import confirm_duplicates
    data        = request.get_json()
    poleid_list = data.get('poleids', [])
    switches    = data.get('switches', [])
    result      = confirm_duplicates(poleid_list, switches)
    return jsonify(result)

@app.route('/cleaner/fix-pole-typos', methods=['POST'])
def cleaner_fix_pole_typos():
    from csv_cleaner import apply_pole_typo_fixes, apply_single_pole_fix
    data   = request.get_json()
    fixes  = data.get('fixes', [])
    poleid = data.get('poleid', None)
    print(f"  poleid received: {poleid!r}")
    print(f"  fixes received: {fixes}")
    if poleid:
        result = apply_single_pole_fix(poleid, fixes[0]['replacement'])
    else:
        result = apply_pole_typo_fixes(fixes)
    return jsonify(result)

@app.route('/cleaner/analyze', methods=['POST'])
def cleaner_analyze():
    from csv_cleaner import analyze_csv
    file = request.files.get('csv_file')
    if not file:
        return jsonify({"error": "No file uploaded"}), 400
    try:
        file_bytes = file.read()
        result = analyze_csv(file_bytes)
        return jsonify(result)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route('/cleaner/apply-fixes', methods=['POST'])
def cleaner_apply_fixes():
    from csv_cleaner import apply_fixes
    data   = request.get_json()
    fixes  = data.get('fixes', [])
    result = apply_fixes(fixes)
    return jsonify(result)


@app.route('/cleaner/download')
def cleaner_download():
    import zipfile
    from csv_cleaner import get_download_csv, get_municipality_zip
    download_type = request.args.get('type', 'all')
    value         = request.args.get('value', '')

    split = request.args.get('split', 'true')

    if download_type == 'municipality' and split == 'true':
        zip_bytes = get_municipality_zip(value)
        return send_file(
            io.BytesIO(zip_bytes),
            mimetype='application/zip',
            as_attachment=True,
            download_name=f'{value}.zip'.replace(' ', '_')
        )

    csv_content = get_download_csv(download_type, value)
    return send_file(
        io.BytesIO(csv_content.encode('utf-8')),
        mimetype='text/csv',
        as_attachment=True,
        download_name=f'{download_type}_{value}.csv'.replace(' ', '_')
    )

@app.route('/cleaner/get-sequence')
def cleaner_get_sequence():
    from csv_cleaner import get_sequence
    return jsonify(get_sequence())


@app.route('/cleaner/apply-sequence', methods=['POST'])
def cleaner_apply_sequence():
    from csv_cleaner import apply_sequence
    data     = request.get_json()
    sequence = data.get('sequence', [])
    result   = apply_sequence(sequence)
    return jsonify(result)

@app.route('/cleaner/download-kml')
def cleaner_download_kml():
    from csv_cleaner import generate_kml
    import io
    kml_content = generate_kml()
    return send_file(
        io.BytesIO(kml_content.encode('utf-8')),
        mimetype='application/vnd.google-earth.kml+xml',
        as_attachment=True,
        download_name='pole_survey.kml'
    )


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