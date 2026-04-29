from flask import Flask, request, jsonify, send_file, render_template
from pypdf import PdfReader, PdfWriter
import os, io, zipfile, tempfile, uuid

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500 MB

# Temp store: session_id -> {filename: bytes}
SESSIONS = {}

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/unlock', methods=['POST'])
def unlock():
    password = request.form.get('password', '')
    session_id = request.form.get('session_id', str(uuid.uuid4()))
    files = request.files.getlist('pdfs')

    has_session_data = session_id in SESSIONS and 'locked_data' in SESSIONS[session_id] and SESSIONS[session_id]['locked_data']
    if not files and not has_session_data:
        return jsonify({'error': 'No files uploaded'}), 400

    unlocked = []
    locked = []

    # If there's already a session (retry for locked files), use stored locked files
    pending_files = {}
    if session_id in SESSIONS and 'locked_data' in SESSIONS[session_id]:
        pending_files = SESSIONS[session_id]['locked_data']
        # Also re-check files sent again (from retry upload)
        for f in files:
            if f.filename:
                pending_files[f.filename] = f.read()
    else:
        for f in files:
            if f.filename:
                pending_files[f.filename] = f.read()

    unlocked_data = SESSIONS.get(session_id, {}).get('unlocked_data', {})
    newly_locked = {}

    for filename, file_bytes in pending_files.items():
        try:
            reader = PdfReader(io.BytesIO(file_bytes))
            if reader.is_encrypted:
                result = reader.decrypt(password)
                if result == 0:
                    # Wrong password
                    newly_locked[filename] = file_bytes
                    locked.append(filename)
                    continue
            # Write unlocked PDF
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)
            buf = io.BytesIO()
            writer.write(buf)
            unlocked_data[filename] = buf.getvalue()
            unlocked.append(filename)
        except Exception as e:
            newly_locked[filename] = file_bytes
            locked.append(filename)

    SESSIONS[session_id] = {
        'unlocked_data': unlocked_data,
        'locked_data': newly_locked
    }

    return jsonify({
        'session_id': session_id,
        'unlocked': unlocked,
        'locked': locked,
        'total_unlocked': len(unlocked_data)
    })

@app.route('/download/<session_id>', methods=['GET'])
def download(session_id):
    if session_id not in SESSIONS:
        return jsonify({'error': 'Session not found'}), 404

    unlocked_data = SESSIONS[session_id].get('unlocked_data', {})
    if not unlocked_data:
        return jsonify({'error': 'No unlocked files'}), 404

    if len(unlocked_data) == 1:
        filename, data = next(iter(unlocked_data.items()))
        return send_file(
            io.BytesIO(data),
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f"unlocked_{filename}"
        )

    # Multiple files → ZIP
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for filename, data in unlocked_data.items():
            zf.writestr(f"unlocked_{filename}", data)
    zip_buf.seek(0)
    return send_file(
        zip_buf,
        mimetype='application/zip',
        as_attachment=True,
        download_name='unlocked_pdfs.zip'
    )

@app.route('/clear/<session_id>', methods=['POST'])
def clear(session_id):
    SESSIONS.pop(session_id, None)
    return jsonify({'ok': True})

if __name__ == '__main__':
    app.run(debug=True, port=5000)
