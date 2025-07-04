from flask import Flask, request, send_file, jsonify
from payroll_processor import process_payroll
import os, tempfile

app = Flask(__name__)

@app.route('/process', methods=['POST'])
def process_endpoint():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    f = request.files['file']
    if f.filename == '':
        return jsonify({"error": "No selected file"}), 400

    # Guarda el Excel de entrada en un temporal
    tmp_in = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    f.save(tmp_in.name)
    tmp_out_path = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx').name

    # Procesa con tu script
    process_payroll(tmp_in.name, tmp_out_path)

    # Devuelve el resultado
    return send_file(tmp_out_path,
                     as_attachment=True,
                     download_name=f"payroll_consolidado_{os.path.basename(f.filename)}")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)