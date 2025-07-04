from flask import Flask, request, send_file, jsonify
import requests
import tempfile
import os
from payroll_processor import process_payroll

app = Flask(__name__)

@app.route('/process', methods=['POST'])
def process_endpoint():
    # 1) Caso: JSON con URL pública
    if request.is_json and 'url' in request.json:
        # Descarga el Excel desde la URL
        tmp_in = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
        resp = requests.get(request.json['url'])
        if resp.status_code != 200:
            return jsonify({"error": "No se pudo descargar el archivo desde URL"}), 400
        tmp_in.write(resp.content)
        tmp_in.flush()
        infile = tmp_in.name

    # 2) Caso: form-data con archivo
    else:
        if 'file' not in request.files:
            return jsonify({"error": "No file part"}), 400
        f = request.files['file']
        if f.filename == '':
            return jsonify({"error": "No selected file"}), 400
        tmp_in = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
        f.save(tmp_in.name)
        infile = tmp_in.name

    # Procesa con tu script
    outfile = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx').name
    process_payroll(infile, outfile)

    # Devuelve el resultado
    return send_file(
        outfile,
        as_attachment=True,
        download_name=f"payroll_consolidado_{os.path.basename(infile)}"
    )

if __name__ == '__main__':
    # En Render se usa PORT=10000; en local 5000. Ajusta si es necesario:
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
