from flask import Flask, request, render_template, send_file
import requests
import pandas as pd
import io

app = Flask(__name__)

N8N_WEBHOOK_URL_PDF = "http://localhost:5678/webhook/7c5ac6f8-aeac-40b6-9f73-3d4f0481e91b"
N8N_WEBHOOK_URL_EXCEL = "https://n8n-3-xfpz.onrender.com/home/workflows/webhook-test/54455645"


@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload_pdf')
def upload_pdf_form():  
    return render_template('pdf.html')


@app.route('/upload', methods=['POST'])
def upload_pdf():
    if 'pdf_file' not in request.files:
        return "No file part", 400

    file = request.files['pdf_file']
    files = {'data': (file.filename, file.stream, 'application/pdf')}
    response = requests.post(N8N_WEBHOOK_URL_PDF, files=files)

    # return f"<h3>Sent to n8n:</h3><pre>{response.text}</pre>"
    return "<h3 style='color:green;'>PDF successfully uploaded!</h3>"


@app.route('/upload_excel')
def upload_excel_form():
    return render_template('excel.html')


@app.route('/upload_excel_file', methods=['POST'])
def upload_excel():
    if 'file' not in request.files:
        return "No file part", 400

    file = request.files['file']

    # Read Excel file into DataFrame
    try:
        df = pd.read_excel(file)
    except Exception as e:
        return f"<h3 style='color:red;'>Error reading Excel file: {e}</h3>"

    # Reset file stream to send again to n8n
    file.stream.seek(0)

    files = {
        'file': (file.filename, file.stream, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    }

    try:
        response = requests.post(N8N_WEBHOOK_URL_EXCEL, files=files)
    except Exception as e:
        return f"<h3 style='color:red;'>Request failed: {e}</h3>"

    return render_n8n_response(response, df)


def render_n8n_response(response, df):
    """Combine Excel Question column with n8n outputs and return downloadable Excel."""
    try:
        data = response.json()
    except Exception:
        return f"<h3>Error parsing n8n response:</h3><pre>{response.text}</pre>"

    if not isinstance(data, list) or not all('output' in d for d in data):
        return f"<h3>Unexpected n8n response:</h3><pre>{response.text}</pre>"

    # Detect question column (case-insensitive)
    question_col = next((col for col in df.columns if 'question' in col.lower()), None)

    if question_col:
        questions = df[question_col].fillna('').tolist()
    else:
        questions = ['(No Question Column Found)'] * len(data)

    # Combine into new DataFrame
    outputs = [d['output'] for d in data]
    result_df = pd.DataFrame({
        'Question': questions[:len(outputs)],
        'Output': outputs
    })

    # Save combined data to memory
    output_stream = io.BytesIO()
    with pd.ExcelWriter(output_stream, engine='openpyxl') as writer:
        result_df.to_excel(writer, index=False, sheet_name='Results')

    output_stream.seek(0)

    # Return downloadable Excel file
    return send_file(
        output_stream,
        as_attachment=True,
        download_name='n8n_results.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    ),"done"


if __name__ == '__main__':
    app.run(port=5000, debug=True)