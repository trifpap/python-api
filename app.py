from flask import Flask, request, send_file, jsonify
import pandas as pd
import io

app = Flask(__name__)

@app.route("/")
def home():
    return "Excel API is running!"

@app.route("/process-excel", methods=["POST"])
def process_excel():

    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']

    df = pd.read_excel(file)

    # Example processing
    df.dropna(how='all', inplace=True)
    df.columns = [col.strip().upper() for col in df.columns]
    df = df.drop_duplicates()

    output = io.BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    return send_file(
        output,
        download_name="processed.xlsx",
        as_attachment=True
    )
