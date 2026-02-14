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

    try:
        df = pd.read_excel(file)

        # Example processing 
        
        # Remove completely empty rows
        df.dropna(how='all', inplace=True)

        # Clean column names
        df.columns = [col.strip().upper() for col in df.columns]

        # Remove duplicate columns (keep first occurrence)
        df = df.loc[:, ~df.columns.duplicated()]

        # Remove duplicate rows
        df = df.drop_duplicates()

        output = io.BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)

        return send_file(
            output,
            download_name="processed.xlsx",
            as_attachment=True,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500
