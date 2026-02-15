from flask import Flask, request, jsonify
import pandas as pd
import io
import base64
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Table, TableStyle

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

        # CLEANING
        df.dropna(how='all', inplace=True)
        df.columns = [col.strip().upper() for col in df.columns]
        df.columns = df.columns.str.replace(r'\.\d+$', '', regex=True)
        df = df.loc[:, ~df.columns.duplicated()]
        duplicate_rows = df.duplicated().sum()
        df = df.drop_duplicates()

        # BASIC METRICS
        num_rows = len(df)
        num_columns = len(df.columns)
        null_count = df.isnull().sum().sum()
        total_cells = df.size
        quality_score = round((1 - null_count/total_cells) * 100, 2)

        numeric_df = df.select_dtypes(include='number')
        means = numeric_df.mean() if not numeric_df.empty else pd.Series()

        # ---------------- AI STYLE SUMMARY ----------------
        summary_text = f"""
        Dataset contains {num_rows} rows and {num_columns} columns.
        {duplicate_rows} duplicate rows were removed.
        Data quality score is {quality_score}%.
        """

        if not means.empty:
            top_metric = means.idxmax()
            summary_text += f"Highest average numeric column is {top_metric} with mean {round(means[top_metric],2)}."

        if "COUNTRY" in df.columns:
            top_country = df["COUNTRY"].value_counts().idxmax()
            summary_text += f" Most frequent country is {top_country}."

        # ---------------- EXCEL GENERATION ----------------
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="DATA", index=False)
        excel_buffer.seek(0)

        excel_base64 = base64.b64encode(excel_buffer.read()).decode('utf-8')

        # ---------------- PDF GENERATION ----------------
        pdf_buffer = io.BytesIO()
        doc = SimpleDocTemplate(pdf_buffer)
        elements = []
        styles = getSampleStyleSheet()

        elements.append(Paragraph("Excel Data Analysis Report", styles['Title']))
        elements.append(Spacer(1, 0.3 * inch))
        elements.append(Paragraph(summary_text, styles['Normal']))
        elements.append(Spacer(1, 0.5 * inch))

        table_data = [["Metric", "Value"],
                      ["Rows", num_rows],
                      ["Columns", num_columns],
                      ["Duplicates Removed", duplicate_rows],
                      ["Quality Score (%)", quality_score]]

        table = Table(table_data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.grey),
            ('GRID', (0,0), (-1,-1), 1, colors.black)
        ]))

        elements.append(table)
        doc.build(elements)

        pdf_buffer.seek(0)
        pdf_base64 = base64.b64encode(pdf_buffer.read()).decode('utf-8')

        return jsonify({
            "excel_file": excel_base64,
            "pdf_file": pdf_base64,
            "summary_text": summary_text
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500