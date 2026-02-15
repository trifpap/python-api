from flask import Flask, request, jsonify
import pandas as pd
import io
import base64
import datetime
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet

app = Flask(__name__)

@app.route("/")
def home():
    return "Excel API is running!"

@app.route("/process-excel", methods=["POST"])
def process_excel():

    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    uploaded_file = request.files['file']
    original_filename = uploaded_file.filename

    try:
        df = pd.read_excel(uploaded_file)

        # ---------------- CLEANING ----------------
        df.dropna(how='all', inplace=True)
        df.columns = [col.strip().upper() for col in df.columns]
        df.columns = df.columns.str.replace(r'\.\d+$', '', regex=True)
        df = df.loc[:, ~df.columns.duplicated()]
        duplicate_rows = df.duplicated().sum()
        df = df.drop_duplicates()

        # ---------------- METRICS ----------------
        num_rows = len(df)
        num_columns = len(df.columns)
        null_counts = df.isnull().sum()
        total_cells = df.size
        total_nulls = null_counts.sum()
        quality_score = round((1 - total_nulls/total_cells) * 100, 2)

        numeric_df = df.select_dtypes(include='number')
        stats_df = pd.DataFrame()

        if not numeric_df.empty:
            stats_df = pd.DataFrame({
                "Mean": numeric_df.mean(),
                "Median": numeric_df.median(),
                "Std Dev": numeric_df.std(),
                "Min": numeric_df.min(),
                "Max": numeric_df.max()
            })

        country_freq = pd.DataFrame()
        if "COUNTRY" in df.columns:
            country_freq = df["COUNTRY"].value_counts().reset_index()
            country_freq.columns = ["Country", "Count"]

        summary_df = pd.DataFrame({
            "Metric": [
                "Rows",
                "Columns",
                "Duplicate Rows Removed",
                "Total Null Values",
                "Data Quality Score (%)"
            ],
            "Value": [
                num_rows,
                num_columns,
                duplicate_rows,
                total_nulls,
                quality_score
            ]
        })

        null_df = null_counts.reset_index()
        null_df.columns = ["Column", "Null Count"]

        # ---------------- EXCEL GENERATION ----------------
        excel_filename = f"processed_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="DATA", index=False)
            summary_df.to_excel(writer, sheet_name="SUMMARY", index=False)
            stats_df.to_excel(writer, sheet_name="NUMERIC_STATS")
            null_df.to_excel(writer, sheet_name="NULL_COUNTS", index=False)

            if not country_freq.empty:
                country_freq.to_excel(writer, sheet_name="COUNTRY_FREQ", index=False)

        excel_buffer.seek(0)
        excel_base64 = base64.b64encode(excel_buffer.read()).decode('utf-8')

        # ---------------- AI STYLE SUMMARY TEXT ----------------      
        original_df = pd.read_excel(request.files['file'])
        original_columns = len(original_df.columns)

        summary_text = f"""
        The uploaded file '{original_filename}' originally contained 
        {original_columns} columns.

        After cleaning and standardization, the processed dataset 
        ('{excel_filename}') contains {num_rows} rows and {num_columns} columns.

        Data Quality Score: {quality_score}%.
        Duplicate Rows Removed: {duplicate_rows}.
        Total Null Values: {total_nulls}.
        """

        if not numeric_df.empty:
            top_metric = numeric_df.mean().idxmax()
            summary_text += f" Highest average numeric column is {top_metric}."

        if "COUNTRY" in df.columns:
            top_country = df["COUNTRY"].value_counts().idxmax()
            summary_text += f" Most frequent country is {top_country}."

        # ---------------- PDF GENERATION ----------------
        pdf_filename = f"report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"

        pdf_buffer = io.BytesIO()
        doc = SimpleDocTemplate(pdf_buffer)
        elements = []
        styles = getSampleStyleSheet()

        elements.append(Paragraph("Excel Data Analysis Report", styles['Title']))
        elements.append(Spacer(1, 0.3 * inch))

        elements.append(Paragraph(f"Original File: {original_filename}", styles['Normal']))
        elements.append(Paragraph(f"Processed Excel File: {excel_filename}", styles['Normal']))
        elements.append(Paragraph(f"Generated PDF File: {pdf_filename}", styles['Normal']))
        elements.append(Paragraph(f"Generated On: {datetime.datetime.now()}", styles['Normal']))
        elements.append(Spacer(1, 0.4 * inch))

        elements.append(Paragraph(summary_text, styles['Normal']))
        elements.append(Spacer(1, 0.5 * inch))

        table_data = summary_df.values.tolist()
        table_data.insert(0, list(summary_df.columns))

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
            "summary_text": summary_text,
            "excel_filename": excel_filename,
            "pdf_filename": pdf_filename
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500