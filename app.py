from flask import Flask, request, send_file, jsonify
import pandas as pd
import io

app = Flask(__name__)

@app.route("/")
def home():
    return "Excel API is running!"

#@app.route("/process-excel", methods=["POST"])
#def process_excel():

#    if 'file' not in request.files:
#        return jsonify({"error": "No file uploaded"}), 400

#    file = request.files['file']

#    try:
#        df = pd.read_excel(file)

#        # Example processing 
        
        # Remove completely empty rows
#        df.dropna(how='all', inplace=True)     

        # Clean column names
#        df.columns = [col.strip().upper() for col in df.columns]

        # Remove pandas duplicate suffix (.1, .2 etc.)
#        df.columns = df.columns.str.replace(r'\.\d+$', '', regex=True)

        # Remove duplicate columns (keep first occurrence)
#        df = df.loc[:, ~df.columns.duplicated()]

        # Remove duplicate rows
#        df = df.drop_duplicates()

#        output = io.BytesIO()
#        df.to_excel(output, index=False)
#        output.seek(0)

#        return send_file(
#            output,
#            download_name="processed.xlsx",
#            as_attachment=True,
#            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#        )

#    except Exception as e:
#        return jsonify({"error": str(e)}), 500
    

@app.route("/process-excel", methods=["POST"])
def process_excel():

    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']

    try:
        df = pd.read_excel(file)

        # ---------------- CLEANING ----------------
        df.dropna(how='all', inplace=True)
        df.columns = [col.strip().upper() for col in df.columns]
        df.columns = df.columns.str.replace(r'\.\d+$', '', regex=True)
        df = df.loc[:, ~df.columns.duplicated()]

        duplicate_rows = df.duplicated().sum()
        df = df.drop_duplicates()

        # ---------------- BASIC STATS ----------------
        num_rows = len(df)
        num_columns = len(df.columns)
        column_names = ", ".join(df.columns)

        null_counts = df.isnull().sum()

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

        # ---------------- COUNTRY FREQUENCY ----------------
        country_freq = pd.DataFrame()
        if "COUNTRY" in df.columns:
            country_freq = df["COUNTRY"].value_counts().reset_index()
            country_freq.columns = ["Country", "Count"]

        # ---------------- DATA QUALITY SCORE ----------------
        total_cells = df.size
        total_nulls = df.isnull().sum().sum()
        quality_score = round((1 - (total_nulls / total_cells)) * 100, 2) if total_cells > 0 else 100

        summary_df = pd.DataFrame({
            "Metric": [
                "Number of Rows",
                "Number of Columns",
                "Duplicate Rows Removed",
                "Data Quality Score (%)",
                "Column Names"
            ],
            "Value": [
                num_rows,
                num_columns,
                duplicate_rows,
                quality_score,
                column_names
            ]
        })

        null_df = null_counts.reset_index()
        null_df.columns = ["Column", "Null Count"]

        # ---------------- WRITE TO EXCEL ----------------
        output = io.BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="DATA", index=False)
            summary_df.to_excel(writer, sheet_name="SUMMARY", index=False)
            stats_df.to_excel(writer, sheet_name="NUMERIC_STATS", index=True)
            null_df.to_excel(writer, sheet_name="NULL_COUNTS", index=False)

            if not country_freq.empty:
                country_freq.to_excel(writer, sheet_name="COUNTRY_FREQ", index=False)

        output.seek(0)

        return send_file(
            output,
            download_name="processed_with_analytics.xlsx",
            as_attachment=True,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500

