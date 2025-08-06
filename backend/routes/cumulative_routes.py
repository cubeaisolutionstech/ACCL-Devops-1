import os
import traceback
import logging
from flask import Blueprint, request, jsonify, send_file, current_app
import pandas as pd
from io import BytesIO
from werkzeug.utils import secure_filename

api_bp = Blueprint('api', __name__)

logger = logging.getLogger(__name__)
UPLOAD_FOLDER = "uploads"
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def read_excel_file(file_path, month, filename, skip_first_row=False):
    try:
        df = pd.read_excel(file_path, header=None if skip_first_row else 0, engine="openpyxl")
        if skip_first_row and not df.empty:
            df.columns = df.iloc[0]
            df = df.iloc[1:]
        if not df.empty:
            df.columns = [str(col).strip() for col in df.columns]

            # Do NOT override month if it's already in file
            if "Month" not in df.columns:
                df["Month"] = month

            df["Source File"] = filename
            return df
        return None
    except Exception as e:
        logger.error(f"Error reading {file_path}: {str(e)}")
        return None

@api_bp.route("/process", methods=["POST"])
def process_files():
    try:
        skip_first_row = request.form.get('skipFirstRow', 'false').lower() == 'true'
        files = request.files
        all_dfs = []
        errors = []
        uploaded_months = []

        # ONLY iterate through uploaded files
        for month, file in files.items():
            if file and file.filename.lower().endswith(('.xlsx', '.xls')):
                try:
                    filename = secure_filename(file.filename)
                    temp_path = os.path.join(current_app.config["UPLOAD_FOLDER"], filename)
                    file.save(temp_path)

                    df = read_excel_file(temp_path, month, filename, skip_first_row)
                    if df is not None:
                        all_dfs.append(df)
                        uploaded_months.append(month)

                    os.remove(temp_path)
                except Exception as e:
                    errors.append(f"Error processing {month}: {str(e)}")

        if not all_dfs:
            return jsonify({
                "success": False,
                "message": "No valid files could be processed",
                "errors": errors
            }), 400

        final_df = pd.concat(all_dfs, ignore_index=True)
        preview_data = final_df.head(5).replace({pd.NA: None}).to_dict(orient="records")

        return jsonify({
            "success": True,
            "message": f"Processed {len(all_dfs)} files",
            "preview": preview_data,
            "uploaded_months": uploaded_months,
            "warnings": errors if errors else None
        })

    except Exception as e:
        logger.error(f"Processing error: {str(e)}\n{traceback.format_exc()}")
        return jsonify({
            "success": False,
            "message": f"Server error: {str(e)}",
            "error": str(e)
        }), 500


@api_bp.route("/download", methods=["POST"])
def download_file():
    try:
        skip_first_row = request.form.get('skipFirstRow', 'false').lower() == 'true'
        files = request.files
        all_dfs = []
        errors = []

        # ONLY iterate through uploaded files
        for month, file in files.items():
            if file and file.filename.lower().endswith(('.xlsx', '.xls')):
                filename = secure_filename(file.filename)
                file_path = os.path.join(current_app.config["UPLOAD_FOLDER"], filename)
                file.save(file_path)
                df = read_excel_file(file_path, month, filename, skip_first_row)
                if df is not None and not df.empty:
                    all_dfs.append(df)
                os.remove(file_path)

        if not all_dfs:
            return jsonify({
                "success": False,
                "message": "No valid files to process.",
                "errors": errors
            }), 400

        combined_df = pd.concat(all_dfs, ignore_index=True)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            combined_df.to_excel(writer, index=False, sheet_name="Combined Sales")
        output.seek(0)

        return send_file(
            output,
            download_name="Combined_Sales_Report.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True
        )

    except Exception as e:
        logger.error(f"Download error: {str(e)}")
        return jsonify({
            "success": False,
            "message": f"Error generating file: {str(e)}",
            "errors": errors
        }), 500
