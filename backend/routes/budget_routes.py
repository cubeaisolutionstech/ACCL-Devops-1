from flask import Blueprint, request, jsonify, send_file
import pandas as pd
from io import BytesIO
from services.mapping_service import process_budget_file
import base64

budget_bp = Blueprint("budget", __name__)

@budget_bp.route("/upload-budget-file", methods=["POST"])
def upload_budget_file():
    file = request.files["file"]
    sheet_name = request.form["sheet_name"]
    header_row = int(request.form["header_row"])

    df = pd.read_excel(file, sheet_name=sheet_name, header=header_row)

    # Get mapped columns
    customer_col = request.form["customer_col"]
    exec_code_col = request.form["exec_code_col"]
    exec_name_col = request.form["exec_name_col"]
    branch_col = request.form["branch_col"]
    region_col = request.form["region_col"]
    cust_name_col = request.form.get("cust_name_col") or None

    processed = process_budget_file(df, customer_col, exec_code_col, exec_name_col, branch_col, region_col, cust_name_col)
    

    # Prepare download file
    output = BytesIO()
    processed.to_excel(output, index=False)
    output.seek(0)
    encoded_excel = base64.b64encode(output.read()).decode("utf-8")

    # Return preview and stats
    return {
        "preview": processed.head(10).fillna("").to_dict(orient="records"),
        "counts": {
            "total": int(len(processed)),
            "execs": int(processed["Executive Name"].astype(bool).sum()),
            "branches": int((processed["Branch"] != "").sum()),
            "regions": int((processed["Region"] != "").sum()),
        },
        "columns": list(processed.columns),
        "file_data": encoded_excel,
        
    }


@budget_bp.route("/save-budget-file", methods=["POST"])
def save_budget_file():
    data = request.get_json()
    base64_excel = data.get("file_data")
    filename = data.get("filename", "Processed_Budget.xlsx")

    if not filename.lower().endswith(".xlsx"):
        filename += ".xlsx"

    if not base64_excel:
        return jsonify({"error": "No file data provided"}), 400

    from base64 import b64decode
    from models.schema import BudgetFile
    from extensions import db

    binary_data = b64decode(base64_excel)
    budget_file = BudgetFile(filename=filename, file_data=binary_data)
    db.session.add(budget_file)
    db.session.commit()

    return jsonify({"message": "File saved", "id": budget_file.id})

@budget_bp.route("/budget-files", methods=["GET"])
def list_budget_files():
    from models.schema import BudgetFile
    files = BudgetFile.query.order_by(BudgetFile.uploaded_at.desc()).all()
    return jsonify([
        {"id": f.id, "filename": f.filename, "uploaded_at": f.uploaded_at.isoformat()}
        for f in files
    ])
    
@budget_bp.route("/budget-files/<int:file_id>/download", methods=["GET"])
def download_budget_file(file_id):
    from models.schema import BudgetFile
    import io
    file = BudgetFile.query.get_or_404(file_id)

    return send_file(
        io.BytesIO(file.file_data),
        download_name=file.filename,
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@budget_bp.route("/budget-files/<int:file_id>", methods=["DELETE"])
def delete_budget_file(file_id):
    from models.schema import BudgetFile
    os_file = BudgetFile.query.get(file_id)
    if not os_file:
        return jsonify({"error": "File not found"}), 404

    from extensions import db
    db.session.delete(os_file)
    db.session.commit()

    return jsonify({"message": "File deleted"})

@budget_bp.route("/budget-files/<int:file_id>/data", methods=["GET"])
def get_budget_file_data(file_id):
    from models.schema import BudgetFile
    from utils.file_utils import read_excel_from_binary  # Make sure this utility exists
    file = BudgetFile.query.get(file_id)
    if not file:
        return jsonify({"error": "Budget file not found"}), 404

    try:
        df = read_excel_from_binary(file.file_data)
        return jsonify(df.head(100).to_dict(orient="records"))
    except Exception as e:
        return jsonify({"error": f"Failed to read Excel: {str(e)}"}), 500
