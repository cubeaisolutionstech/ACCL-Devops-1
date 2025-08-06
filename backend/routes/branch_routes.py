from flask import Blueprint, request, jsonify, send_file
import os
import pandas as pd
import datetime
import logging
import json 
from werkzeug.utils import secure_filename
from utils.budget_vs_billed import calculate_budget_vs_billed, auto_map_budget_columns
from utils.ppt_generator import generate_budget_ppt, create_od_ppt_updated,create_product_growth_ppt,create_nbc_individual_ppt,create_od_individual_ppt,create_consolidated_ppt
from utils.od_target import auto_map_od_columns,calculate_od_values_updated,create_region_branch_mapping,create_dynamic_regional_summary,get_cumulative_branches,get_cumulative_regions
from utils.product_growth import calculate_product_growth,auto_map_product_growth_columns,standardize_name
from utils.nbc_od_utils import auto_map_nbc_columns,auto_map_od_target_columns,create_customer_table,filter_os_qty,nbc_branch_mapping

branch_bp = Blueprint('branch', __name__)
logger = logging.getLogger(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs("static", exist_ok=True)


@branch_bp.route('/upload', methods=['POST'])
def upload_file():
    file = request.files.get('file')
    if not file:
        return jsonify({'error': 'No file uploaded'}), 400
    filename = secure_filename(file.filename)
    path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(path)
    return jsonify({'message': 'File uploaded successfully', 'filename': filename})

@branch_bp.route('/sheets', methods=['POST'])
def get_sheet_names():
    file = request.json.get('filename')
    path = os.path.join(UPLOAD_FOLDER, file)
    try:
        xl = pd.ExcelFile(path)
        return jsonify({'sheets': xl.sheet_names})
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
@branch_bp.route('/get_columns', methods=['POST'])
def get_columns_and_auto_map():
    data = request.json
    path = os.path.join('uploads', data['filename'])
    header_row = int(data['header']) - 1  # user sends 1-based
    sheet_name = data['sheet_name']

    try:
        df = pd.read_excel(path, sheet_name=sheet_name, header=header_row)
        columns = df.columns.tolist()
        return jsonify({'columns': columns})
    except Exception as e:
        import traceback
        logger.error(f"❌ Error in get_columns: {traceback.format_exc()}")
        return jsonify({'error': str(e)}), 500
 
@branch_bp.route('/auto_map_columns', methods=['POST'])
def auto_map_columns():
    try:
        data = request.json
        sales_cols = data.get('sales_columns', [])
        budget_cols = data.get('budget_columns', [])
        sales_mapping, budget_mapping = auto_map_budget_columns(sales_cols, budget_cols)
        return jsonify({
            'sales_mapping': sales_mapping,
            'budget_mapping': budget_mapping
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
@branch_bp.route('/get_exec_branch_options', methods=['POST'])
def get_exec_branch_options():
    try:
        data = request.json
        sales_path = os.path.join('uploads', data['sales_filename'])
        budget_path = os.path.join('uploads', data['budget_filename'])

        sales_df = pd.read_excel(sales_path, sheet_name=data['sales_sheet'], header=data['sales_header'] - 1)
        budget_df = pd.read_excel(budget_path, sheet_name=data['budget_sheet'], header=data['budget_header'] - 1)

        exec_sales_col = data['sales_exec_col']
        exec_budget_col = data['budget_exec_col']
        area_sales_col = data['sales_area_col']
        area_budget_col = data['budget_area_col']

        from utils.budget_vs_billed import map_branch

        sales_execs = sorted(sales_df[exec_sales_col].dropna().unique().tolist())
        budget_execs = sorted(budget_df[exec_budget_col].dropna().unique().tolist())

        combined = pd.concat([sales_df[area_sales_col], budget_df[area_budget_col]], ignore_index=True).dropna()
        branches = sorted(set(map(map_branch, combined)))

        

        return jsonify({
            'sales_executives': sales_execs,
            'budget_executives': budget_execs,
            'branches': branches
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
@branch_bp.route('/extract_months', methods=['POST'])
def extract_months():
    try:
        data = request.json
        sales_path = os.path.join('uploads', data['sales_filename'])
        df = pd.read_excel(sales_path, sheet_name=data['sales_sheet'], header=data['sales_header'] - 1)
        date_col = data['sales_date_col']

        months = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce').dt.strftime('%b %y').dropna().unique().tolist()
        return jsonify({'months': months})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@branch_bp.route('/calculate_budget_vs_billed', methods=['POST'])
def budget_vs_billed():
    data = request.json
    try:
        results = calculate_budget_vs_billed(data)
        return jsonify(results)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@branch_bp.route('/generate_consolidated_branch_ppt', methods=['POST'])
def generate_consolidated_branch_ppt():
    try:
        data = request.get_json()
        report_title = data.get('reportTitle', 'ACCLLP Consolidated Report')
        all_dfs_info = data.get('allDfsInfo', [])
        

        parsed_sections = []
        for item in all_dfs_info:
            df_data = item.get("df")
            title = item.get("title", "Untitled Section")
            percent_cols = item.get("percent_cols", [])
            if df_data:
                df = pd.DataFrame(df_data)
                parsed_sections.append({
                    "title": title,
                    "df": df,
                    "percent_cols": percent_cols
                })

        # Optional logo - for now, ignore unless you want to add a file upload
        ppt_stream = create_consolidated_ppt(parsed_sections, logo_file=None, title=report_title)

        return send_file(
            ppt_stream,
            as_attachment=True,
            download_name=f"{report_title.replace(' ', '_')}.pptx",
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500
    
@branch_bp.route('/download_ppt', methods=['POST'])
def download_ppt():
    try:
        payload = request.get_json()
        print("📦 PPT Payload Keys:", list(payload.keys()))
        ppt_path = generate_budget_ppt(payload)
        return send_file(ppt_path, as_attachment=True)
    except Exception as e:
        print("❌ PPT generation failed:", str(e))
        return jsonify({'error': str(e)}), 500
    
@branch_bp.route('/get_od_columns', methods=['POST'])
def get_od_columns():
    data = request.get_json()
    os_prev_cols = data.get('os_prev_columns', [])
    os_curr_cols = data.get('os_curr_columns', [])
    sales_cols = data.get('sales_columns', [])

    try:
        os_jan_map, os_feb_map, sales_map = auto_map_od_columns(os_prev_cols, os_curr_cols, sales_cols)
        return jsonify({
            'os_jan_mapping': os_jan_map,
            'os_feb_mapping': os_feb_map,
            'sales_mapping': sales_map
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
@branch_bp.route('/get_od_filter_options', methods=['POST'])
def get_od_filter_options():
    data = request.get_json()
    try:
        os_prev = pd.read_excel(f"uploads/{data['os_prev_filename']}", sheet_name=data['os_prev_sheet'], header=data['os_prev_header'] - 1)
        os_curr = pd.read_excel(f"uploads/{data['os_curr_filename']}", sheet_name=data['os_curr_sheet'], header=data['os_curr_header'] - 1)
        sales = pd.read_excel(f"uploads/{data['sales_filename']}", sheet_name=data['sales_sheet'], header=data['sales_header'] - 1)

        # Pull mapped column names
        os_prev_exec = data['os_prev_mapping'].get('executive')
        os_curr_exec = data['os_curr_mapping'].get('executive')
        sales_exec = data['sales_mapping'].get('executive')

        os_prev_branch = data['os_prev_mapping'].get('branch')
        os_curr_branch = data['os_curr_mapping'].get('branch')
        sales_branch = data['sales_mapping'].get('branch')

        os_prev_region = data['os_prev_mapping'].get('region')
        os_curr_region = data['os_curr_mapping'].get('region')
        sales_region = data['sales_mapping'].get('region')

        execs = set()
        branches = set()
        regions = set()

        if os_prev_exec in os_prev.columns:
            execs.update(os_prev[os_prev_exec].dropna().unique())
        if os_curr_exec in os_curr.columns:
            execs.update(os_curr[os_curr_exec].dropna().unique())
        if sales_exec in sales.columns:
            execs.update(sales[sales_exec].dropna().unique())

        if os_prev_branch in os_prev.columns:
            branches.update(os_prev[os_prev_branch].dropna().unique())
        if os_curr_branch in os_curr.columns:
            branches.update(os_curr[os_curr_branch].dropna().unique())
        if sales_branch in sales.columns:
            branches.update(sales[sales_branch].dropna().unique())

        if os_prev_region and os_prev_region in os_prev.columns:
            regions.update(os_prev[os_prev_region].dropna().unique())
        if os_curr_region and os_curr_region in os_curr.columns:
            regions.update(os_curr[os_curr_region].dropna().unique())
        if sales_region and sales_region in sales.columns:
            regions.update(sales[sales_region].dropna().unique())

        return jsonify({
            "executives": sorted(list(execs)),
            "branches": sorted(list(branches)),
            "regions": sorted(list(regions))
        })

    except Exception as e:
        return jsonify({'error': f'Failed to fetch filters: {str(e)}'}), 500

@branch_bp.route('/calculate_od_target', methods=['POST'])
def calculate_od_target():
    try:
        data = request.get_json()

        print("🔧 OD Calculation Payload Keys:", list(data.keys()))
        print("🧩 Mappings Preview (os_prev):", data.get("os_prev_mapping", {}))
        print("🧩 Mappings Preview (sales):", data.get("sales_mapping", {}))
        print("📄 Filenames:", data['os_prev_filename'], data['sales_filename'])

        # Load the 3 Excel files
        os_prev = pd.read_excel(f"uploads/{data['os_prev_filename']}", sheet_name=data['os_prev_sheet'], header=data['os_prev_header'] - 1)
        os_curr = pd.read_excel(f"uploads/{data['os_curr_filename']}", sheet_name=data['os_curr_sheet'], header=data['os_curr_header'] - 1)
        sales = pd.read_excel(f"uploads/{data['sales_filename']}", sheet_name=data['sales_sheet'], header=data['sales_header'] - 1)

        # === Apply filters and calculate final output ===
        final, regional, region_map = calculate_od_values_updated(
            os_prev, os_curr, sales,
            data['selected_month'],
            data['os_prev_mapping']['due_date'],
            data['os_prev_mapping']['ref_date'],
            data['os_prev_mapping']['branch'],
            data['os_prev_mapping']['net_value'],
            data['os_prev_mapping']['executive'],
            data['os_prev_mapping'].get('region'),
            data['os_curr_mapping']['due_date'],
            data['os_curr_mapping']['ref_date'],
            data['os_curr_mapping']['branch'],
            data['os_curr_mapping']['net_value'],
            data['os_curr_mapping']['executive'],
            data['os_curr_mapping'].get('region'),
            data['sales_mapping']['bill_date'],
            data['sales_mapping']['due_date'],
            data['sales_mapping']['branch'],
            data['sales_mapping']['value'],
            data['sales_mapping']['executive'],
            data['sales_mapping'].get('region'),
            data.get('selected_executives', []),
            data.get('selected_branches', []),
            data.get('selected_regions', [])
        )

        return jsonify({
            "branch_summary": final.to_dict(orient='records') if final is not None else [],
            "regional_summary": regional.to_dict(orient='records') if regional is not None else [],
            "region_mapping": region_map
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'OD Target calculation failed: {str(e)}'}), 500

@branch_bp.route('/download_od_ppt', methods=['POST'])
def download_od_ppt():
    try:
        data = request.get_json()

        branch_data = data.get('branch_summary', {}).get('data', [])
        branch_columns = data.get('branch_summary', {}).get('columns', [])

        regional_data = data.get('regional_summary', {}).get('data', [])
        regional_columns = data.get('regional_summary', {}).get('columns', [])

        title = data.get('title', 'OD Target vs Collection')

        def to_ordered_df(data, columns):
            df = pd.DataFrame(data)
            if columns:
                df = df[columns]
            for col in df.columns:
                if pd.api.types.is_numeric_dtype(df[col]):
                    df[col] = df[col].round(0).astype('Int64')
            return df
        
        branch_df = to_ordered_df(branch_data, branch_columns)
        regional_df = to_ordered_df(regional_data, regional_columns)

        ppt_buffer = create_od_ppt_updated(branch_df, regional_df, title)

        if ppt_buffer:
            return send_file(
                ppt_buffer,
                mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                as_attachment=True,
                download_name=f"OD_Target_vs_Collection_{title}.pptx"
            )
        else:
            return jsonify({'error': 'PPT generation failed'}), 500
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@branch_bp.route('/auto_map_product_growth', methods=['POST'])
def auto_map_pg_columns():
    data = request.get_json()
    ly_mapping, cy_mapping, budget_mapping = auto_map_product_growth_columns(
        data['ly_columns'], data['cy_columns'], data['budget_columns']
    )
    return jsonify({
        'ly_mapping': ly_mapping,
        'cy_mapping': cy_mapping,
        'budget_mapping': budget_mapping
    })

@branch_bp.route("/get_product_growth_filters", methods=["POST"])
def get_product_growth_filters():
    try:
        data = request.get_json()

        ly_filename = data["ly_filename"]
        cy_filename = data["cy_filename"]
        budget_filename = data["budget_filename"]

        ly_sheet = data["ly_sheet"]
        cy_sheet = data["cy_sheet"]
        budget_sheet = data["budget_sheet"]

        ly_header = data["ly_header"] - 1
        cy_header = data["cy_header"] - 1
        budget_header = data["budget_header"] - 1

        # Column names provided from frontend
        ly_date_col = data["ly_date_col"]
        cy_date_col = data["cy_date_col"]

        ly_exec_col = data["ly_exec_col"]
        cy_exec_col = data["cy_exec_col"]
        budget_exec_col = data["budget_exec_col"]

        ly_group_col = data["ly_group_col"]
        cy_group_col = data["cy_group_col"]
        budget_group_col = data["budget_group_col"]

        # Load DataFrames
        ly_df = pd.read_excel(f"uploads/{ly_filename}", sheet_name=ly_sheet, header=ly_header)
        cy_df = pd.read_excel(f"uploads/{cy_filename}", sheet_name=cy_sheet, header=cy_header)
        budget_df = pd.read_excel(f"uploads/{budget_filename}", sheet_name=budget_sheet, header=budget_header)

        # ✅ Extract unique months from LY and CY
        ly_months = pd.to_datetime(ly_df[ly_date_col], dayfirst=True, errors="coerce").dt.strftime("%b %y").dropna().unique().tolist()
        cy_months = pd.to_datetime(cy_df[cy_date_col], dayfirst=True, errors="coerce").dt.strftime("%b %y").dropna().unique().tolist()

        # ✅ Get unique executives
        executives = pd.concat([
            ly_df[ly_exec_col].dropna(),
            cy_df[cy_exec_col].dropna(),
            budget_df[budget_exec_col].dropna()
        ]).dropna().astype(str).unique().tolist()

        # ✅ Get unique company groups
        all_groups = pd.concat([
            ly_df[ly_group_col].dropna(),
            cy_df[cy_group_col].dropna(),
            budget_df[budget_group_col].dropna()
        ]).dropna().astype(str).map(standardize_name)

        company_groups = sorted(set(all_groups))

        return jsonify({
            "ly_months": sorted(ly_months),
            "cy_months": sorted(cy_months),
            "executives": sorted(executives),
            "company_groups": sorted(company_groups)
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500

@branch_bp.route('/calculate_product_growth', methods=['POST'])
def calculate_product_growth_api():
    try:
        data = request.get_json()

        # === Load Excel files directly from 'uploads/' ===
        ly_df = pd.read_excel(f"uploads/{data['ly_filename']}", sheet_name=data['ly_sheet'], header=data['ly_header'] - 1)
        cy_df = pd.read_excel(f"uploads/{data['cy_filename']}", sheet_name=data['cy_sheet'], header=data['cy_header'] - 1)
        budget_df = pd.read_excel(f"uploads/{data['budget_filename']}", sheet_name=data['budget_sheet'], header=data['budget_header'] - 1)

        result = calculate_product_growth(
            ly_df, cy_df, budget_df,
            data['ly_months'], data['cy_months'],
            data['ly_date_col'], data['cy_date_col'],
            data['ly_qty_col'], data['cy_qty_col'],
            data['ly_value_col'], data['cy_value_col'],
            data['budget_qty_col'], data['budget_value_col'],
            data['ly_product_col'], data['cy_product_col'],
            data['ly_company_group_col'], data['cy_company_group_col'],
            data['budget_company_group_col'], data['budget_product_group_col'],
            data['ly_exec_col'], data['cy_exec_col'], data['budget_exec_col'],
            data.get('selected_executives', []),
            data.get('selected_company_groups', [])
        )
        if not result:
            return jsonify({'status': 'error', 'error': 'Calculation failed or no data returned'}), 400


        # Convert DataFrames to JSON
        response = {
            group: {
                'qty_df': df_pair['qty_df'].to_dict(orient='records'),
                'value_df': df_pair['value_df'].to_dict(orient='records')
            } for group, df_pair in result.items()
        }

        return jsonify({'status': 'success', 'results': response})
    except Exception as e:
        import traceback
        print("❌ Error in calculate_product_growth:", e)
        print(traceback.format_exc())
        return jsonify({'status': 'error', 'error': str(e)}), 500

@branch_bp.route('/download_product_growth_ppt', methods=['POST'])
def download_product_growth_ppt():
    try:
        data = request.get_json()
        group_results = data.get("group_results", {})
        month_title = data.get("month_title", "Product Growth")

        parsed_results = {}

        for group, content in group_results.items():
            # ✅ If content is a JSON string, parse it
            if isinstance(content, str):
                try:
                    content = json.loads(content)
                except Exception as parse_err:
                    logger.warning(f"⚠️ Failed to parse JSON for group '{group}': {parse_err}")
                    continue

            # ✅ Safely extract DataFrames
            qty_df = pd.DataFrame(content.get("qty_df", []))
            value_df = pd.DataFrame(content.get("value_df", []))

            if qty_df.empty or value_df.empty:
                logger.warning(f"⚠️ Empty DataFrame for group: {group}")
                continue

            parsed_results[group] = {
                "qty_df": qty_df,
                "value_df": value_df
            }

        if not parsed_results:
            return jsonify({"error": "No data available to generate PPT"}), 400

        ppt_buffer = create_product_growth_ppt(parsed_results, month_title)

        if ppt_buffer:
            return send_file(
                ppt_buffer,
                mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                as_attachment=True,
                download_name=f"Product_Growth_{month_title.replace(' ', '_')}.pptx"
            )
        else:
            return jsonify({"error": "PPT generation failed"}), 500

    except Exception as e:
        import traceback
        logger.error("❌ Error in download_product_growth_ppt: %s", traceback.format_exc())
        return jsonify({"error": str(e)}), 500

@branch_bp.route('/get_nbc_columns', methods=['POST'])
def get_nbc_columns():
    try:
        data = request.get_json()
        filename = data['filename']
        sheet_name = data['sheet_name']
        header = data['header'] - 1

        df = pd.read_excel(f"uploads/{filename}", sheet_name=sheet_name, header=header)
        columns = df.columns.tolist()
        mapped = auto_map_nbc_columns(columns)

        return jsonify({"columns": columns, "mapping": mapped})
    except Exception as e:
        import traceback
        print("❌ Error in get_nbc_columns:", traceback.format_exc())
        return jsonify({"error": str(e)}), 500
    
@branch_bp.route("/get_nbc_filters", methods=["POST"])
def get_nbc_filters():
    try:
        data = request.get_json()
        filename = data["filename"]
        sheet_name = data["sheet_name"]
        header_row = data["header"] - 1

        date_col = data["date_col"]
        branch_col = data["branch_col"]
        executive_col = data["executive_col"]

        df = pd.read_excel(f"uploads/{filename}", sheet_name=sheet_name, header=header_row)

        # Clean column names
        df.columns = [str(col).strip() for col in df.columns]

        # Basic validation
        for col in [date_col, branch_col, executive_col]:
            if col not in df.columns:
                return jsonify({"error": f"Column '{col}' not found in data"}), 400

        # Normalize branch names
        raw_branches = df[branch_col].dropna().astype(str).str.upper().unique().tolist()
        all_branches = sorted(set([
            nbc_branch_mapping.get(branch.split(" - ")[-1], branch.split(" - ")[-1])
            for branch in raw_branches
        ]))

        all_executives = sorted(df[executive_col].dropna().astype(str).unique().tolist())

        return jsonify({
            "branches": all_branches,
            "executives": all_executives
        })

    except Exception as e:
        import traceback
        logger.error(f"❌ Error in get_nbc_filters: {traceback.format_exc()}")
        return jsonify({"error": str(e)}), 500


@branch_bp.route('/calculate_nbc_table', methods=['POST'])
def calculate_nbc_table():
    try:
        data = request.get_json()

        filename = data["filename"]
        sheet_name = data["sheet_name"]
        header = data["header"] - 1

        date_col = data["date_col"]
        customer_id_col = data["customer_id_col"]
        branch_col = data["branch_col"]
        executive_col = data["executive_col"]

        selected_branches = data.get("selected_branches", [])
        selected_executives = data.get("selected_executives", [])

        df = pd.read_excel(f"uploads/{filename}", sheet_name=sheet_name, header=header)

        results = create_customer_table(
            df, date_col, branch_col, customer_id_col, executive_col,
            selected_branches=selected_branches,
            selected_executives=selected_executives
        )

        final_output = {}
        for fy, (result_df, sorted_months) in results.items():
            final_output[fy] = {
                "data": result_df.to_dict(orient="records"),
                "months": sorted_months
            }

        return jsonify({"results": final_output})
    except Exception as e:
        import traceback
        logger.error(f"❌ Error in calculate_nbc_table: {traceback.format_exc()}")
        return jsonify({"error": str(e)}), 500

@branch_bp.route('/download_nbc_ppt', methods=['POST'])
def download_nbc_ppt():
    try:
        data = request.get_json()
        df_data = data.get("data", [])
        title = data.get("title", "NBC Report")
        months = data.get("months", [])
        fy = data.get("financial_year", "")
        logo_filename = data.get("logo_file")

        if not df_data:
            return jsonify({"error": "No data provided"}), 400

        df = pd.DataFrame(df_data)

        logo_path = f"uploads/{logo_filename}" if logo_filename else None
        ppt_buffer = create_nbc_individual_ppt(df, title, months, fy, logo_path)

        if ppt_buffer:
            return send_file(
                ppt_buffer,
                mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                as_attachment=True,
                download_name=f"{title.replace(' ', '_')}_{fy}.pptx"
            )
        else:
            return jsonify({"error": "Failed to generate PPT"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@branch_bp.route("/get_od_target_columns", methods=["POST"])
def get_od_target_columns():
    try:
        data = request.get_json()
        columns = data.get("columns", [])
        mapping = auto_map_od_target_columns(columns)
        return jsonify({"mapping": mapping})
    except Exception as e:
        logger.error(f"❌ Error in get_od_target_columns: {e}")
        return jsonify({"error": str(e)}), 500
@branch_bp.route('/get_column_unique_values', methods=['POST'])
def get_column_unique_values():
    try:
        data = request.get_json()
        filename = data["filename"]
        sheet_name = data["sheet_name"]
        header = data["header"] - 1
        column_names = data.get("column_names", [])

        df = pd.read_excel(f"uploads/{filename}", sheet_name=sheet_name, header=header)

        response = {}

        for col in column_names:
            if col not in df.columns:
                continue

            col_data = df[col].dropna()

            # Detect if this column is due date, area, or executive
            lower_col = col.lower()
            if "date" in lower_col:
                col_data = pd.to_datetime(col_data, errors='coerce')
                valid_years = sorted(col_data.dt.year.dropna().astype(int).unique())
                response[col] = {"years": [str(y) for y in valid_years]}
            elif "area" in lower_col or "branch" in lower_col:
                from utils.nbc_od_utils import extract_area_name
                cleaned = sorted(set(filter(None, col_data.map(extract_area_name))))
                response[col] = {"values": cleaned}
            elif "executive" in lower_col:
                from utils.nbc_od_utils import extract_executive_name
                cleaned = sorted(set(filter(None, col_data.map(extract_executive_name))))
                response[col] = {"values": cleaned}
            else:
                cleaned = sorted(set(col_data.astype(str).str.strip().unique()))
                response[col] = {"values": cleaned}

        return jsonify(response)

    except Exception as e:
        logger.error(f"❌ Error in get_column_unique_values: {e}")
        return jsonify({"error": str(e)}), 500

@branch_bp.route('/calculate_od_target_table', methods=['POST'])
def calculate_od_target_table():
    try:
        data = request.get_json()
        filename = data["filename"]
        sheet_name = data["sheet_name"]
        header = data["header"] - 1

        area_col = data["area_col"]
        due_date_col = data["due_date_col"]
        qty_col = data["qty_col"]
        executive_col = data["executive_col"]

        selected_branches = data.get("selected_branches", [])
        selected_executives = data.get("selected_executives", [])
        selected_years = data.get("selected_years", [])
        till_month = data.get("till_month")

        df = pd.read_excel(f"uploads/{filename}", sheet_name=sheet_name, header=header)

        result_df, start_date, end_date = filter_os_qty(
            df,
            os_area_col=area_col,
            os_qty_col=qty_col,
            os_due_date_col=due_date_col,
            os_exec_col=executive_col,
            selected_branches=selected_branches,
            selected_years=selected_years,
            till_month=till_month,
            selected_executives=selected_executives
        )

        if result_df is None:
            return jsonify({"error": "Failed to generate OD Target table"}), 400

        start_str = start_date.strftime('%b %Y') if start_date else "Start"
        end_str = end_date.strftime('%b %Y') if end_date else "End"

        return jsonify({
            "table": result_df.to_dict(orient="records"),
            "start": start_str,
            "end": end_str
        })

    except Exception as e:
        logger.error(f"❌ Error in calculate_od_target_table: {e}")
        return jsonify({"error": str(e)}), 500


@branch_bp.route('/download_od_target_ppt', methods=['POST'])
def download_od_target_ppt():
    try:
        data = request.get_json()

        df_data = data.get("result", [])
        title = data.get("title", "OD Target")
        logo_filename = data.get("logo_file")

        if not df_data:
            return jsonify({"error": "No data provided"}), 400

        df = pd.DataFrame(df_data)
        logo_path = f"uploads/{logo_filename}" if logo_filename else None
        ppt_buffer = create_od_individual_ppt(df, title, logo_path)

        if ppt_buffer:
            return send_file(
                ppt_buffer,
                mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                as_attachment=True,
                download_name=f"{title.replace(' ', '_')}.pptx"
            )
        else:
            return jsonify({"error": "Failed to generate PPT"}), 500
    except Exception as e:
        import traceback
        logger.error("❌ Error in download_od_target_ppt: %s", traceback.format_exc())
        return jsonify({"error": str(e)}), 500
