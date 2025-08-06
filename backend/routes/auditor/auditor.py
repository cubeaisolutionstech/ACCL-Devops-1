from flask import Blueprint, request, jsonify, session
import pandas as pd
import re
import traceback
from utils.auditor.data_processor import DataProcessor

# Create the blueprint
auditor_bp = Blueprint('auditor', __name__)

def detect_analysis_type(sheet_name):
    """Detect the type of analysis based on sheet name"""
    sheet_lower = sheet_name.lower()
    
    is_region_analysis = 'region' in sheet_lower
    is_product_analysis = 'product' in sheet_lower or 'ts-pw' in sheet_lower or 'ero-pw' in sheet_lower
    is_sales_analysis_month_wise = bool(re.search(r'sales\s*analysis\s*month\s*wise', sheet_lower))
    
    return {
        'is_region_analysis': is_region_analysis,
        'is_product_analysis': is_product_analysis,
        'is_sales_analysis_month_wise': is_sales_analysis_month_wise
    }

def get_session_key(filepath, sheet_name, table_type):
    """Generate a unique session key for each table"""
    # Create a clean key from filepath and sheet name
    file_key = filepath.split('/')[-1].replace('.xlsx', '').replace('.xls', '')
    return f"auditor_{file_key}_{sheet_name}_{table_type}"

@auditor_bp.route('/process-auditor-auto', methods=['POST'])
def process_auditor_auto():
    """Auto-process auditor table and store both tables in session"""
    try:
        data = request.json
        filepath = data.get('filepath')
        sheet_name = data.get('sheet_name')
        table_choice = data.get('table_choice', 'Table 2: SALES in Value')
        force_refresh = data.get('force_refresh', False)  # Option to force reprocessing
        
        if not filepath or not sheet_name:
            return jsonify({'error': 'Missing filepath or sheet_name'}), 400
        
        # Check if both tables are already in session (unless force refresh)
        table1_key = get_session_key(filepath, sheet_name, 'Table1')
        table2_key = get_session_key(filepath, sheet_name, 'Table2')
        
        if not force_refresh and table1_key in session and table2_key in session:
            # Return the requested table from session
            if table_choice == "Table 1: SALES in MT/Tonage" and table1_key in session:
                stored_data = session[table1_key]
                return jsonify(stored_data)
            elif table_choice == "Table 2: SALES in Value" and table2_key in session:
                stored_data = session[table2_key]
                return jsonify(stored_data)
        
        # Initialize DataProcessor
        processor = DataProcessor()
        
        # Read the Excel sheet
        df_sheet = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
        
        # Detect analysis type
        analysis_info = detect_analysis_type(sheet_name)
        
        # Define possible headers for both tables
        table1_possible_headers = [
            "SALES in MT", "SALES IN MT", "Sales in MT", "SALES IN TONNAGE", "SALES IN TON",
            "Tonnage", "TONNAGE", "Tonnage Sales", "Sales Tonnage", "Metric Tons", "MT Sales"
        ]
        
        table2_possible_headers = [
            "SALES in Value", "SALES IN VALUE", "Sales in Value", "SALES IN RS", "VALUE SALES",
            "Value", "VALUE", "Sales Value"
        ]
        
        try:
            # Extract both tables using DataProcessor
            idx1, data_start1 = processor.extract_tables(
                df_sheet, 
                table1_possible_headers, 
                is_product_analysis=analysis_info['is_product_analysis']
            )
            
            idx2, data_start2 = processor.extract_tables(
                df_sheet, 
                table2_possible_headers, 
                is_product_analysis=analysis_info['is_product_analysis']
            )
            
            # Process and store Table 1 if found
            table1_data = None
            if idx1 is not None:
                table_end1 = idx2 if idx2 is not None and idx2 > idx1 else len(df_sheet)
                table1_df = df_sheet.iloc[data_start1:table_end1].dropna(how='all').reset_index(drop=True)
                
                if not table1_df.empty:
                    # Set headers and process data
                    original_headers1 = df_sheet.iloc[idx1].tolist()
                    table1_df.columns = [str(col) for col in original_headers1]
                    
                    # Process with DataProcessor
                    new_column_names1 = processor.rename_columns(table1_df.columns.tolist())
                    table1_df.columns = new_column_names1
                    table1_df = processor.handle_duplicate_columns(table1_df)
                    table1_df = processor.clean_and_convert_numeric(table1_df)
                    
                    # Get column information
                    act_columns1 = [col for col in table1_df.columns if col.startswith('Act-')]
                    gr_columns1 = [col for col in table1_df.columns if col.startswith('Gr-')]
                    ach_columns1 = [col for col in table1_df.columns if col.startswith('Ach-')]
                    budget_columns1 = [col for col in table1_df.columns if col.startswith('Budget-')]
                    ly_columns1 = [col for col in table1_df.columns if col.startswith('LY-')]
                    
                    # Prepare Table 1 data for session storage
                    table1_data = {
                        'success': True,
                        'table_name': 'Table 1: SALES in MT/Tonage',
                        'data': table1_df.to_dict('records'),
                        'columns': table1_df.columns.tolist(),
                        'shape': [int(table1_df.shape[0]), int(table1_df.shape[1])],
                        'analysis_type': analysis_info,
                        'column_info': {
                            'act_columns': act_columns1,
                            'gr_columns': gr_columns1,
                            'ach_columns': ach_columns1,
                            'budget_columns': budget_columns1,
                            'ly_columns': ly_columns1,
                            'total_columns': int(len(table1_df.columns))
                        },
                        'original_headers': [str(h) for h in original_headers1],
                        'processing_info': {
                            'original_column_count': len(original_headers1),
                            'final_column_count': len(table1_df.columns),
                            'processing_successful': True
                        }
                    }
                    
                    # Store in session
                    session[table1_key] = table1_data
            
            # Process and store Table 2 if found
            table2_data = None
            if idx2 is not None:
                table2_df = df_sheet.iloc[data_start2:len(df_sheet)].dropna(how='all').reset_index(drop=True)
                
                if not table2_df.empty:
                    # Set headers and process data
                    original_headers2 = df_sheet.iloc[idx2].tolist()
                    table2_df.columns = [str(col) for col in original_headers2]
                    
                    # Process with DataProcessor
                    new_column_names2 = processor.rename_columns(table2_df.columns.tolist())
                    table2_df.columns = new_column_names2
                    table2_df = processor.handle_duplicate_columns(table2_df)
                    table2_df = processor.clean_and_convert_numeric(table2_df)
                    
                    # Get column information
                    act_columns2 = [col for col in table2_df.columns if col.startswith('Act-')]
                    gr_columns2 = [col for col in table2_df.columns if col.startswith('Gr-')]
                    ach_columns2 = [col for col in table2_df.columns if col.startswith('Ach-')]
                    budget_columns2 = [col for col in table2_df.columns if col.startswith('Budget-')]
                    ly_columns2 = [col for col in table2_df.columns if col.startswith('LY-')]
                    
                    # Prepare Table 2 data for session storage
                    table2_data = {
                        'success': True,
                        'table_name': 'Table 2: SALES in Value',
                        'data': table2_df.to_dict('records'),
                        'columns': table2_df.columns.tolist(),
                        'shape': [int(table2_df.shape[0]), int(table2_df.shape[1])],
                        'analysis_type': analysis_info,
                        'column_info': {
                            'act_columns': act_columns2,
                            'gr_columns': gr_columns2,
                            'ach_columns': ach_columns2,
                            'budget_columns': budget_columns2,
                            'ly_columns': ly_columns2,
                            'total_columns': int(len(table2_df.columns))
                        },
                        'original_headers': [str(h) for h in original_headers2],
                        'processing_info': {
                            'original_column_count': len(original_headers2),
                            'final_column_count': len(table2_df.columns),
                            'processing_successful': True
                        }
                    }
                    
                    # Store in session
                    session[table2_key] = table2_data
            
            # Return the requested table
            if table_choice == "Table 1: SALES in MT/Tonage" and table1_data:
                return jsonify(table1_data)
            elif table_choice == "Table 2: SALES in Value" and table2_data:
                return jsonify(table2_data)
            else:
                # Fallback: return any available table
                if table2_data:
                    return jsonify(table2_data)
                elif table1_data:
                    return jsonify(table1_data)
                else:
                    return jsonify({
                        'error': 'No valid tables found in the sheet',
                        'debug_info': {
                            'tried_headers_table1': table1_possible_headers,
                            'tried_headers_table2': table2_possible_headers,
                            'sheet_preview': df_sheet.head(10).to_dict('records')
                        }
                    }), 400
            
        except Exception as e:
            return jsonify({
                'error': f'Error extracting tables: {str(e)}',
                'traceback': traceback.format_exc()
            }), 500
        
    except Exception as e:
        return jsonify({
            'error': f'Error processing auditor file: {str(e)}',
            'traceback': traceback.format_exc()
        }), 500

@auditor_bp.route('/get-available-tables', methods=['POST'])
def get_available_tables():
    """Get list of available tables in the auditor sheet and check session cache"""
    try:
        data = request.json
        filepath = data.get('filepath')
        sheet_name = data.get('sheet_name')
        
        if not filepath or not sheet_name:
            return jsonify({'error': 'Missing filepath or sheet_name'}), 400
        
        # Check session first
        table1_key = get_session_key(filepath, sheet_name, 'Table1')
        table2_key = get_session_key(filepath, sheet_name, 'Table2')
        
        available_tables = []
        table_info = {}
        
        # Check if tables are in session
        if table1_key in session:
            available_tables.append("Table 1: SALES in MT/Tonage")
            table_info["Table 1: SALES in MT/Tonage"] = {
                'source': 'session_cache',
                'estimated_rows': session[table1_key]['shape'][0]
            }
        
        if table2_key in session:
            available_tables.append("Table 2: SALES in Value")
            table_info["Table 2: SALES in Value"] = {
                'source': 'session_cache',
                'estimated_rows': session[table2_key]['shape'][0]
            }
        
        # If both tables are in session, return cached info
        if len(available_tables) == 2:
            return jsonify({
                'success': True,
                'available_tables': available_tables,
                'table_info': table_info,
                'analysis_type': session[table2_key]['analysis_type'] if table2_key in session else session[table1_key]['analysis_type'],
                'default_table': "Table 2: SALES in Value" if "Table 2: SALES in Value" in available_tables else available_tables[0],
                'from_cache': True
            })
        
        # If not in session, process the file to detect tables
        # Initialize DataProcessor
        processor = DataProcessor()
        
        # Read the Excel sheet
        df_sheet = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
        
        # Detect analysis type
        analysis_info = detect_analysis_type(sheet_name)
        
        # Define possible headers
        table1_possible_headers = [
            "SALES in MT", "SALES IN MT", "Sales in MT", "SALES IN TONNAGE", "SALES IN TON",
            "Tonnage", "TONNAGE", "Tonnage Sales", "Sales Tonnage", "Metric Tons", "MT Sales"
        ]
        
        table2_possible_headers = [
            "SALES in Value", "SALES IN VALUE", "Sales in Value", "SALES IN RS", "VALUE SALES",
            "Value", "VALUE", "Sales Value"
        ]
        
        available_tables = []
        table_info = {}
        
        # Check for Table 1 (MT/Tonnage) using DataProcessor
        try:
            idx1, data_start1 = processor.extract_tables(
                df_sheet, 
                table1_possible_headers, 
                is_product_analysis=analysis_info['is_product_analysis']
            )
            if idx1 is not None:
                available_tables.append("Table 1: SALES in MT/Tonage")
                table_info["Table 1: SALES in MT/Tonage"] = {
                    'header_row': idx1,
                    'data_start': data_start1,
                    'estimated_rows': len(df_sheet) - data_start1,
                    'source': 'file_detection'
                }
        except Exception as e:
            print(f"Could not find Table 1: {e}")
        
        # Check for Table 2 (Value) using DataProcessor
        try:
            idx2, data_start2 = processor.extract_tables(
                df_sheet, 
                table2_possible_headers, 
                is_product_analysis=analysis_info['is_product_analysis']
            )
            if idx2 is not None:
                available_tables.append("Table 2: SALES in Value")
                table_info["Table 2: SALES in Value"] = {
                    'header_row': idx2,
                    'data_start': data_start2,
                    'estimated_rows': len(df_sheet) - data_start2,
                    'source': 'file_detection'
                }
        except Exception as e:
            print(f"Could not find Table 2: {e}")
        
        return jsonify({
            'success': True,
            'available_tables': available_tables,
            'table_info': table_info,
            'analysis_type': analysis_info,
            'default_table': "Table 2: SALES in Value" if "Table 2: SALES in Value" in available_tables else (available_tables[0] if available_tables else None),
            'sheet_info': {
                'total_rows': len(df_sheet),
                'total_columns': len(df_sheet.columns)
            },
            'from_cache': False
        })
        
    except Exception as e:
        return jsonify({
            'error': f'Error getting available tables: {str(e)}',
            'traceback': traceback.format_exc()
        }), 500

@auditor_bp.route('/get-cached-table', methods=['POST'])
def get_cached_table():
    """Get a specific table from session cache"""
    try:
        data = request.json
        filepath = data.get('filepath')
        sheet_name = data.get('sheet_name')
        table_choice = data.get('table_choice')
        
        if not all([filepath, sheet_name, table_choice]):
            return jsonify({'error': 'Missing required parameters'}), 400
        
        # Determine which table to get
        if table_choice == "Table 1: SALES in MT/Tonage":
            table_key = get_session_key(filepath, sheet_name, 'Table1')
        elif table_choice == "Table 2: SALES in Value":
            table_key = get_session_key(filepath, sheet_name, 'Table2')
        else:
            return jsonify({'error': 'Invalid table choice'}), 400
        
        # Check if table exists in session
        if table_key in session:
            return jsonify(session[table_key])
        else:
            return jsonify({'error': 'Table not found in cache. Please refresh.'}), 404
            
    except Exception as e:
        return jsonify({
            'error': f'Error getting cached table: {str(e)}',
            'traceback': traceback.format_exc()
        }), 500

@auditor_bp.route('/clear-session-cache', methods=['POST'])
def clear_session_cache():
    """Clear session cache for specific file/sheet or all"""
    try:
        data = request.json
        filepath = data.get('filepath')
        sheet_name = data.get('sheet_name')
        
        if filepath and sheet_name:
            # Clear specific file/sheet cache
            table1_key = get_session_key(filepath, sheet_name, 'Table1')
            table2_key = get_session_key(filepath, sheet_name, 'Table2')
            
            session.pop(table1_key, None)
            session.pop(table2_key, None)
            
            return jsonify({
                'success': True,
                'message': f'Cache cleared for {sheet_name}'
            })
        else:
            # Clear all auditor cache
            keys_to_remove = [key for key in session.keys() if key.startswith('auditor_')]
            for key in keys_to_remove:
                session.pop(key, None)
            
            return jsonify({
                'success': True,
                'message': f'Cleared {len(keys_to_remove)} cached tables'
            })
            
    except Exception as e:
        return jsonify({
            'error': f'Error clearing cache: {str(e)}'
        }), 500

@auditor_bp.route('/export-auditor-data', methods=['POST'])
def export_auditor_data():
    """Export processed auditor data from session cache"""
    try:
        data = request.json
        filepath = data.get('filepath')
        sheet_name = data.get('sheet_name')
        table_choice = data.get('table_choice', 'Table 2: SALES in Value')
        export_format = data.get('format', 'csv')
        
        if not filepath or not sheet_name:
            return jsonify({'error': 'Missing filepath or sheet_name'}), 400
        
        # Try to get data from session cache first
        if table_choice == "Table 1: SALES in MT/Tonage":
            table_key = get_session_key(filepath, sheet_name, 'Table1')
        else:
            table_key = get_session_key(filepath, sheet_name, 'Table2')
        
        if table_key in session:
            # Get data from session
            cached_data = session[table_key]
            df = pd.DataFrame(cached_data['data'])
        else:
            # If not in cache, process fresh (fallback)
            process_response = process_auditor_auto()
            if isinstance(process_response, tuple):
                process_data = process_response[0].get_json()
            else:
                process_data = process_response.get_json()
            
            if not process_data.get('success'):
                return jsonify(process_data), 400
            
            df = pd.DataFrame(process_data['data'])
        
        # Generate filename
        clean_table_name = table_choice.replace(':', '').replace(' ', '_').lower()
        filename = f"auditor_{clean_table_name}_{sheet_name}_{export_format}"
        
        if export_format == 'csv':
            from io import StringIO
            output = StringIO()
            df.to_csv(output, index=False)
            csv_data = output.getvalue()
            
            return jsonify({
                'success': True,
                'data': csv_data,
                'filename': f"{filename}.csv",
                'format': 'csv'
            })
        else:
            return jsonify({'error': 'Only CSV export is currently supported'}), 400
                
    except Exception as e:
        return jsonify({
            'error': f'Error exporting auditor data: {str(e)}',
            'traceback': traceback.format_exc()
        }), 500