import pandas as pd
import numpy as np
import os
from pathlib import Path
import logging

logger = logging.getLogger(__name__)

def extract_executive_name(executive):
    """Extract and standardize executive name"""
    if pd.isna(executive) or str(executive).strip() == '':
        return 'BLANK'
    return str(executive).strip().upper()

def calculate_executive_budget_vs_billed(
    sales_file_path, budget_file_path, 
    sales_date, sales_value, sales_quantity, sales_product_group, sales_sl_code, sales_executive, sales_area,
    budget_value, budget_quantity, budget_product_group, budget_sl_code, budget_executive, budget_area,
    selected_executives, selected_months, selected_branches=None
):
    """
    Calculate executive budget vs billed analysis following the exact Streamlit logic
    """
    try:
        print("Starting executive budget vs billed calculation...")
        
        # Load data files
        sales_df = load_data(sales_file_path)
        budget_df = load_data(budget_file_path)
        
        print(f"Sales data loaded: {len(sales_df)} rows")
        print(f"Budget data loaded: {len(budget_df)} rows")
        
        # Create copies to avoid modifying original DataFrames
        sales_df = sales_df.copy()
        budget_df = budget_df.copy()

        # Validate column existence
        required_sales_cols = [sales_date, sales_value, sales_quantity, sales_executive,
                              sales_product_group, sales_sl_code, sales_area]
        required_budget_cols = [budget_value, budget_quantity, budget_executive,
                               budget_product_group, budget_sl_code, budget_area]
        
        missing_sales_cols = [col for col in required_sales_cols if col and col not in sales_df.columns]
        missing_budget_cols = [col for col in required_budget_cols if col and col not in budget_df.columns]
        
        if missing_sales_cols:
            return {"error": f"Missing columns in sales data: {missing_sales_cols}"}
        if missing_budget_cols:
            return {"error": f"Missing columns in budget data: {missing_budget_cols}"}

        # Convert date and numeric columns
        sales_df[sales_date] = pd.to_datetime(sales_df[sales_date], dayfirst=True, errors='coerce')
        sales_df[sales_value] = pd.to_numeric(sales_df[sales_value], errors='coerce').fillna(0)
        sales_df[sales_quantity] = pd.to_numeric(sales_df[sales_quantity], errors='coerce').fillna(0)
        budget_df[budget_value] = pd.to_numeric(budget_df[budget_value], errors='coerce').fillna(0)
        budget_df[budget_quantity] = pd.to_numeric(budget_df[budget_quantity], errors='coerce').fillna(0)

        # Filter sales data for the selected months
        if selected_months:
            filtered_sales_df = sales_df[sales_df[sales_date].dt.strftime('%b %y').isin(selected_months)].copy()
            if filtered_sales_df.empty:
                return {"error": f"No sales data found for selected months: {selected_months}"}
        else:
            filtered_sales_df = sales_df.copy()

        # Standardize branch and executive names
        filtered_sales_df[sales_area] = filtered_sales_df[sales_area].astype(str).str.strip().str.upper()
        filtered_sales_df[sales_executive] = filtered_sales_df[sales_executive].astype(str).str.strip().str.upper()
        budget_df[budget_area] = budget_df[budget_area].astype(str).str.strip().str.upper()
        budget_df[budget_executive] = budget_df[budget_executive].astype(str).str.strip().str.upper()
        
        # Apply branch filter if provided
        if selected_branches:
            filtered_sales_df = filtered_sales_df[filtered_sales_df[sales_area].isin([b.upper() for b in selected_branches])]
            budget_df = budget_df[budget_df[budget_area].isin([b.upper() for b in selected_branches])]
            if filtered_sales_df.empty or budget_df.empty:
                return {"error": f"No data found for selected branches: {', '.join(selected_branches)}"}

        # Determine executives to display
        if selected_branches:
            branch_sales_df = filtered_sales_df[filtered_sales_df[sales_area].isin([b.upper() for b in selected_branches])]
            branch_budget_df = budget_df[budget_df[budget_area].isin([b.upper() for b in selected_branches])]
            branch_executives = sorted(set(branch_sales_df[sales_executive].dropna().unique()) | 
                                     set(branch_budget_df[budget_executive].dropna().unique()))
            
            if selected_executives:
                selected_execs_upper = [str(exec).strip().upper() for exec in selected_executives]
                executives_to_display = [exec for exec in branch_executives if exec in selected_execs_upper]
            else:
                executives_to_display = branch_executives
        else:
            executives_to_display = [str(exec).strip().upper() for exec in selected_executives] if selected_executives else \
                                    sorted(set(filtered_sales_df[sales_executive].dropna().unique()) | 
                                          set(budget_df[budget_executive].dropna().unique()))

        # Filter by selected executives
        filtered_sales_df = filtered_sales_df[filtered_sales_df[sales_executive].isin(executives_to_display)].copy()
        budget_filtered = budget_df[budget_df[budget_executive].isin(executives_to_display)].copy()
        
        if filtered_sales_df.empty or budget_filtered.empty:
            return {"error": "No data found for selected executives."}

        # Standardize SL codes and product groups
        filtered_sales_df[sales_sl_code] = filtered_sales_df[sales_sl_code].astype(str).str.strip().str.upper()
        filtered_sales_df[sales_product_group] = filtered_sales_df[sales_product_group].astype(str).str.strip().str.upper()
        budget_filtered[budget_sl_code] = budget_filtered[budget_sl_code].astype(str).str.strip().str.upper()
        budget_filtered[budget_product_group] = budget_filtered[budget_product_group].astype(str).str.strip().str.upper()

        # Process Budget Data
        budget_grouped = budget_filtered.groupby([
            budget_executive, 
            budget_sl_code, 
            budget_product_group
        ]).agg({
            budget_quantity: 'sum',
            budget_value: 'sum'
        }).reset_index()
        
        budget_valid = budget_grouped[
            (budget_grouped[budget_quantity] > 0) & 
            (budget_grouped[budget_value] > 0)
        ].copy()
        
        if budget_valid.empty:
            return {"error": "No valid budget data found (with qty > 0 and value > 0)."}
        
        # Process Sales Data
        final_results = []
        
        for _, budget_row in budget_valid.iterrows():
            executive = budget_row[budget_executive]
            sl_code = budget_row[budget_sl_code]
            product = budget_row[budget_product_group]
            budget_qty = budget_row[budget_quantity]
            budget_val = budget_row[budget_value]
            
            matching_sales = filtered_sales_df[
                (filtered_sales_df[sales_executive] == executive) &
                (filtered_sales_df[sales_sl_code] == sl_code) &
                (filtered_sales_df[sales_product_group] == product)
            ]
            
            if not matching_sales.empty:
                sales_qty_total = matching_sales[sales_quantity].sum()
                sales_value_total = matching_sales[sales_value].sum()
            else:
                sales_qty_total = 0
                sales_value_total = 0
            
            final_qty = budget_qty if sales_qty_total > budget_qty else sales_qty_total
            final_value = budget_val if sales_value_total > budget_val else sales_value_total
            
            final_results.append({
                'Executive': executive,
                'SL_Code': sl_code,
                'Product': product,
                'Budget_Qty': budget_qty,
                'Sales_Qty': sales_qty_total,
                'Final_Qty': final_qty,
                'Budget_Value': budget_val,
                'Sales_Value': sales_value_total,
                'Final_Value': final_value
            })
        
        results_df = pd.DataFrame(final_results)
        
        # Aggregate by Executive
        exec_qty_summary = results_df.groupby('Executive').agg({
            'Budget_Qty': 'sum',
            'Final_Qty': 'sum'
        }).reset_index()
        exec_qty_summary.columns = ['Executive', 'Budget Qty', 'Billed Qty']
        
        exec_value_summary = results_df.groupby('Executive').agg({
            'Budget_Value': 'sum',
            'Final_Value': 'sum'
        }).reset_index()
        exec_value_summary.columns = ['Executive', 'Budget Value', 'Billed Value']
        
        # NUCLEAR FIX: Completely rebuild DataFrames from scratch
        # Build QUANTITY DataFrame from scratch
        qty_data = []
        for exec_name in executives_to_display:
            # Find matching data in summary
            exec_qty_row = exec_qty_summary[exec_qty_summary['Executive'] == exec_name]
            
            if not exec_qty_row.empty:
                budget_val = round(float(exec_qty_row['Budget Qty'].iloc[0]), 2)
                billed_val = round(float(exec_qty_row['Billed Qty'].iloc[0]), 2)
            else:
                budget_val = 0.0
                billed_val = 0.0
            
            # Calculate percentage from scratch
            if budget_val > 0:
                percentage = round((billed_val / budget_val) * 100, 2)
            else:
                percentage = 0.0
            
            qty_data.append({
                'Executive': exec_name,
                'Budget Qty': budget_val,
                'Billed Qty': billed_val,
                '%': percentage
            })
        
        # Create fresh DataFrame
        budget_vs_billed_qty_df = pd.DataFrame(qty_data)
        
        # Build VALUE DataFrame from scratch
        value_data = []
        for exec_name in executives_to_display:
            # Find matching data in summary
            exec_value_row = exec_value_summary[exec_value_summary['Executive'] == exec_name]
            
            if not exec_value_row.empty:
                budget_val = round(float(exec_value_row['Budget Value'].iloc[0]), 2)
                billed_val = round(float(exec_value_row['Billed Value'].iloc[0]), 2)
            else:
                budget_val = 0.0
                billed_val = 0.0
            
            # Calculate percentage from scratch
            if budget_val > 0:
                percentage = round((billed_val / budget_val) * 100, 2)
            else:
                percentage = 0.0
            
            value_data.append({
                'Executive': exec_name,
                'Budget Value': budget_val,
                'Billed Value': billed_val,
                '%': percentage
            })
        
        # Create fresh DataFrame
        budget_vs_billed_value_df = pd.DataFrame(value_data)
        
        # ADDITIONAL SAFETY CHECK: Force fix any remaining anomalies
        # Check for any impossible percentages (non-zero % with zero budget and billed)
        anomaly_mask_qty = (
            (budget_vs_billed_qty_df['Budget Qty'] == 0) & 
            (budget_vs_billed_qty_df['Billed Qty'] == 0) & 
            (budget_vs_billed_qty_df['%'] != 0.0)
        )
        if anomaly_mask_qty.any():
            budget_vs_billed_qty_df.loc[anomaly_mask_qty, '%'] = 0.0
            print(f"🔧 SAFETY FIX: Reset {anomaly_mask_qty.sum()} anomalous quantity percentages to 0.0%")
        
        anomaly_mask_value = (
            (budget_vs_billed_value_df['Budget Value'] == 0) & 
            (budget_vs_billed_value_df['Billed Value'] == 0) & 
            (budget_vs_billed_value_df['%'] != 0.0)
        )
        if anomaly_mask_value.any():
            budget_vs_billed_value_df.loc[anomaly_mask_value, '%'] = 0.0
            print(f"🔧 SAFETY FIX: Reset {anomaly_mask_value.sum()} anomalous value percentages to 0.0%")
        
        # Create Overall Sales DataFrames
        overall_sales_data = filtered_sales_df.groupby(sales_executive).agg({
            sales_quantity: 'sum',
            sales_value: 'sum'
        }).reset_index()
        overall_sales_data.columns = ['Executive', 'Overall_Sales_Qty', 'Overall_Sales_Value']
        
        budget_totals = results_df.groupby('Executive').agg({
            'Budget_Qty': 'sum',
            'Budget_Value': 'sum'
        }).reset_index()
        
        overall_sales_qty_df = pd.DataFrame({'Executive': executives_to_display})
        overall_sales_qty_df = pd.merge(
            overall_sales_qty_df,
            budget_totals[['Executive', 'Budget_Qty']].rename(columns={'Budget_Qty': 'Budget Qty'}),
            on='Executive',
            how='left'
        ).fillna({'Budget Qty': 0})
        
        overall_sales_qty_df = pd.merge(
            overall_sales_qty_df,
            overall_sales_data[['Executive', 'Overall_Sales_Qty']].rename(columns={'Overall_Sales_Qty': 'Billed Qty'}),
            on='Executive',
            how='left'
        ).fillna({'Billed Qty': 0})
        
        overall_sales_value_df = pd.DataFrame({'Executive': executives_to_display})
        overall_sales_value_df = pd.merge(
            overall_sales_value_df,
            budget_totals[['Executive', 'Budget_Value']].rename(columns={'Budget_Value': 'Budget Value'}),
            on='Executive',
            how='left'
        ).fillna({'Budget Value': 0})
        
        overall_sales_value_df = pd.merge(
            overall_sales_value_df,
            overall_sales_data[['Executive', 'Overall_Sales_Value']].rename(columns={'Overall_Sales_Value': 'Billed Value'}),
            on='Executive',
            how='left'
        ).fillna({'Billed Value': 0})
        
        # Add Total Rows
        total_budget_qty = round(budget_vs_billed_qty_df['Budget Qty'].sum(), 2)
        total_billed_qty = round(budget_vs_billed_qty_df['Billed Qty'].sum(), 2)
        total_percentage_qty = round((total_billed_qty / total_budget_qty * 100), 2) if total_budget_qty > 0 else 0.0
        
        total_row_qty = pd.DataFrame({
            'Executive': ['TOTAL'],
            'Budget Qty': [total_budget_qty],
            'Billed Qty': [total_billed_qty],
            '%': [total_percentage_qty]
        })
        budget_vs_billed_qty_df = pd.concat([budget_vs_billed_qty_df, total_row_qty], ignore_index=True)
        
        total_budget_value = round(budget_vs_billed_value_df['Budget Value'].sum(), 2)
        total_billed_value = round(budget_vs_billed_value_df['Billed Value'].sum(), 2)
        total_percentage_value = round((total_billed_value / total_budget_value * 100), 2) if total_budget_value > 0 else 0.0
        
        total_row_value = pd.DataFrame({
            'Executive': ['TOTAL'],
            'Budget Value': [total_budget_value],
            'Billed Value': [total_billed_value],
            '%': [total_percentage_value]
        })
        budget_vs_billed_value_df = pd.concat([budget_vs_billed_value_df, total_row_value], ignore_index=True)
        
        total_row_overall_qty = pd.DataFrame({
            'Executive': ['TOTAL'],
            'Budget Qty': [round(overall_sales_qty_df['Budget Qty'].sum(), 2)],
            'Billed Qty': [round(overall_sales_qty_df['Billed Qty'].sum(), 2)]
        })
        overall_sales_qty_df = pd.concat([overall_sales_qty_df, total_row_overall_qty], ignore_index=True)
        
        total_row_overall_value = pd.DataFrame({
            'Executive': ['TOTAL'],
            'Budget Value': [round(overall_sales_value_df['Budget Value'].sum(), 2)],
            'Billed Value': [round(overall_sales_value_df['Billed Value'].sum(), 2)]
        })
        overall_sales_value_df = pd.concat([overall_sales_value_df, total_row_overall_value], ignore_index=True)
        
        # Round all numeric columns
        for df in [budget_vs_billed_qty_df, budget_vs_billed_value_df, overall_sales_qty_df, overall_sales_value_df]:
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            df[numeric_cols] = df[numeric_cols].round(2)
        
        # FINAL NUCLEAR SAFETY CHECK - Force any remaining issues
        for df_name, df in [("QTY", budget_vs_billed_qty_df), ("VALUE", budget_vs_billed_value_df)]:
            if df_name == "QTY":
                budget_col, billed_col = 'Budget Qty', 'Billed Qty'
            else:
                budget_col, billed_col = 'Budget Value', 'Billed Value'
            
            # Check for any remaining anomalies
            final_anomaly_mask = (
                (df[budget_col] == 0.0) & 
                (df[billed_col] == 0.0) & 
                (df['%'] != 0.0)
            )
            
            if final_anomaly_mask.any():
                print(f"🚨 FINAL CHECK: Found {final_anomaly_mask.sum()} anomalies in {df_name} DataFrame")
                print(f"Anomalous executives: {df[final_anomaly_mask]['Executive'].tolist()}")
                df.loc[final_anomaly_mask, '%'] = 0.0
                print(f"✅ Fixed all anomalies in {df_name} DataFrame")
        
        print("Executive budget vs billed calculation completed successfully")
        
        return {
            'success': True,
            'budget_vs_billed_qty': budget_vs_billed_qty_df.to_dict('records'),
            'budget_vs_billed_value': budget_vs_billed_value_df.to_dict('records'),
            'overall_sales_qty': overall_sales_qty_df.to_dict('records'),
            'overall_sales_value': overall_sales_value_df.to_dict('records')
        }
        
    except Exception as e:
        print(f"Error in calculate_executive_budget_vs_billed: {str(e)}")
        import traceback
        traceback.print_exc()
        return {
            'success': False,
            'error': f"Error calculating executive budget vs billed: {str(e)}"
        }

def load_data(file_path):
    """Load data from CSV or Excel file"""
    try:
        # Convert to absolute path
        if not os.path.isabs(file_path):
            file_path = os.path.abspath(file_path)
        
        if file_path.endswith('.csv'):
            return pd.read_csv(file_path)
        elif file_path.endswith(('.xlsx', '.xls')):
            return pd.read_excel(file_path)
        else:
            raise ValueError(f"Unsupported file format: {file_path}")
    except Exception as e:
        print(f"Error loading file {file_path}: {str(e)}")
        raise

def auto_map_executive_columns(sales_df, budget_df):
    """
    Auto-map columns for executive analysis following the Streamlit logic
    """
    try:
        def find_column(columns, target_names, default_index=0):
            for target in target_names:
                for col in columns:
                    if col.lower() == target.lower():
                        return col
            return columns[default_index] if columns else None

        # Enhanced column mappings - Fixed quantity mapping
        column_mappings = {
            'sales_date': ['Date'],
            'sales_value': ['Value', 'Invoice Value'],
            'sales_qty': ['Actual Quantity', 'Quantity', 'Qty', 'Sales Quantity', 'Sales Qty'],
            'sales_product_group': ['Type (Make)', 'Product Group'],
            'sales_sl_code': ['Customer Code', 'SL Code'],
            'sales_area': ['Branch', 'Area'],
            'sales_exec': ['Executive Name', 'Executive'],
            'budget_value': ['Budget Value', 'Value'],
            'budget_qty': ['Budget Quantity', 'Quantity', 'Qty', 'Budget Qty'],
            'budget_product_group': ['Product Group', 'Type (Make)'],
            'budget_sl_code': ['SL Code', 'Customer Code'],
            'budget_area': ['Branch', 'Area'],
            'budget_exec': ['Executive Name', 'Executive']
        }

        sales_mapping = {}
        budget_mapping = {}

        # Auto-map sales columns
        for key, targets in column_mappings.items():
            if key.startswith('sales_'):
                clean_key = key.replace('sales_', '')
                sales_mapping[clean_key] = find_column(sales_df.columns.tolist(), targets)

        # Auto-map budget columns
        for key, targets in column_mappings.items():
            if key.startswith('budget_'):
                clean_key = key.replace('budget_', '')
                budget_mapping[clean_key] = find_column(budget_df.columns.tolist(), targets)

        return {
            'sales_mapping': sales_mapping,
            'budget_mapping': budget_mapping,
            'sales_columns': list(sales_df.columns),
            'budget_columns': list(budget_df.columns)
        }
        
    except Exception as e:
        print(f"Error in auto_map_executive_columns: {str(e)}")
        return {
            'sales_mapping': {},
            'budget_mapping': {},
            'sales_columns': [],
            'budget_columns': []
        }

def get_executives_and_branches(sales_df, budget_df, sales_exec_col, budget_exec_col, sales_area_col, budget_area_col):
    """Get available executives and branches from the data"""
    try:
        # Get executives
        sales_execs = []
        budget_execs = []
        
        if sales_exec_col and sales_exec_col in sales_df.columns:
            sales_execs = sales_df[sales_exec_col].dropna().astype(str).str.strip().str.upper().unique().tolist()
        
        if budget_exec_col and budget_exec_col in budget_df.columns:
            budget_execs = budget_df[budget_exec_col].dropna().astype(str).str.strip().str.upper().unique().tolist()
        
        all_executives = sorted(set(sales_execs + budget_execs))
        
        # Get branches
        sales_branches = []
        budget_branches = []
        
        if sales_area_col and sales_area_col in sales_df.columns:
            sales_branches = sales_df[sales_area_col].dropna().astype(str).str.strip().str.upper().unique().tolist()
        
        if budget_area_col and budget_area_col in budget_df.columns:
            budget_branches = budget_df[budget_area_col].dropna().astype(str).str.strip().str.upper().unique().tolist()
        
        all_branches = sorted(set(sales_branches + budget_branches))
        
        return {
            'executives': all_executives,
            'sales_executives': sales_execs,
            'budget_executives': budget_execs,
            'branches': all_branches,
            'sales_branches': sales_branches,
            'budget_branches': budget_branches
        }
        
    except Exception as e:
        print(f"Error getting executives and branches: {str(e)}")
        return {
            'executives': [],
            'sales_executives': [],
            'budget_executives': [],
            'branches': [],
            'sales_branches': [],
            'budget_branches': []
        }

def get_available_months(sales_df, sales_date_col):
    """Get available months from sales data"""
    try:
        if not sales_date_col or sales_date_col not in sales_df.columns:
            return []
        
        sales_df[sales_date_col] = pd.to_datetime(sales_df[sales_date_col], dayfirst=True, errors='coerce')
        available_months = sorted(sales_df[sales_date_col].dt.strftime('%b %y').dropna().unique().tolist())
        
        return available_months
        
    except Exception as e:
        print(f"Error getting available months: {str(e)}")
        return []