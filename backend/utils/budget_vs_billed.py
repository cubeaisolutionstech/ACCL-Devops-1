import pandas as pd
import numpy as np
import os
from datetime import datetime
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Mappings from branch.py
branch_mapping = {
    'PONDY': 'PUDUCHERRY', 'PDY': 'PUDUCHERRY', 'Puducherry - PDY': 'PUDUCHERRY',
    'COVAI': 'COIMBATORE', 'CBE': 'COIMBATORE', 'Coimbatore - CBE': 'COIMBATORE',
    'ERD': 'ERODE', 'Erode - ERD': 'ERODE', 'ERD002': 'ERODE', 'ERD001': 'ERODE',
    'ERDTD1': 'ERODE', 'ERD003': 'ERODE', 'ERD004': 'ERODE', 'ERD005': 'ERODE',
    'ERD007': 'ERODE',
    'KRR': 'KARUR',
    'Chennai - CHN': 'CHENNAI', 'CHN': 'CHENNAI',
    'Tirupur - TPR': 'TIRUPUR', 'TPR': 'TIRUPUR',
    'Madurai - MDU': 'MADURAI', 'MDU': 'MADURAI',
    'POULTRY': 'POULTRY', 'Poultry - PLT': 'POULTRY',
    'SALEM': 'SALEM', 'Salem - SLM': 'SALEM',
    'HO': 'HO',
    'SLM002': 'SALEM', 'SLMTD1': 'SALEM',
    'BHV1': 'BHAVANI',
    'CBU': 'COIMBATORE',
    'VLR': 'VELLORE',
    'TRZ': 'TRICHY',
    'TVL': 'TIRUNELVELI',
    'NGS': 'NAGERCOIL',
    'PONDICHERRY': 'PUDUCHERRY',
    'BLR': 'BANGALORE', 'BANGALORE': 'BANGALORE', 'BGLR': 'BANGALORE'
}

# -------------------- UTILITY MAPPING FUNCTIONS --------------------

def map_branch(branch_name):
    if pd.isna(branch_name):
        return 'Unknown'
    branch_str = str(branch_name).strip().upper()
    if ' - ' in branch_str:
        branch_str = branch_str.split(' - ')[-1].strip()
    return branch_mapping.get(branch_str, branch_str)

def find_column_by_names(columns, target_names):
    columns_upper = [col.upper() for col in columns]
    target_names_upper = [name.upper() for name in target_names]
    for target in target_names_upper:
        if target in columns_upper:
            return columns[columns_upper.index(target)]
    return None

def auto_map_budget_columns(sales_columns, budget_columns):
    sales_mapping = {
        'date': find_column_by_names(sales_columns, ['Date', 'Bill Date', 'Invoice Date']),
        'value': find_column_by_names(sales_columns, ['Value', 'Invoice Value', 'Amount']),
        'product_group': find_column_by_names(sales_columns, ['Type (Make)', 'Product Group', 'Type', 'Make']),
        'area': find_column_by_names(sales_columns, ['Branch', 'Area', 'Location']),
        'quantity': find_column_by_names(sales_columns, ['Actual Quantity', 'Quantity', 'Qty']),
        'sl_code': find_column_by_names(sales_columns, ['Customer Code', 'SL Code', 'Customer ID']),
        'executive': find_column_by_names(sales_columns, ['Executive', 'Sales Executive', 'Executive Name'])
    }
    budget_mapping = {
        'area': find_column_by_names(budget_columns, ['Branch', 'Area', 'Location']),
        'quantity': find_column_by_names(budget_columns, ["Qty – Apr'25", 'Quantity', 'Qty', 'Budget Qty']),
        'sl_code': find_column_by_names(budget_columns, ['SL Code', 'Customer Code', 'Customer ID']),
        'value': find_column_by_names(budget_columns, ["Value – Apr'25", 'Value', 'Budget Value', 'Amount']),
        'product_group': find_column_by_names(budget_columns, ['Product Group', 'Type(Make)', 'Product', 'Type']),
        'executive': find_column_by_names(budget_columns, ['Executive Name', 'Executive', 'Sales Executive'])
    }
    return sales_mapping, budget_mapping

# -------------------- CORE BUDGET VS BILLED CALCULATION --------------------

def clean_and_format_df(df, percent_column='%', column_order=None):
    df = df.copy()

    # Round numeric columns
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            df[col] = df[col].round(0).astype('Int64')

    # Ensure percent column moves to the end if no order is given
    if column_order:
        ordered_cols = [col for col in column_order if col in df.columns]
        remaining_cols = [col for col in df.columns if col not in ordered_cols]
        df = df[ordered_cols + remaining_cols]
    elif percent_column in df.columns:
        # Default behavior: move % to end
        percent_series = df.pop(percent_column)
        df[percent_column] = percent_series

    return df



def calculate_budget_vs_billed(data):
    import os
    import pandas as pd

    # Load files
    sales_path = os.path.join('uploads', data['sales_filename'])
    budget_path = os.path.join('uploads', data['budget_filename'])

    sales_df = pd.read_excel(sales_path, sheet_name=data['sales_sheet'], header=data['sales_header'] - 1, dtype=str)
    budget_df = pd.read_excel(budget_path, sheet_name=data['budget_sheet'], header=data['budget_header'] - 1, dtype=str)

    # Convert and clean numeric and date columns
    sales_df[data['sales_date_col']] = pd.to_datetime(sales_df[data['sales_date_col']], dayfirst=True, errors='coerce')
    sales_df[data['sales_value_col']] = pd.to_numeric(sales_df[data['sales_value_col']], errors='coerce').fillna(0)
    sales_df[data['sales_qty_col']] = pd.to_numeric(sales_df[data['sales_qty_col']], errors='coerce').fillna(0)

    budget_df[data['budget_value_col']] = pd.to_numeric(budget_df[data['budget_value_col']], errors='coerce').fillna(0)
    budget_df[data['budget_qty_col']] = pd.to_numeric(budget_df[data['budget_qty_col']], errors='coerce').fillna(0)

    # Filters
    selected_month = data['selected_month']
    selected_sales_execs = data['selected_sales_execs']
    selected_budget_execs = data['selected_budget_execs']
    selected_branches = data['selected_branches']

    sales_df = sales_df[sales_df[data['sales_exec_col']].isin(selected_sales_execs)]
    budget_df = budget_df[budget_df[data['budget_exec_col']].isin(selected_budget_execs)]

    sales_df = sales_df[sales_df[data['sales_date_col']].dt.strftime('%b %y') == selected_month]

    # Map branches
    sales_df[data['sales_area_col']] = sales_df[data['sales_area_col']].map(map_branch)
    budget_df[data['budget_area_col']] = budget_df[data['budget_area_col']].map(map_branch)

    # Filter branches
    sales_df = sales_df[sales_df[data['sales_area_col']].isin(selected_branches)]
    budget_df = budget_df[budget_df[data['budget_area_col']].isin(selected_branches)]

    # ---------------- Budget Matching Logic ----------------
    results = []
    budget_grouped = budget_df.groupby([
        data['budget_area_col'],
        data['budget_sl_code_col'],
        data['budget_product_group_col']
    ]).agg({
        data['budget_qty_col']: 'sum',
        data['budget_value_col']: 'sum'
    }).reset_index()

    valid_budget = budget_grouped[
        (budget_grouped[data['budget_qty_col']] > 0) &
        (budget_grouped[data['budget_value_col']] > 0)
    ]

    for _, row in valid_budget.iterrows():
        branch = row[data['budget_area_col']]
        sl_code = row[data['budget_sl_code_col']]
        product = row[data['budget_product_group_col']]
        budget_qty = row[data['budget_qty_col']]
        budget_val = row[data['budget_value_col']]

        sales_match = sales_df[
            (sales_df[data['sales_area_col']] == branch) &
            (sales_df[data['sales_sl_code_col']] == sl_code) &
            (sales_df[data['sales_product_group_col']] == product)
        ]

        sales_qty = sales_match[data['sales_qty_col']].sum()
        sales_val = sales_match[data['sales_value_col']].sum()

        results.append({
            'Branch': branch,
            'SL_Code': sl_code,
            'Product': product,
            'Budget_Qty': budget_qty,
            'Sales_Qty': sales_qty,
            'Final_Qty': budget_qty if sales_qty > budget_qty else sales_qty,
            'Budget_Value': budget_val,
            'Sales_Value': sales_val,
            'Final_Value': budget_val if sales_val > budget_val else sales_val
        })

    df = pd.DataFrame(results)

    # ---- Quantity Summary
    qty_summary = df.groupby('Branch')[['Budget_Qty', 'Final_Qty']].sum().reset_index()
    qty_summary.columns = ['Area', 'Budget Qty', 'Billed Qty']
    qty_summary['%'] = (qty_summary['Billed Qty'] / qty_summary['Budget Qty'] * 100).fillna(0).astype(int)

    # ---- Value Summary
    val_summary = df.groupby('Branch')[['Budget_Value', 'Final_Value']].sum().reset_index()
    val_summary.columns = ['Area', 'Budget Value', 'Billed Value']
    val_summary['%'] = (val_summary['Billed Value'] / val_summary['Budget Value'] * 100).fillna(0).astype(int)

    # ---------------- Overall Sales (Raw) Logic ----------------
    overall_sales = sales_df.groupby(data['sales_area_col']).agg({
        data['sales_qty_col']: 'sum',
        data['sales_value_col']: 'sum'
    }).reset_index()

    # Match branch names
    overall_sales = overall_sales.rename(columns={
        data['sales_area_col']: 'Area',
        data['sales_qty_col']: 'Billed Qty',
        data['sales_value_col']: 'Billed Value'
    })

    # Fill missing budget values per area
    total_budget = df.groupby('Branch')[['Budget_Qty', 'Budget_Value']].sum().reset_index()
    total_budget = total_budget.rename(columns={'Branch': 'Area'})

    overall_qty = pd.merge(overall_sales[['Area', 'Billed Qty']], total_budget[['Area', 'Budget_Qty']], on='Area', how='left').fillna(0)
    overall_qty = overall_qty.rename(columns={'Budget_Qty': 'Budget Qty'})

    overall_val = pd.merge(overall_sales[['Area', 'Billed Value']], total_budget[['Area', 'Budget_Value']], on='Area', how='left').fillna(0)
    overall_val = overall_val.rename(columns={'Budget_Value': 'Budget Value'})

    # ---------------- Add TOTAL rows ----------------
    def add_total_row(df, keys):
        total_row = {'Area': 'TOTAL'}
        for key in keys:
            total_row[key] = int(df[key].sum())
        if '%' in df.columns:
            total = (total_row.get(keys[1], 0) / total_row.get(keys[0], 1)) * 100 if total_row.get(keys[0], 0) else 0
            total_row['%'] = int(total)
        return df.append(total_row, ignore_index=True)
    
    qty_summary = clean_and_format_df(qty_summary, column_order=['Area', 'Budget Qty', 'Billed Qty', '%'])
    print("Actual columns in qty_summary:", qty_summary.columns.tolist())
    val_summary = clean_and_format_df(val_summary, column_order=['Area', 'Budget Value', 'Billed Value', '%'])
    overall_qty = clean_and_format_df(overall_qty, column_order=['Area', 'Budget Qty', 'Billed Qty'])
    overall_val = clean_and_format_df(overall_val, column_order=['Area', 'Budget Value', 'Billed Valu'])

    qty_summary = add_total_row(qty_summary, ['Budget Qty', 'Billed Qty'])
    val_summary = add_total_row(val_summary, ['Budget Value', 'Billed Value'])
    overall_qty = add_total_row(overall_qty, ['Budget Qty', 'Billed Qty'])
    overall_val = add_total_row(overall_val, ['Budget Value', 'Billed Value'])

    return {
        'budget_vs_billed_qty':{
            'data': qty_summary.to_dict(orient='records'),
            'columns': qty_summary.columns.tolist(),
        },
        'budget_vs_billed_value': {
            'data': val_summary.to_dict(orient='records'),
            'columns': val_summary.columns.tolist()
        },
        'overall_sales_qty': {
            'data': overall_qty.to_dict(orient='records'),
            'columns': overall_qty.columns.tolist()
        },
        'overall_sales_value': {
            'data': overall_val.to_dict(orient='records'),
            'columns': overall_val.columns.tolist()
        }
}


