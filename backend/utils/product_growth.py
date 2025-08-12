import pandas as pd
import numpy as np
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt
from utils.ppt_generator import create_title_slide, add_table_slide
import logging

logger = logging.getLogger(__name__)

def standardize_name(name):
    if pd.isna(name) or not name:
        return ""
    name = str(name).strip().lower()
    name = ''.join(c for c in name if c.isalnum() or c.isspace())
    name = ' '.join(word.capitalize() for word in name.split())
    general_variants = ['general', 'gen', 'generals', 'general ', 'genral', 'generl']
        
    if any(variant in name.lower() for variant in general_variants):
        return 'General'
    return name

def log_non_numeric_values(df, col):
    """Log non-numeric values in a column before conversion."""
    if col in df.columns:
        non_numeric = df[df[col].apply(lambda x: not isinstance(x, (int, float)) and pd.notna(x))][col].unique()
        if len(non_numeric) > 0:
            logger.warning(f"Non-numeric values in {col}: {non_numeric}")

def find_column_by_names(columns, target_names):
    """Find a column by checking multiple possible names (case-insensitive)."""
    columns_upper = [col.upper() for col in columns]
    target_names_upper = [name.upper() for name in target_names]
    
    for target in target_names_upper:
        if target in columns_upper:
            return columns[columns_upper.index(target)]
    return None

def auto_map_product_growth_columns(ly_columns, cy_columns, budget_columns):

    ly_mapping = {
        'date': find_column_by_names(ly_columns, ['Date', 'Bill Date', 'Invoice Date']),
        'product_group': find_column_by_names(ly_columns, ['Type (Make)', 'Product Group', 'Type', 'Make']),
        'quantity': find_column_by_names(ly_columns, ['Actual Quantity', 'Quantity', 'Qty']),
        'company_group': find_column_by_names(ly_columns, ['Company Group', 'Company', 'Group']),
        'value': find_column_by_names(ly_columns, ['Value', 'Invoice Value', 'Amount']),
        'executive': find_column_by_names(ly_columns, ['Executive', 'Sales Executive', 'Executive Name'])
    }

    cy_mapping = {
        'date': find_column_by_names(cy_columns, ['Month Format', 'Date', 'Bill Date', 'Invoice Date']),
        'product_group': find_column_by_names(cy_columns, ['Type (Make)', 'Product Group', 'Type', 'Make']),
        'quantity': find_column_by_names(cy_columns, ['Actual Quantity', 'Quantity', 'Qty']),
        'company_group': find_column_by_names(cy_columns, ['Company Group', 'Company', 'Group']),
        'value': find_column_by_names(cy_columns, ['Value', 'Invoice Value', 'Amount']),
        'executive': find_column_by_names(cy_columns, ['Executive', 'Sales Executive', 'Executive Name'])
    }

    budget_mapping = {
        'quantity': find_column_by_names(budget_columns, ["Qty – Apr'25", 'Quantity', 'Qty', 'Budget Qty']),
        'company_group': find_column_by_names(budget_columns, ['Company Group', 'Company', 'Group']),
        'value': find_column_by_names(budget_columns, ["Value – Apr'25", 'Value', 'Budget Value', 'Amount']),
        'executive': find_column_by_names(budget_columns, ['Executive Name', 'Executive', 'Sales Executive']),
        'product_group': find_column_by_names(budget_columns, ['Product Group', 'Type(Make)', 'Product', 'Type'])
    }

    return ly_mapping, cy_mapping, budget_mapping


def calculate_product_growth(
    ly_df, cy_df, budget_df, ly_months, cy_months,
    ly_date_col, cy_date_col, ly_qty_col, cy_qty_col,
    ly_value_col, cy_value_col, budget_qty_col, budget_value_col,
    ly_product_col, cy_product_col, ly_company_group_col, cy_company_group_col,
    budget_company_group_col, budget_product_group_col,
    ly_exec_col, cy_exec_col, budget_exec_col,
    selected_executives=None, selected_company_groups=None
):
    try:
        ly_df = ly_df.copy()
        cy_df = cy_df.copy()
        budget_df = budget_df.copy()

        required_cols = [
            (ly_df, [ly_date_col, ly_qty_col, ly_value_col, ly_product_col, ly_company_group_col, ly_exec_col], "Last Year"),
            (cy_df, [cy_date_col, cy_qty_col, cy_value_col, cy_product_col, cy_company_group_col, cy_exec_col], "Current Year"),
            (budget_df, [budget_qty_col, budget_value_col, budget_product_group_col, budget_company_group_col, budget_exec_col], "Budget")
        ]
        for df, cols, name in required_cols:
            for col in cols:
                if col not in df.columns:
                    raise ValueError(f"Missing column '{col}' in {name} data")

        if selected_executives:
            ly_df = ly_df[ly_df[ly_exec_col].isin(selected_executives)]
            cy_df = cy_df[cy_df[cy_exec_col].isin(selected_executives)]
            budget_df = budget_df[budget_df[budget_exec_col].isin(selected_executives)]

        if ly_df.empty or cy_df.empty or budget_df.empty:
            raise ValueError("No data remains after executive filtering")

        ly_df[ly_date_col] = pd.to_datetime(ly_df[ly_date_col], dayfirst=True, errors='coerce')
        cy_df[cy_date_col] = pd.to_datetime(cy_df[cy_date_col], dayfirst=True, errors='coerce')

        ly_filtered_df = ly_df[ly_df[ly_date_col].dt.strftime('%b %y').isin(ly_months)]
        cy_filtered_df = cy_df[cy_df[cy_date_col].dt.strftime('%b %y').isin(cy_months)]

        if ly_filtered_df.empty or cy_filtered_df.empty:
            raise ValueError("No data found for selected LY or CY months")

        # Fix the standardization to avoid SettingWithCopyWarning
        ly_filtered_df = ly_filtered_df.copy()
        cy_filtered_df = cy_filtered_df.copy()
        budget_df = budget_df.copy()
        
        for df, col in [
            (ly_filtered_df, ly_product_col), (cy_filtered_df, cy_product_col), (budget_df, budget_product_group_col),
            (ly_filtered_df, ly_company_group_col), (cy_filtered_df, cy_company_group_col), (budget_df, budget_company_group_col)
        ]:
            df[col] = df[col].fillna("").astype(str).apply(standardize_name)

        if selected_company_groups:
            selected_company_groups = [standardize_name(g) for g in selected_company_groups]
            ly_filtered_df = ly_filtered_df[ly_filtered_df[ly_company_group_col].isin(selected_company_groups)]
            cy_filtered_df = cy_filtered_df[cy_filtered_df[cy_company_group_col].isin(selected_company_groups)]
            budget_df = budget_df[budget_df[budget_company_group_col].isin(selected_company_groups)]

        if ly_filtered_df.empty or cy_filtered_df.empty or budget_df.empty:
            raise ValueError("No data remains after company group filtering")

        for df, qty_col, val_col in [
            (ly_filtered_df, ly_qty_col, ly_value_col),
            (cy_filtered_df, cy_qty_col, cy_value_col),
            (budget_df, budget_qty_col, budget_value_col)
        ]:
            df[qty_col] = pd.to_numeric(df[qty_col], errors='coerce').fillna(0)
            df[val_col] = pd.to_numeric(df[val_col], errors='coerce').fillna(0)

        company_groups = selected_company_groups if selected_company_groups else pd.concat([
            ly_filtered_df[ly_company_group_col],
            cy_filtered_df[cy_company_group_col],
            budget_df[budget_company_group_col]
        ]).dropna().unique().tolist()

        # Helper function to calculate growth percentage
        def calculate_growth_percentage(current_value, last_year_value):
            """Calculate growth percentage with proper handling of edge cases"""
            if pd.isna(last_year_value) or last_year_value == 0:
                if pd.isna(current_value) or current_value == 0:
                    return 0.00
                else:
                    return 100.00  # or could be a large positive number
            return round(((current_value - last_year_value) / last_year_value) * 100, 2)

        result = {}

        for group in company_groups:
            ly_group_df = ly_filtered_df[ly_filtered_df[ly_company_group_col] == group]
            cy_group_df = cy_filtered_df[cy_filtered_df[cy_company_group_col] == group]
            budget_group_df = budget_df[budget_df[budget_company_group_col] == group]

            if ly_group_df.empty and cy_group_df.empty and budget_group_df.empty:
                continue

            ly_products = ly_group_df[ly_product_col].dropna().apply(standardize_name).unique().tolist()
            cy_products = cy_group_df[cy_product_col].dropna().apply(standardize_name).unique().tolist()
            budget_products = budget_group_df[budget_product_group_col].dropna().apply(standardize_name).unique().tolist()
            group_products = sorted(set(ly_products + cy_products + budget_products))

            if group != 'General':
                group_products = [p for p in group_products if p != 'Gc']
            if not group_products:
                continue

            ly_qty = ly_group_df.groupby(ly_product_col)[ly_qty_col].sum().reset_index().rename(columns={ly_product_col: 'PRODUCT NAME', ly_qty_col: 'LY_QTY'})
            cy_qty = cy_group_df.groupby(cy_product_col)[cy_qty_col].sum().reset_index().rename(columns={cy_product_col: 'PRODUCT NAME', cy_qty_col: 'CY_QTY'})
            ly_value = ly_group_df.groupby(ly_product_col)[ly_value_col].sum().reset_index().rename(columns={ly_product_col: 'PRODUCT NAME', ly_value_col: 'LY_VALUE'})
            cy_value = cy_group_df.groupby(cy_product_col)[cy_value_col].sum().reset_index().rename(columns={cy_product_col: 'PRODUCT NAME', cy_value_col: 'CY_VALUE'})
            budget_qty = budget_group_df.groupby(budget_product_group_col)[budget_qty_col].sum().reset_index().rename(columns={budget_product_group_col: 'PRODUCT NAME', budget_qty_col: 'BUDGET_QTY'})
            budget_value = budget_group_df.groupby(budget_product_group_col)[budget_value_col].sum().reset_index().rename(columns={budget_product_group_col: 'PRODUCT NAME', budget_value_col: 'BUDGET_VALUE'})

            all_products_df = pd.DataFrame({'PRODUCT NAME': group_products})

            qty_df = all_products_df.merge(ly_qty, on='PRODUCT NAME', how='left') \
                .merge(cy_qty, on='PRODUCT NAME', how='left') \
                .merge(budget_qty, on='PRODUCT NAME', how='left').fillna(0)

            value_df = all_products_df.merge(ly_value, on='PRODUCT NAME', how='left') \
                .merge(cy_value, on='PRODUCT NAME', how='left') \
                .merge(budget_value, on='PRODUCT NAME', how='left').fillna(0)

            # Calculate achievement percentages for individual products using the helper function
            qty_df['ACHIEVEMENT %'] = qty_df.apply(lambda row: calculate_growth_percentage(row['CY_QTY'], row['LY_QTY']), axis=1)
            value_df['ACHIEVEMENT %'] = value_df.apply(lambda row: calculate_growth_percentage(row['CY_VALUE'], row['LY_VALUE']), axis=1)

            # Round numeric columns
            qty_df[['LY_QTY', 'BUDGET_QTY', 'CY_QTY']] = qty_df[['LY_QTY', 'BUDGET_QTY', 'CY_QTY']].round(2)
            value_df[['LY_VALUE', 'BUDGET_VALUE', 'CY_VALUE']] = value_df[['LY_VALUE', 'BUDGET_VALUE', 'CY_VALUE']].round(2)

            # Reorder columns
            qty_df = qty_df[['PRODUCT NAME', 'LY_QTY', 'BUDGET_QTY', 'CY_QTY', 'ACHIEVEMENT %']]
            value_df = value_df[['PRODUCT NAME', 'LY_VALUE', 'BUDGET_VALUE', 'CY_VALUE', 'ACHIEVEMENT %']]

            # Calculate totals correctly
            total_ly_qty = qty_df['LY_QTY'].sum()
            total_cy_qty = qty_df['CY_QTY'].sum()
            total_budget_qty = qty_df['BUDGET_QTY'].sum()
            
            total_ly_value = value_df['LY_VALUE'].sum()
            total_cy_value = value_df['CY_VALUE'].sum()
            total_budget_value = value_df['BUDGET_VALUE'].sum()

            # Calculate total growth percentages correctly (based on total values, not average of percentages)
            total_qty_growth = calculate_growth_percentage(total_cy_qty, total_ly_qty)
            total_value_growth = calculate_growth_percentage(total_cy_value, total_ly_value)

            qty_totals = pd.DataFrame({
                'PRODUCT NAME': ['TOTAL'],
                'LY_QTY': [round(total_ly_qty, 2)],
                'BUDGET_QTY': [round(total_budget_qty, 2)],
                'CY_QTY': [round(total_cy_qty, 2)],
                'ACHIEVEMENT %': [total_qty_growth]
            })
            qty_df = pd.concat([qty_df, qty_totals], ignore_index=True)

            value_totals = pd.DataFrame({
                'PRODUCT NAME': ['TOTAL'],
                'LY_VALUE': [round(total_ly_value, 2)],
                'BUDGET_VALUE': [round(total_budget_value, 2)],
                'CY_VALUE': [round(total_cy_value, 2)],
                'ACHIEVEMENT %': [total_value_growth]
            })
            value_df = pd.concat([value_df, value_totals], ignore_index=True)

            # Set column attributes for reference
            qty_df.attrs["columns"] = ["PRODUCT NAME", "LY_QTY", "BUDGET_QTY", "CY_QTY", "ACHIEVEMENT %"]
            value_df.attrs["columns"] = ["PRODUCT NAME", "LY_VALUE", "BUDGET_VALUE", "CY_VALUE", "ACHIEVEMENT %"]

            result[group] = {'qty_df': qty_df, 'value_df': value_df}
            print(f"✅ [{group}] qty_df shape: {qty_df.shape}, value_df shape: {value_df.shape}")
            print(f"Total Growth - Qty: {total_qty_growth}%, Value: {total_value_growth}%")
            print(qty_df.head())
            print(value_df.head())

            if qty_df.empty or value_df.empty:
                print(f"⚠️ No data found for {group}")

        if not result:
            raise ValueError("No result generated after product growth computation")
        return result

    except Exception as e:
        import traceback
        logging.error(f"Error in product growth calculation: {e}", exc_info=True)
        logging.error(traceback.format_exc())
        raise RuntimeError(f"Product Growth Calculation failed: {str(e)}")
