import pandas as pd
import numpy as np
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import logging

logger = logging.getLogger(__name__)

# =========================
# Auto-mapping (KEEP ORIGINAL)
# =========================
def find_column_by_names(columns, possible_names):
    for name in possible_names:
        for col in columns:
            if str(col).strip().lower() == str(name).strip().lower():
                return col
    return None

def auto_map_od_columns(os_first_columns, os_second_columns, sales_columns):
    os_first_mapping = {
        'due_date': find_column_by_names(os_first_columns, ['Due Date', 'Due_Date', 'DueDate']),
        'branch': find_column_by_names(os_first_columns, ['Branch', 'Unit', 'Area', 'Location']),
        'ref_date': find_column_by_names(os_first_columns, ['Ref. Date', 'Ref Date', 'Reference Date']),
        'net_value': find_column_by_names(os_first_columns, ['Net Value', 'NetValue', 'Amount', 'Value']),
        'executive': find_column_by_names(os_first_columns, ['Executive Name', 'Executive', 'Sales Executive']),
        'region': find_column_by_names(os_first_columns, ['Region', 'Area Region', 'Zone'])
    }
    os_second_mapping = {
        'due_date': find_column_by_names(os_second_columns, ['Due Date', 'Due_Date', 'DueDate']),
        'branch': find_column_by_names(os_second_columns, ['Branch', 'Unit', 'Area', 'Location']),
        'ref_date': find_column_by_names(os_second_columns, ['Ref. Date', 'Ref Date', 'Reference Date']),
        'net_value': find_column_by_names(os_second_columns, ['Net Value', 'NetValue', 'Amount', 'Value']),
        'executive': find_column_by_names(os_second_columns, ['Executive Name', 'Executive', 'Sales Executive']),
        'region': find_column_by_names(os_second_columns, ['Region', 'Area Region', 'Zone'])
    }
    sales_mapping = {
        'bill_date': find_column_by_names(sales_columns, ['Date', 'Bill Date', 'Invoice Date']),
        'branch': find_column_by_names(sales_columns, ['Branch', 'Unit', 'Area', 'Location']),
        'due_date': find_column_by_names(sales_columns, ['Due Date', 'Due_Date', 'DueDate']),
        'value': find_column_by_names(sales_columns, ['Invoice Value', 'Value', 'Amount']),
        'executive': find_column_by_names(sales_columns, ['Executive', 'Sales Executive', 'Executive Name']),
        'region': find_column_by_names(sales_columns, ['Region', 'Area Region', 'Zone'])
    }
    return os_first_mapping, os_second_mapping, sales_mapping

# =========================
# Branch Mapping Helper (KEEP ORIGINAL)
# =========================
def map_branch(val, case='title'):
    if not val or not isinstance(val, str):
        return 'Unknown'
    val = val.strip()
    if case == 'title':
        return val.title()
    return val.upper()

# =========================
# Cumulative Helpers (KEEP ORIGINAL)
# =========================
def get_cumulative_branches(os_first, os_second, total_sale, os_first_unit_col, os_second_unit_col, sale_branch_col):
    branches = set()
    if os_first_unit_col in os_first.columns:
        branches.update(os_first[os_first_unit_col].dropna().apply(lambda x: map_branch(x, 'title')).unique())
    if os_second_unit_col in os_second.columns:
        branches.update(os_second[os_second_unit_col].dropna().apply(lambda x: map_branch(x, 'title')).unique())
    if sale_branch_col in total_sale.columns:
        branches.update(total_sale[sale_branch_col].dropna().apply(lambda x: map_branch(x, 'title')).unique())
    return sorted(b for b in branches if b and b != 'Unknown')

def get_cumulative_regions(os_first, os_second, total_sale, os_first_region_col, os_second_region_col, sale_region_col):
    regions = set()
    if os_first_region_col in os_first.columns:
        regions.update(os_first[os_first_region_col].dropna().astype(str).str.strip().unique())
    if os_second_region_col in os_second.columns:
        regions.update(os_second[os_second_region_col].dropna().astype(str).str.strip().unique())
    if sale_region_col in total_sale.columns:
        regions.update(total_sale[sale_region_col].dropna().astype(str).str.strip().unique())
    return sorted(r for r in regions if r)

# =========================
# Region Mapping (KEEP ORIGINAL)
# =========================
def create_region_branch_mapping(os_first, os_second, total_sale,
                                  os_first_unit_col, os_first_region_col,
                                  os_second_unit_col, os_second_region_col,
                                  sale_branch_col, sale_region_col):
    region_map = {}
    combined = []

    for df, unit_col, region_col in [
        (os_first, os_first_unit_col, os_first_region_col),
        (os_second, os_second_unit_col, os_second_region_col),
        (total_sale, sale_branch_col, sale_region_col)
    ]:
        if unit_col and region_col and unit_col in df.columns and region_col in df.columns:
            temp = df[[unit_col, region_col]].dropna()
            temp['Branch'] = temp[unit_col].apply(lambda x: map_branch(x, 'title'))
            temp['Region'] = temp[region_col].astype(str).str.strip()
            combined.append(temp[['Branch', 'Region']])

    if combined:
        all_mappings = pd.concat(combined).drop_duplicates()
        for region, group in all_mappings.groupby('Region'):
            branches = sorted(b for b in group['Branch'].unique() if b and b != 'Unknown')
            if branches:
                region_map[region.strip()] = branches

    return region_map

# =========================
# SIMPLE Regional Summary - ONLY sum branch values
# =========================
def create_dynamic_regional_summary(df, region_map):
    """SIMPLIFIED: Just sum branch values by region"""
    if df.empty or not region_map:
        return None
    
    # Filter out TOTAL row for regional calculation
    branch_data = df[df['Branch'] != 'TOTAL'].copy()
    
    if branch_data.empty:
        return None
    
    # Map branches to regions
    branch_to_region = {}
    for region, branches in region_map.items():
        for branch in branches:
            branch_to_region[branch] = region
    
    # Add Region column
    branch_data['Region'] = branch_data['Branch'].map(branch_to_region)
    
    # Filter only branches that have region mapping
    regional_data = branch_data[branch_data['Region'].notna()].copy()
    
    if regional_data.empty:
        return None

    # Simple aggregation by region
    numeric_cols = ["Due Target", "Collection Achieved", "For the month Overdue", "For the month Collection"]
    regional = regional_data.groupby('Region')[numeric_cols].sum().reset_index()

    # Calculate percentages
    regional["Overall % Achieved"] = np.where(
        regional["Due Target"] > 0, 
        (regional["Collection Achieved"] / regional["Due Target"]) * 100, 
        0
    )
    regional["% Achieved (Selected Month)"] = np.where(
        regional["For the month Overdue"] > 0, 
        (regional["For the month Collection"] / regional["For the month Overdue"]) * 100, 
        0
    )

    # Round values
    regional[["Due Target", "Collection Achieved"]] = regional[["Due Target", "Collection Achieved"]].round(2)
    regional[["For the month Overdue", "For the month Collection"]] = regional[["For the month Overdue", "For the month Collection"]].round(2)
    regional[["Overall % Achieved", "% Achieved (Selected Month)"]] = regional[["Overall % Achieved", "% Achieved (Selected Month)"]].round(2)

    # Add total row
    total_row = {
        "Region": "TOTAL",
        "Due Target": round(regional["Due Target"].sum(), 2),
        "Collection Achieved": round(regional["Collection Achieved"].sum(), 2),
        "For the month Overdue": round(regional["For the month Overdue"].sum(), 2),
        "For the month Collection": round(regional["For the month Collection"].sum(), 2),
        "Overall % Achieved": round(
            (regional["Collection Achieved"].sum() / regional["Due Target"].sum() * 100) 
            if regional["Due Target"].sum() > 0 else 0, 2
        ),
        "% Achieved (Selected Month)": round(
            (regional["For the month Collection"].sum() / regional["For the month Overdue"].sum() * 100) 
            if regional["For the month Overdue"].sum() > 0 else 0, 2
        )
    }
    regional = pd.concat([regional, pd.DataFrame([total_row])], ignore_index=True)
    
    return regional[["Region", "Due Target", "Collection Achieved", "Overall % Achieved", 
                    "For the month Overdue", "For the month Collection", "% Achieved (Selected Month)"]]

# =========================
# Main Calculation - KEEP YOUR ORIGINAL WORKING LOGIC
# =========================
def calculate_od_values_updated(
    os_first, os_second, total_sale, selected_month_str,
    os_first_due_date_col, os_first_ref_date_col, os_first_unit_col, os_first_net_value_col, os_first_exec_col, os_first_region_col,
    os_second_due_date_col, os_second_ref_date_col, os_second_unit_col, os_second_net_value_col, os_second_exec_col, os_second_region_col,
    sale_bill_date_col, sale_due_date_col, sale_branch_col, sale_value_col, sale_exec_col, sale_region_col,
    selected_executives, selected_branches, selected_regions
):
    try:
        os_first = os_first.copy()
        os_second = os_second.copy()
        total_sale = total_sale.copy()

        # ✅ Validate numeric columns - KEEP ORIGINAL
        for df, col, name in [
            (os_first, os_first_net_value_col, "OS First"),
            (os_second, os_second_net_value_col, "OS Second"),
            (total_sale, sale_value_col, "Sales")
        ]:
            df[col] = pd.to_numeric(df[col], errors='coerce')
            if df[col].isna().all():
                raise ValueError(f"Column '{col}' in {name} contains no valid numeric data.")

        # ✅ Remove negative values - KEEP ORIGINAL
        os_first = os_first[os_first[os_first_net_value_col] >= 0]
        os_second = os_second[os_second[os_second_net_value_col] >= 0]

        # ✅ Date conversion & branch mapping - KEEP ORIGINAL
        os_first[os_first_due_date_col] = pd.to_datetime(os_first[os_first_due_date_col], errors='coerce')
        os_first[os_first_ref_date_col] = pd.to_datetime(os_first[os_first_ref_date_col], errors='coerce') if os_first_ref_date_col else None
        os_first["Branch"] = os_first[os_first_unit_col].apply(lambda x: map_branch(x, case='title'))

        os_second[os_second_due_date_col] = pd.to_datetime(os_second[os_second_due_date_col], errors='coerce')
        os_second[os_second_ref_date_col] = pd.to_datetime(os_second[os_second_ref_date_col], errors='coerce') if os_second_ref_date_col else None
        os_second["Branch"] = os_second[os_second_unit_col].apply(lambda x: map_branch(x, case='title'))

        total_sale[sale_bill_date_col] = pd.to_datetime(total_sale[sale_bill_date_col], errors='coerce')
        total_sale[sale_due_date_col] = pd.to_datetime(total_sale[sale_due_date_col], errors='coerce')
        total_sale["Branch"] = total_sale[sale_branch_col].apply(lambda x: map_branch(x, case='title'))

        # ✅ Region-branch mapping - KEEP ORIGINAL
        region_map = create_region_branch_mapping(
            os_first, os_second, total_sale,
            os_first_unit_col, os_first_region_col,
            os_second_unit_col, os_second_region_col,
            sale_branch_col, sale_region_col
        )

        # Remove rows where Branch is null or "Unknown" - KEEP ORIGINAL
        os_first = os_first[os_first["Branch"].notna() & (os_first["Branch"] != "Unknown")]
        os_second = os_second[os_second["Branch"].notna() & (os_second["Branch"] != "Unknown")]
        total_sale = total_sale[total_sale["Branch"].notna() & (total_sale["Branch"] != "Unknown")]

        # ✅ Filters - KEEP ORIGINAL
        if selected_executives:
            os_first = os_first[os_first[os_first_exec_col].isin(selected_executives)]
            os_second = os_second[os_second[os_second_exec_col].isin(selected_executives)]
            total_sale = total_sale[total_sale[sale_exec_col].isin(selected_executives)]

        if selected_branches:
            os_first = os_first[os_first["Branch"].isin(selected_branches)]
            os_second = os_second[os_second["Branch"].isin(selected_branches)]
            total_sale = total_sale[total_sale["Branch"].isin(selected_branches)]

        if selected_regions and region_map:
            allowed = [b for r in selected_regions for b in region_map.get(r, [])]
            os_first = os_first[os_first["Branch"].isin(allowed)]
            os_second = os_second[os_second["Branch"].isin(allowed)]
            total_sale = total_sale[total_sale["Branch"].isin(allowed)]

        if os_first.empty or os_second.empty or total_sale.empty:
            logger.error("Error in calculate_od_values_updated: One or more datasets are empty after filtering. Cannot compute results.")
            raise ValueError("One or more datasets are empty after filtering. Cannot compute results.")

        # ✅ Date range - KEEP ORIGINAL
        month_str = selected_month_str.replace(" ", "-")
        specified_date = pd.to_datetime("01-" + month_str, format="%d-%b-%y")
        month_end = specified_date + pd.offsets.MonthEnd(0)

        # === Branch-wise Calculations - KEEP ORIGINAL ===
        due_target = os_first[os_first[os_first_due_date_col] <= month_end]
        due_target_sum = due_target.groupby("Branch")[os_first_net_value_col].sum().reset_index()
        due_target_sum.columns = ["Branch", "Due Target"]

        jan_coll = os_first[os_first[os_first_due_date_col] <= month_end].groupby("Branch")[os_first_net_value_col].sum().reset_index()
        jan_coll.columns = ["Branch", "OS Jan Coll"]

        feb_coll = os_second[(os_second[os_second_due_date_col] <= month_end) & 
                             (os_second[os_second_ref_date_col] < specified_date)].groupby("Branch")[os_second_net_value_col].sum().reset_index()
        feb_coll.columns = ["Branch", "OS Feb Coll"]

        collection = jan_coll.merge(feb_coll, on="Branch", how="outer").fillna(0)
        collection["Collection Achieved"] = collection["OS Jan Coll"] - collection["OS Feb Coll"]
        collection["Overall % Achieved"] = np.where(collection["OS Jan Coll"] > 0, 
                                                    (collection["Collection Achieved"] / collection["OS Jan Coll"]) * 100, 
                                                    0)
        collection = collection.merge(due_target_sum, on="Branch", how="left").fillna(0)

        # === Overdue + Month Collection - KEEP ORIGINAL ===
        overdue = total_sale[(total_sale[sale_bill_date_col].between(specified_date, month_end)) &
                             (total_sale[sale_due_date_col].between(specified_date, month_end))]
        overdue_sum = overdue.groupby("Branch")[sale_value_col].sum().reset_index()
        overdue_sum.columns = ["Branch", "For the month Overdue"]

        sales_sum = overdue.groupby("Branch")[sale_value_col].sum().reset_index()
        sales_sum.columns = ["Branch", "Sale Value"]

        os_month = os_second[(os_second[os_second_ref_date_col].between(specified_date, month_end)) &
                             (os_second[os_second_due_date_col].between(specified_date, month_end))]
        os_month_sum = os_month.groupby("Branch")[os_second_net_value_col].sum().reset_index()
        os_month_sum.columns = ["Branch", "OS Month Collection"]

        month_result = sales_sum.merge(os_month_sum, on="Branch", how="outer").fillna(0)
        month_result["For the month Collection"] = month_result["Sale Value"] - month_result["OS Month Collection"]

        # === Final Merge - KEEP ORIGINAL ===
        final = collection.drop(columns=["OS Jan Coll", "OS Feb Coll"]).merge(overdue_sum, on="Branch", how="outer") \
                          .merge(month_result[["Branch", "For the month Collection"]], on="Branch", how="outer").fillna(0)

        final["% Achieved (Selected Month)"] = np.where(final["For the month Overdue"] > 0,
                                                        (final["For the month Collection"] / final["For the month Overdue"]) * 100, 0)

        # 💰 Convert to lakhs and round - KEEP ORIGINAL
        value_cols = ["Due Target", "Collection Achieved", "For the month Overdue", "For the month Collection"]
        final[value_cols] = final[value_cols].div(100000)
        round_cols = value_cols + ["Overall % Achieved", "% Achieved (Selected Month)"]
        final[round_cols] = final[round_cols].round(2)

        # Reorder - KEEP ORIGINAL
        final = final[["Branch", "Due Target", "Collection Achieved", "Overall % Achieved", 
                       "For the month Overdue", "For the month Collection", "% Achieved (Selected Month)"]]
        final.sort_values("Branch", inplace=True)

        # 🌍 Regional Summary - SIMPLIFIED TO JUST SUM VALUES
        regional_summary = create_dynamic_regional_summary(final, region_map)

        # ➕ Total Row - FIXED: Use exact formula like regional summary
        total_row = {'Branch': 'TOTAL'}
        for col in final.columns[1:]:
            if col == "Overall % Achieved":
                # Use exact formula: (Total Collection Achieved / Total Due Target) * 100
                total_due = final["Due Target"].sum()
                total_achieved = final["Collection Achieved"].sum()
                total_row[col] = round((total_achieved / total_due * 100) if total_due > 0 else 0, 2)
            elif col == "% Achieved (Selected Month)":
                # Use exact formula: (Total Month Collection / Total Month Overdue) * 100
                total_overdue = final["For the month Overdue"].sum()
                total_month_collection = final["For the month Collection"].sum()
                total_row[col] = round((total_month_collection / total_overdue * 100) if total_overdue > 0 else 0, 2)
            else:
                total_row[col] = round(final[col].sum(), 2)

        final = pd.concat([final, pd.DataFrame([total_row])], ignore_index=True)

        return final, regional_summary, region_map

    except Exception as e:
        logger.error(f"Error in calculate_od_values_updated: {e}")
        raise RuntimeError(f"OD Target Calculation failed: {str(e)}")