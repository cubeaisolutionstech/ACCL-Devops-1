// components/BudgetVsBilled.jsx
import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { useExcelData } from '../context/ExcelDataContext';
import ExecSelector from './ExecSelector';
import BranchSelector from './BranchSelector';
import { addReportToStorage } from '../utils/consolidatedStorage';

const BudgetVsBilled = () => {
  const { selectedFiles } = useExcelData();

  const [salesSheet, setSalesSheet] = useState('Sheet1');
  const [budgetSheet, setBudgetSheet] = useState('Sheet1');
  const [salesSheets, setSalesSheets] = useState([]);
  const [budgetSheets, setBudgetSheets] = useState([]);
  const [salesHeader, setSalesHeader] = useState(1);
  const [budgetHeader, setBudgetHeader] = useState(1);
  const [salesColumns, setSalesColumns] = useState([]);
  const [budgetColumns, setBudgetColumns] = useState([]);
  const [autoMap, setAutoMap] = useState({});
  const [monthOptions, setMonthOptions] = useState([]);
  const [filters, setFilters] = useState({
    selectedMonth: '',
    selectedSalesExecs: [],
    selectedBudgetExecs: [],
    selectedBranches: []
  });
  const [salesExecList, setSalesExecList] = useState([]);
  const [budgetExecList, setBudgetExecList] = useState([]);
  const [branchList, setBranchList] = useState([]);
  const [results, setResults] = useState(null);
  const [loading, setLoading] = useState(false);
  const [loadingColumns, setLoadingColumns] = useState(false); // Loader for Load Columns & Auto Map
  const [loadingReport, setLoadingReport] = useState(false); // Loader for Generate Budget vs Billed Report
  

  // Fetch available sheet names from backend
  const fetchSheets = async (filename, setter) => {
    const res = await axios.post('http://localhost:5000/api/branch/sheets', { filename });
    setter(res.data.sheets);
  };

  useEffect(() => {
    if (selectedFiles.salesFile) fetchSheets(selectedFiles.salesFile, setSalesSheets);
    if (selectedFiles.budgetFile) fetchSheets(selectedFiles.budgetFile, setBudgetSheets);
  }, [selectedFiles]);

  const [columnSelections, setColumnSelections] = useState({
  sales: {},
  budget: {}
  });

  // Fetch columns after user selects sheet + header
  const fetchColumns = async () => {
    if (!salesSheet || !budgetSheet) return;
    setLoadingColumns(true);
    try{
    const getCols = async (filename, sheet_name, header) => {
      const res = await axios.post('http://localhost:5000/api/branch/get_columns', {
        filename,
        sheet_name,
        header
      });
      return res.data.columns || [];
    };

    const [salesCols, budgetCols] = await Promise.all([
      getCols(selectedFiles.salesFile, salesSheet, salesHeader),
      getCols(selectedFiles.budgetFile, budgetSheet, budgetHeader)
    ]);

    setSalesColumns(salesCols);
    setBudgetColumns(budgetCols);

    const res = await axios.post('http://localhost:5000/api/branch/auto_map_columns', {
  sales_columns: salesCols,
  budget_columns: budgetCols
});

const mapping = res.data;

if (
  mapping?.sales_mapping?.executive &&
  mapping?.budget_mapping?.executive &&
  mapping?.sales_mapping?.date
) {
  setAutoMap(mapping);
  setColumnSelections({
  sales: { ...mapping.sales_mapping },
  budget: { ...mapping.budget_mapping }
});

  await fetchExecAndBranches(mapping);
  await fetchMonths(mapping.sales_mapping.date);
} else {
  console.warn("❌ Auto-mapping missing required fields:", mapping);
  alert("Auto-mapping failed: check if 'date' or 'executive' columns were found.");
}
    }catch (err) {
      console.error('Error in fetchcolumn:', err);
    }finally{
      setLoadingColumns(false);
    }
  };

  const fetchExecAndBranches = async (autoMapData) => {
  const salesExec = autoMapData?.sales_mapping?.executive || '';
  const budgetExec = autoMapData?.budget_mapping?.executive || '';
  const salesArea = autoMapData?.sales_mapping?.area || '';
  const budgetArea = autoMapData?.budget_mapping?.area || '';

  const res = await axios.post('http://localhost:5000/api/branch/get_exec_branch_options', {
    sales_filename: selectedFiles.salesFile,
    budget_filename: selectedFiles.budgetFile,
    sales_sheet: salesSheet,
    budget_sheet: budgetSheet,
    sales_header: salesHeader,
    budget_header: budgetHeader,
    sales_exec_col: salesExec,
    budget_exec_col: budgetExec,
    sales_area_col: salesArea,
    budget_area_col: budgetArea
  });

  setSalesExecList(res.data.sales_executives);
  setBudgetExecList(res.data.budget_executives);
  setBranchList(res.data.branches);
};


const fetchMonths = async (dateCol) => {
  if (!dateCol) return;

  const res = await axios.post('http://localhost:5000/api/branch/extract_months', {
    sales_filename: selectedFiles.salesFile,
    sales_sheet: salesSheet,
    sales_header: salesHeader,
    sales_date_col: dateCol
  });

  const months = res.data.months || [];
  setMonthOptions(months);

  // Auto-select first month if available and not already selected
  if (months.length > 0 && !filters.selectedMonth) {
    setFilters((prev) => ({ ...prev, selectedMonth: months[0] }));
  }
};


const addToConsolidatedReports = (resultsData) => {
  try {
    const monthTitle = filters.selectedMonth || 'All Months';

    const columnOrderMap = {
      Qty: ['Area', 'Budget Qty', 'Billed Qty', '%'],
      Value: ['Area', 'Budget Value', 'Billed Value', '%'],
      OverallQty: ['Area', 'Budget Qty', 'Billed Qty'],
      OverallValue: ['Area', 'Budget Value', 'Billed Value']
    };

    const branchReports = [
      {
        df: resultsData.budget_vs_billed_qty.data || [],
        columns: resultsData.budget_vs_billed_qty.columns || [],
        title: `BUDGET AGAINST BILLED (Qty in Mt) - ${monthTitle}`,
        percent_cols: [3]
      },
      {
        df: resultsData.budget_vs_billed_value.data || [],
        columns: resultsData.budget_vs_billed_value.columns || [],
        title: `BUDGET AGAINST BILLED (Value in Lakhs) - ${monthTitle}`,
        percent_cols: [3]
      },
      {
        df: resultsData.overall_sales_qty.data || [],
        columns: resultsData.overall_sales_qty.columns || [],
        title: `OVERALL SALES (Qty in Mt) - ${monthTitle}`,
        percent_cols: []
      },
      {
        df: resultsData.overall_sales_value.data || [],
        columns: resultsData.overall_sales_value.columns || [],
        title: `OVERALL SALES (Value in Lakhs) - ${monthTitle}`,
        percent_cols: []
      }
    ];

    // This version stores ordered columns along with data
    addReportToStorage(branchReports, 'branch_budget_results');

    console.log(`✅ Added ${branchReports.length} branch reports to consolidated storage with column order and category: branch_budget_results`);
  } catch (error) {
    console.error('❌ Error adding branch reports to consolidated storage:', error);
  }
};

const handleCalculate = async () => {
  setLoadingReport(true);
    const payload = {
      sales_filename: selectedFiles.salesFile,
      budget_filename: selectedFiles.budgetFile,
      sales_sheet: salesSheet,
      budget_sheet: budgetSheet,
      sales_header: salesHeader,
      budget_header: budgetHeader,
      selected_month: filters.selectedMonth,
      selected_sales_execs: filters.selectedSalesExecs,
      selected_budget_execs: filters.selectedBudgetExecs,
      selected_branches: filters.selectedBranches,

      sales_date_col: columnSelections.sales.date,
      sales_value_col: columnSelections.sales.value,
      sales_qty_col: columnSelections.sales.quantity,
      sales_area_col: columnSelections.sales.area,
      sales_product_group_col: columnSelections.sales.product_group,
      sales_sl_code_col: columnSelections.sales.sl_code,
      sales_exec_col: columnSelections.sales.executive,

      budget_value_col: columnSelections.budget.value,
      budget_qty_col: columnSelections.budget.quantity,
      budget_area_col: columnSelections.budget.area,
      budget_product_group_col: columnSelections.budget.product_group,
      budget_sl_code_col: columnSelections.budget.sl_code,
      budget_exec_col: columnSelections.budget.executive
    };

    try{
    const res = await axios.post('http://localhost:5000/api/branch/calculate_budget_vs_billed', payload);
    if (res.data && res.data.budget_vs_billed_qty) {
      setResults(res.data);
      addToConsolidatedReports(res.data);
    } else {
      console.warn("No valid data returned from backend.");
    }
  } catch (err) {
    console.error("❌ Error calculating Branch Budget vs Billed:", err);
  } finally{
    setLoadingReport(false);
  }
  };
  
  return (
  <div>
    <h2 className="text-xl font-bold text-blue-800 mb-4">Budget vs Billed Report</h2>

    {/* Sheet Selection */}
    <div className="grid grid-cols-2 gap-6 mb-6">
      <div>
        <label className="block font-semibold mb-1">Sales Sheet Name</label>
        <select className="w-full p-2 border" value={salesSheet} onChange={e => setSalesSheet(e.target.value)}>
          <option value="">Select</option>
          {salesSheets.map(sheet => <option key={sheet}>{sheet}</option>)}
        </select>

        <label className="block mt-4 font-semibold mb-1">Sales Header Row (1-based)</label>
        <input type="number" className="w-full p-2 border" min={1} value={salesHeader}
          onChange={e => setSalesHeader(Number(e.target.value))} />
      </div>

      <div>
        <label className="block font-semibold mb-1">Budget Sheet Name</label>
        <select className="w-full p-2 border" value={budgetSheet} onChange={e => setBudgetSheet(e.target.value)}>
          <option value="">Select</option>
          {budgetSheets.map(sheet => <option key={sheet}>{sheet}</option>)}
        </select>

        <label className="block mt-4 font-semibold mb-1">Budget Header Row (1-based)</label>
        <input type="number" className="w-full p-2 border" min={1} value={budgetHeader}
          onChange={e => setBudgetHeader(Number(e.target.value))} />
      </div>
    </div>

    {/* Load Columns */}
    {!(salesColumns.length > 0 && budgetColumns.length > 0) && (
      <div className="mb-4">
        <button
          className="bg-blue-600 text-white px-6 py-2 rounded hover:bg-blue-700 disabled:bg-gray-400"
          onClick={fetchColumns}
          disabled={!salesSheet || !budgetSheet || loadingColumns}
        >
          {loadingColumns ? 'Loading...' : 'Load Columns & Auto Map'}
        </button>
      </div>
    )}

    {salesColumns.length > 0 && budgetColumns.length > 0 && (
      <>
  <div className="mt-6">
    <h3 className="text-blue-700 font-bold text-lg mb-4">Sales Column Mapping</h3>
    <div className="grid grid-cols-3 gap-4">
      <div>
        <label className="block font-semibold mb-1">Sales Date</label>
        <select
          className="w-full p-2 border"
          value={columnSelections.sales.date || ''}
          onChange={(e) =>
            setColumnSelections(prev => ({
              ...prev,
              sales: { ...prev.sales, date: e.target.value }
            }))
          }
        >
          <option value="">Select</option>
          {salesColumns.map(col => (
            <option key={col} value={col}>{col}</option>
          ))}
        </select>

        <label className="block font-semibold mt-2 mb-1">Sales Area</label>
        <select
          className="w-full p-2 border"
          value={columnSelections.sales.area || ''}
          onChange={(e) =>
            setColumnSelections(prev => ({
              ...prev,
              sales: { ...prev.sales, area: e.target.value }
            }))
          }
        >
          <option value="">Select</option>
          {salesColumns.map(col => (
            <option key={col} value={col}>{col}</option>
          ))}
        </select>
      </div>

      <div>
        <label className="block font-semibold mb-1">Sales Value</label>
        <select
          className="w-full p-2 border"
          value={columnSelections.sales.value || ''}
          onChange={(e) =>
            setColumnSelections(prev => ({
              ...prev,
              sales: { ...prev.sales, value: e.target.value }
            }))
          }
        >
          <option value="">Select</option>
          {salesColumns.map(col => (
            <option key={col} value={col}>{col}</option>
          ))}
        </select>

        <label className="block font-semibold mt-2 mb-1">Sales Quantity</label>
        <select
          className="w-full p-2 border"
          value={columnSelections.sales.quantity || ''}
          onChange={(e) =>
            setColumnSelections(prev => ({
              ...prev,
              sales: { ...prev.sales, quantity: e.target.value }
            }))
          }
        >
          <option value="">Select</option>
          {salesColumns.map(col => (
            <option key={col} value={col}>{col}</option>
          ))}
        </select>
      </div>

      <div>
        <label className="block font-semibold mb-1">Sales Product Group</label>
        <select
          className="w-full p-2 border"
          value={columnSelections.sales.product_group || ''}
          onChange={(e) =>
            setColumnSelections(prev => ({
              ...prev,
              sales: { ...prev.sales, product_group: e.target.value }
            }))
          }
        >
          <option value="">Select</option>
          {salesColumns.map(col => (
            <option key={col} value={col}>{col}</option>
          ))}
        </select>

        <label className="block font-semibold mt-2 mb-1">Sales SL Code</label>
        <select
          className="w-full p-2 border"
          value={columnSelections.sales.sl_code || ''}
          onChange={(e) =>
            setColumnSelections(prev => ({
              ...prev,
              sales: { ...prev.sales, sl_code: e.target.value }
            }))
          }
        >
          <option value="">Select</option>
          {salesColumns.map(col => (
            <option key={col} value={col}>{col}</option>
          ))}
        </select>

        <label className="block font-semibold mt-2 mb-1">Sales Executive</label>
        <select
          className="w-full p-2 border"
          value={columnSelections.sales.executive || ''}
          onChange={(e) =>
            setColumnSelections(prev => ({
              ...prev,
              sales: { ...prev.sales, executive: e.target.value }
            }))
          }
        >
          <option value="">Select</option>
          {salesColumns.map(col => (
            <option key={col} value={col}>{col}</option>
          ))}
        </select>
      </div>
    </div>

    <h3 className="text-blue-700 font-bold text-lg mt-6 mb-4">Budget Column Mapping</h3>
    <div className="grid grid-cols-3 gap-4">
      <div>
        <label className="block font-semibold mb-1">Budget Area</label>
        <select
          className="w-full p-2 border"
          value={columnSelections.budget.area || ''}
          onChange={(e) =>
            setColumnSelections(prev => ({
              ...prev,
              budget: { ...prev.budget, area: e.target.value }
            }))
          }
        >
          <option value="">Select</option>
          {budgetColumns.map(col => (
            <option key={col} value={col}>{col}</option>
          ))}
        </select>

        <label className="block font-semibold mt-2 mb-1">Budget Value</label>
        <select
          className="w-full p-2 border"
          value={columnSelections.budget.value || ''}
          onChange={(e) =>
            setColumnSelections(prev => ({
              ...prev,
              budget: { ...prev.budget, value: e.target.value }
            }))
          }
        >
          <option value="">Select</option>
          {budgetColumns.map(col => (
            <option key={col} value={col}>{col}</option>
          ))}
        </select>
      </div>

      <div>
        <label className="block font-semibold mb-1">Budget Quantity</label>
        <select
          className="w-full p-2 border"
          value={columnSelections.budget.quantity || ''}
          onChange={(e) =>
            setColumnSelections(prev => ({
              ...prev,
              budget: { ...prev.budget, quantity: e.target.value }
            }))
          }
        >
          <option value="">Select</option>
          {budgetColumns.map(col => (
            <option key={col} value={col}>{col}</option>
          ))}
        </select>

        <label className="block font-semibold mt-2 mb-1">Budget Product Group</label>
        <select
          className="w-full p-2 border"
          value={columnSelections.budget.product_group || ''}
          onChange={(e) =>
            setColumnSelections(prev => ({
              ...prev,
              budget: { ...prev.budget, product_group: e.target.value }
            }))
          }
        >
          <option value="">Select</option>
          {budgetColumns.map(col => (
            <option key={col} value={col}>{col}</option>
          ))}
        </select>
      </div>

      <div>
        <label className="block font-semibold mb-1">Budget SL Code</label>
        <select
          className="w-full p-2 border"
          value={columnSelections.budget.sl_code || ''}
          onChange={(e) =>
            setColumnSelections(prev => ({
              ...prev,
              budget: { ...prev.budget, sl_code: e.target.value }
            }))
          }
        >
          <option value="">Select</option>
          {budgetColumns.map(col => (
            <option key={col} value={col}>{col}</option>
          ))}
        </select>

        <label className="block font-semibold mt-2 mb-1">Budget Executive</label>
        <select
          className="w-full p-2 border"
          value={columnSelections.budget.executive || ''}
          onChange={(e) =>
            setColumnSelections(prev => ({
              ...prev,
              budget: { ...prev.budget, executive: e.target.value }
            }))
          }
        >
          <option value="">Select</option>
          {budgetColumns.map(col => (
            <option key={col} value={col}>{col}</option>
          ))}
        </select>
      </div>
    </div>
  </div>

<ExecSelector
  salesExecList={salesExecList}
  budgetExecList={budgetExecList}
  onChange={({ sales, budget }) => {
    setFilters(prev => ({
      ...prev,
      selectedSalesExecs: sales,
      selectedBudgetExecs: budget
    }));
  }}
/>
<BranchSelector
  branchList={branchList}
  onChange={(selected) =>
    setFilters(prev => ({ ...prev, selectedBranches: selected }))
  }
/>
{monthOptions.length > 0 && (
  <div className="mt-6">
    <label className="block font-semibold mb-1">Select Month</label>
    <select
      className="w-full p-2 border"
      value={filters.selectedMonth}
      onChange={(e) => setFilters((prev) => ({ ...prev, selectedMonth: e.target.value }))}
    >
      {monthOptions.map(month => (
        <option key={month} value={month}>{month}</option>
      ))}
    </select>
  </div>
)}


<button
  onClick={handleCalculate}
  className="mt-6 bg-red-600 text-white px-6 py-2 rounded hover:bg-red-700"
  disabled={loadingReport}
>
  {loadingReport ? 'Generating...' : 'Generate Budget vs Billed Report'}
</button>

{results && (
  <div className="mt-8">
    <h3 className="text-lg font-bold text-blue-700 mb-2">Results</h3>

    {['Qty', 'Value', 'OverallQty', 'OverallValue'].map((type, index) => {
      const labelMap = {
        Qty: 'Budget vs Billed Quantity',
        Value: 'Budget vs Billed Value',
        OverallQty: 'Overall Sales Quantity',
        OverallValue: 'Overall Sales Value'
      };
      const resultKeyMap = {
        Qty: 'budget_vs_billed_qty',
        Value: 'budget_vs_billed_value',
        OverallQty: 'overall_sales_qty',
        OverallValue: 'overall_sales_value'
      };
      const resultBlock = results[resultKeyMap[type]] || {};
const rows = resultBlock.data || [];
const orderedCols = resultBlock.columns || [];
      return (
        <div key={type} className="mt-6">
          <h4 className="font-semibold mb-1">{labelMap[type]}</h4>
          <div className="overflow-x-auto">
            <table className="table-auto w-full border text-sm">
              <thead>
                <tr>
                  {orderedCols.map(col => (
                    <th key={col} className="border px-2 py-1 bg-gray-200">{col}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.map((row, i) => (
                  <tr key={i}>
                    {orderedCols.map(col => (
                      <td key={col} className="border px-2 py-1">{row[col]}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      );
    })}
  </div>
)}

{results && (
  <div className="mt-6">
    <button
      className="bg-green-600 text-white px-6 py-2 rounded hover:bg-green-700"
      onClick={async () => {
        const res = await axios.post(
          'http://localhost:5000/api/branch/download_ppt',
          {
            month: filters.selectedMonth,
            budget_vs_billed_qty: results.budget_vs_billed_qty,
            budget_vs_billed_value: results.budget_vs_billed_value,
            overall_sales_qty: results.overall_sales_qty,
            overall_sales_value: results.overall_sales_value,
          },
          { responseType: 'blob' }
        );

        const blob = new Blob([res.data], {
          type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Budget_vs_Billed_${filters.selectedMonth}.pptx`;
        a.click();
        window.URL.revokeObjectURL(url);
      }}
    >
      Download Budget VS Billed PPT
    </button>
  </div>
)}
</>
)}
    </div>
  );
};

export default BudgetVsBilled;
