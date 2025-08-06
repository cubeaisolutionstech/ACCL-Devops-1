import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { useExcelData } from '../context/ExcelDataContext';
import { addReportToStorage } from '../utils/consolidatedStorage';



const ProductGrowth = () => {
  const { selectedFiles } = useExcelData();
  const [loadingFilters, setLoadingFilters] = useState(false);
  const [loadingReport, setLoadingReport] = useState(false); // Loader for Generate Product Growth Report

  const [sheets, setSheets] = useState({ ly: [], cy: [], budget: [] });
  const [selectedSheet, setSelectedSheet] = useState({ ly: 'Sheet1', cy: 'Sheet1', budget: 'Sheet1' });
  const [headers, setHeaders] = useState({ ly: 1, cy: 1, budget: 1 });

  const [columns, setColumns] = useState({ ly: {}, cy: {}, budget: {} });
  const [allCols, setAllCols] = useState({ ly: [], cy: [], budget: [] });
  const [mappingsOverride, setMappingsOverride] = useState({ ly: {}, cy: {}, budget: {} });

  const [availableMonths, setAvailableMonths] = useState({ ly: [], cy: [] });
  const [selectedMonths, setSelectedMonths] = useState({ ly: [], cy: [] });

  const [executives, setExecutives] = useState([]);
  const [companyGroups, setCompanyGroups] = useState([]);
  const [selectedExecutives, setSelectedExecutives] = useState([]);
  const [selectedGroups, setSelectedGroups] = useState([]);

  const [groupResults, setGroupResults] = useState({});

  useEffect(() => {
    if (selectedFiles.lastYearSalesFile && selectedFiles.salesFile && selectedFiles.budgetFile) {
      fetchSheetNames();
    }
  }, [ selectedFiles]);

  const fetchSheetNames = async () => {
    const load = async (filename) => {
      const res = await axios.post('http://localhost:5000/api/branch/sheets', { filename });
      return res.data.sheets || [];
    };
    const [lySheets, cySheets, budgetSheets] = await Promise.all([
      load(selectedFiles.lastYearSalesFile),
      load(selectedFiles.salesFile),
      load(selectedFiles.budgetFile)
    ]);
    setSheets({ ly: lySheets, cy: cySheets, budget: budgetSheets });
  };

  const handleAutoMap = async () => {
    setLoadingFilters(true);
    const fetchColumns = async (filename, sheet, header) => {
      const res = await axios.post('http://localhost:5000/api/branch/get_columns', {
        filename,
        sheet_name: sheet,
        header
      });
      return res.data.columns || [];
    };

    const [lyCols, cyCols, budgetCols] = await Promise.all([
      fetchColumns(selectedFiles.lastYearSalesFile, selectedSheet.ly, headers.ly),
      fetchColumns(selectedFiles.salesFile, selectedSheet.cy, headers.cy),
      fetchColumns(selectedFiles.budgetFile, selectedSheet.budget, headers.budget)
    ]);

    const res = await axios.post('http://localhost:5000/api/branch/auto_map_product_growth', {
      ly_columns: lyCols,
      cy_columns: cyCols,
      budget_columns: budgetCols
    });

    const newMappings = {
      ly: res.data.ly_mapping,
      cy: res.data.cy_mapping,
      budget: res.data.budget_mapping
    };

    setColumns(newMappings);
    setAllCols({ ly: lyCols, cy: cyCols, budget: budgetCols });
    setMappingsOverride(newMappings);
    setLoadingFilters(false);

  };

  // ✅ Run fetchFilters only when mappingsOverride is fully ready
useEffect(() => {
  const hasValidMappings = mappingsOverride.ly?.date && mappingsOverride.cy?.date && mappingsOverride.budget?.executive;

  if (
    selectedFiles.lastYearSalesFile &&
    selectedFiles.salesFile &&
    selectedFiles.budgetFile &&
    selectedSheet.ly &&
    selectedSheet.cy &&
    selectedSheet.budget &&
    hasValidMappings
  ) {
    fetchFilters();
  }
}, [mappingsOverride, selectedSheet, selectedFiles]);

const standardizeName = (name) => {
  if (!name) return "";
  let cleaned = name.trim().toLowerCase().replace(/[^a-zA-Z0-9\s]/g, '');
  cleaned = cleaned.replace(/\s+/g, ' ');
  const titleCase = cleaned.split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');

  const generalVariants = ['general', 'gen', 'generals', 'general ', 'genral', 'generl'];
  if (generalVariants.includes(cleaned)) {
    return 'General';
  }
  return titleCase;
};
useEffect(() => {
  if (availableMonths.ly.length > 0) {
    const defaultLY = availableMonths.ly.includes("0") ? ["0"] : [];
    setSelectedMonths((prev) => ({ ...prev, ly: defaultLY }));
  }

  if (availableMonths.cy.length > 0) {
    setSelectedMonths((prev) => ({ ...prev, cy: [availableMonths.cy[0]] }));
  }
}, [availableMonths]);


  const fetchFilters = async () => {
    setLoadingFilters(true);
  const payload = {
    ly_filename: selectedFiles.lastYearSalesFile,
    ly_sheet: selectedSheet.ly,
    ly_header: headers.ly,
    ly_date_col: mappingsOverride.ly?.date,
    ly_exec_col: mappingsOverride.ly?.executive,
    ly_group_col: mappingsOverride.ly?.company_group,

    cy_filename: selectedFiles.salesFile,
    cy_sheet: selectedSheet.cy,
    cy_header: headers.cy,
    cy_date_col: mappingsOverride.cy?.date,
    cy_exec_col: mappingsOverride.cy?.executive,
    cy_group_col: mappingsOverride.cy?.company_group,

    budget_filename: selectedFiles.budgetFile,
    budget_sheet: selectedSheet.budget,
    budget_header: headers.budget,
    budget_exec_col: mappingsOverride.budget?.executive,
    budget_group_col: mappingsOverride.budget?.company_group
  };

  console.log("📦 fetchFilters payload:", payload);

  try {
    const res = await axios.post('http://localhost:5000/api/branch/get_product_growth_filters', payload);
    console.log("✅ Filters response:", res.data);

    setAvailableMonths({ ly: res.data.ly_months, cy: res.data.cy_months });
    setExecutives(res.data.executives);
    setCompanyGroups( Array.from(new Set(res.data.company_groups.map(standardizeName))).sort());
    setSelectedExecutives(res.data.executives);
    setSelectedGroups(res.data.company_groups);
  } catch (error) {
    console.error("❌ Failed to fetch filters:", error);
    alert("Error loading month and filter options. Check console.");
  }finally {
    setLoadingFilters(false); // ✅ Stop loading
  }
}

  const addBranchProductGrowthToStorage = (resultsData) => {
  try {
    const branchProductGrowthReports = [];

    const lyMonthsLabel = selectedMonths.ly?.join(', ') || 'N/A';
    const cyMonthsLabel = selectedMonths.cy?.join(', ') || 'N/A';

    const resultGroups = resultsData && typeof resultsData === 'object' ? resultsData : {};

    Object.entries(resultGroups).forEach(([branchName, data]) => {
      if (data.qty_df?.length) {
        branchProductGrowthReports.push({
          df: data.qty_df,
          title: `${branchName} - Quantity Growth (Qty in Mt) - LY: ${lyMonthsLabel} vs CY: ${cyMonthsLabel}`,
          percent_cols: [4],
          columns: ["PRODUCT NAME", "LY_QTY", "BUDGET_QTY", "CY_QTY", "ACHIEVEMENT %"]
        });
      }

      if (data.value_df?.length) {
        branchProductGrowthReports.push({
          df: data.value_df,
          title: `${branchName} - Value Growth (Value in Lakhs) - LY: ${lyMonthsLabel} vs CY: ${cyMonthsLabel}`,
          percent_cols: [4],
          columns: ["PRODUCT NAME", "LY_VALUE", "BUDGET_VALUE", "CY_VALUE", "ACHIEVEMENT %"]
        });
      }
    });

    if (resultsData.overall_growth_qty?.length) {
      branchProductGrowthReports.push({
        df: resultsData.overall_growth_qty,
        title: `Overall Branch Growth Quantity Summary (Mt) - LY: ${lyMonthsLabel} vs CY: ${cyMonthsLabel}`,
        percent_cols: [],
        columns: resultsData.overall_growth_qty[0] ? Object.keys(resultsData.overall_growth_qty[0]) : []
      });
    }

    if (resultsData.overall_growth_value?.length) {
      branchProductGrowthReports.push({
        df: resultsData.overall_growth_value,
        title: `Overall Branch Growth Value Summary (Lakhs) - LY: ${lyMonthsLabel} vs CY: ${cyMonthsLabel}`,
        percent_cols: [],
        columns: resultsData.overall_growth_value[0] ? Object.keys(resultsData.overall_growth_value[0]) : []
      });
    }

    if (branchProductGrowthReports.length > 0) {
      console.log("📦 Prepared Product Growth Reports:", branchProductGrowthReports);
      addReportToStorage(branchProductGrowthReports, 'branch_product_results');
      console.log(`✅ Stored ${branchProductGrowthReports.length} Branch Product Growth reports`);
    } else {
      console.warn("⚠️ No product growth reports generated to store.");
    }

  } catch (err) {
    console.error("❌ Error storing branch product growth data:", err);
  }
};



  const handleGenerateReport = async () => {
    setLoadingReport(true);
  try {
    const payload = {
      ly_filename: selectedFiles.lastYearSalesFile,
      ly_sheet: selectedSheet.ly,
      ly_header: headers.ly,
      cy_filename: selectedFiles.salesFile,
      cy_sheet: selectedSheet.cy,
      cy_header: headers.cy,
      budget_filename: selectedFiles.budgetFile,
      budget_sheet: selectedSheet.budget,
      budget_header: headers.budget,

      ly_months: selectedMonths.ly,
      cy_months: selectedMonths.cy,

      ly_date_col: mappingsOverride.ly.date,
      cy_date_col: mappingsOverride.cy.date,
      ly_qty_col: mappingsOverride.ly.quantity,
      cy_qty_col: mappingsOverride.cy.quantity,
      ly_value_col: mappingsOverride.ly.value,
      cy_value_col: mappingsOverride.cy.value,

      budget_qty_col: mappingsOverride.budget.quantity,
      budget_value_col: mappingsOverride.budget.value,

      ly_product_col: mappingsOverride.ly.product_group,
      cy_product_col: mappingsOverride.cy.product_group,
      budget_product_group_col: mappingsOverride.budget.product_group,

      ly_company_group_col: mappingsOverride.ly.company_group,
      cy_company_group_col: mappingsOverride.cy.company_group,
      budget_company_group_col: mappingsOverride.budget.company_group,

      ly_exec_col: mappingsOverride.ly.executive,
      cy_exec_col: mappingsOverride.cy.executive,
      budget_exec_col: mappingsOverride.budget.executive,

      selected_executives: selectedExecutives,
      selected_company_groups: selectedGroups
    };

    const res = await axios.post("http://localhost:5000/api/branch/calculate_product_growth", payload);
    console.log("🧪 Product Growth Backend Result:", res.data);
    if (res && res.data && res.data.results) {
    setGroupResults({ results: res.data.results }); // ✅ Ensure groupResults.results works
    addBranchProductGrowthToStorage(res.data.results);
    }else {
    alert("No results received.");
    }}catch (error) {
    console.error("❌ Failed to calculate product growth:", error);
    alert("Error generating product growth report. See console for details.");
  } finally{
    setLoadingReport(false);
  }
};

const handleDownloadPPT = async () => {
  try {
   const filteredGroups = groupResults?.results
      ? Object.fromEntries(
          Object.entries(groupResults.results).filter(
            ([_, v]) => v && typeof v === "object" && v.qty_df && v.value_df
          )
        )
      : {};

    const payload = {
      group_results: filteredGroups,
      month_title: `LY: ${selectedMonths.ly.join(', ')} vs CY: ${selectedMonths.cy.join(', ')}`
    };

    const response = await axios.post(
      "http://localhost:5000/api/branch/download_product_growth_ppt",
      payload,
      { responseType: "blob" }
    );

    const blob = new Blob([response.data], {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    });

    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Product_Growth_Report_${Date.now()}.pptx`;
    a.click();
    window.URL.revokeObjectURL(url);
  } catch (error) {
    console.error("❌ PPT download failed:", error);
    alert("PPT generation failed. Check console.");
  }
};
// const isArrayNotEmpty = (arr) => Array.isArray(arr) && arr.length > 0;

  return (
    <div className="p-4">
      <h2 className="text-xl font-bold text-blue-800 mb-4">Product Growth</h2>
      <div className="grid grid-cols-3 gap-4 mb-6">
      {['ly', 'cy', 'budget'].map((key) => (
    <div key={key}>
      <label className="font-semibold block mb-1">{key === 'ly' ? 'Last Year Sales' : key === 'cy' ? 'Current Year Sales' : 'Budget'} Sheet</label>
      <select
        className="w-full p-2 border"
        value={selectedSheet[key]}
        onChange={(e) => setSelectedSheet(prev => ({ ...prev, [key]: e.target.value }))}
      >
        <option value="">Select</option>
        {sheets[key].map(s => <option key={s}>{s}</option>)}
      </select>

      <label className="block mt-2">Header Row</label>
      <input
        type="number"
        className="w-full p-2 border"
        value={headers[key]}
        min={1}
        onChange={(e) => setHeaders(prev => ({ ...prev, [key]: Number(e.target.value) }))}
      />
    </div>
  ))}
</div>

{/* Button at the top: only Load Columns & Auto Map visible before columns are loaded */}
{Object.keys(mappingsOverride.ly || {}).length === 0 && (
  <div className="mb-4">
    <button
      onClick={handleAutoMap}
      className="bg-blue-600 text-white px-4 py-2 rounded flex items-center justify-center min-w-[220px]"
      disabled={loadingFilters}
    >
      {loadingFilters ? (
        <>
          <svg className="animate-spin h-5 w-5 mr-2 text-white" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
            <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
            <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"></path>
          </svg>
          Loading...
        </>
      ) : (
        'Load Columns & Auto Map'
      )}
    </button>
  </div>
)}

{/* Column Mapping Preview */}
    {Object.keys(mappingsOverride.ly || {}).length > 0 && (
      <>
      <div className="space-y-10">
        {['ly', 'cy', 'budget'].map((key) => (
          <div key={key}>
            <h3 className="text-lg font-bold text-blue-700 mb-4">
              {key === 'ly' ? 'Last Year Sales' : key === 'cy' ? 'Current Year Sales' : 'Budget'} Mapping
            </h3>
            <div className="grid grid-cols-3 gap-4">
              {Object.entries(columns[key]).map(([colKey, label]) => (
                <div key={colKey}>
                  <label className="block font-semibold mb-1 capitalize">{colKey.replace(/_/g, ' ')}</label>
                  <select
                    className="w-full p-2 border"
                    value={mappingsOverride[key]?.[colKey] || ''}
                    onChange={(e) =>
                      setMappingsOverride((prev) => ({
                        ...prev,
                        [key]: { ...prev[key], [colKey]: e.target.value }
                      }))
                    }
                  >
                    <option value="">Select</option>
                    {allCols[key]?.map((col) => (
                      <option key={col}>{col}</option>
                    ))}
                  </select>
                </div>
              ))}
            </div>
          </div>
        ))}
      </div>
    {loadingFilters && (
  <div className="flex justify-center items-center mt-6 mb-6">
    <div className="animate-spin rounded-full h-8 w-8 border-t-4 border-blue-500 border-opacity-75"></div>
    <span className="ml-3 text-blue-700 font-medium">Loading filters...</span>
  </div>
)}
    
   {/* 🗓️ Month Range Selection */}
{availableMonths.ly.length > 0 && availableMonths.cy.length > 0 && (
  <div className="mt-10">
    <h3 className="text-blue-700 font-semibold text-lg mb-2">📅 Select Month Range</h3>

    <div className="grid grid-cols-2 gap-6">
      
      {/* Last Year Months (with checkboxes) */}
      <div>
        <label className="block mb-1 font-medium">Last Year Months</label>

        {/* Select All */}
        <div className="flex items-center mb-2">
          <input
            type="checkbox"
            className="mr-2"
            checked={selectedMonths.ly.length === availableMonths.ly.length}
            onChange={(e) => {
              setSelectedMonths((prev) => ({
                ...prev,
                ly: e.target.checked ? [...availableMonths.ly] : [],
              }));
            }}
          />
          <span>Select All</span>
        </div>

        {/* Month Checkboxes */}
        <div className="grid grid-cols-2 gap-2">
          {availableMonths.ly.map((month) => (
            <label key={month} className="flex items-center">
              <input
                type="checkbox"
                className="mr-2"
                value={month}
                checked={selectedMonths.ly.includes(month)}
                onChange={(e) => {
                  const checked = e.target.checked;
                  setSelectedMonths((prev) => {
                    const updated = checked
                      ? [...prev.ly, month]
                      : prev.ly.filter((m) => m !== month);
                    return { ...prev, ly: updated };
                  });
                }}
              />
              {month}
            </label>
          ))}
        </div>
      </div>

      {/* Current Year Month (single-select) */}
      <div>
        <label className="block mb-1 font-medium">Current Year Month</label>
        <select
          className="w-full border p-2"
          value={selectedMonths.cy[0] || ''}
          onChange={(e) =>
            setSelectedMonths((prev) => ({
              ...prev,
              cy: [e.target.value],
            }))
          }
        >
          {availableMonths.cy.map((m) => (
            <option key={m} value={m}>{m}</option>
          ))}
        </select>
      </div>
    </div>
  </div>
)}

{loadingFilters && (
  <div className="flex justify-center items-center mt-6 mb-6">
    <div className="animate-spin rounded-full h-8 w-8 border-t-4 border-blue-500 border-opacity-75"></div>
    <span className="ml-3 text-blue-700 font-medium">Loading filters...</span>
  </div>
)}

    {/* 👤 Filters Section */}
{executives.length > 0 && (
  <div className="mt-10">
    <h3 className="text-blue-700 font-semibold text-lg mb-2">🎯 Filter Options</h3>

    <div className="grid grid-cols-2 gap-6">
      {/* Executives */}
      <div>
        <label className="block mb-1 font-medium">Executives</label>
        <div className="flex items-center mb-2">
          <input
            type="checkbox"
            className="mr-2"
            checked={selectedExecutives.length === executives.length}
            onChange={(e) =>
              setSelectedExecutives(e.target.checked ? [...executives] : [])
            }
          />
          <span>Select All</span>
        </div>
        <div className="max-h-40 overflow-y-auto border p-2 rounded">
              {executives.map((exec) => (
                <label key={exec} className="flex items-center mb-1">
                  <input
                    type="checkbox"
                    className="mr-2"
                    value={exec}
                    checked={selectedExecutives.includes(exec)}
                    onChange={(e) => {
                      const checked = e.target.checked;
                      setSelectedExecutives((prev) => {
                        if (checked) return [...prev, exec];
                        return prev.filter((x) => x !== exec);
                      });
                    }}
                  />
                  {exec}
                </label>
              ))}
            </div>
          </div>

      {/* Company Groups */}
      <div>
        <label className="block mb-1 font-medium">Company Groups</label>
        <div className="flex items-center mb-2">
          <input
            type="checkbox"
            className="mr-2"
            checked={selectedGroups.length === companyGroups.length}
            onChange={(e) =>
              setSelectedGroups(e.target.checked ? [...companyGroups] : [])
            }
          />
          <span>Select All</span>
        </div>
        <div className="max-h-40 overflow-y-auto border p-2 rounded">
              {companyGroups.map((grp) => (
                <label key={grp} className="flex items-center mb-1">
                  <input
                    type="checkbox"
                    className="mr-2"
                    value={grp}
                    checked={selectedGroups.includes(grp)}
                    onChange={(e) => {
                      const checked = e.target.checked;
                      setSelectedGroups((prev) => {
                        if (checked) return [...prev, grp];
                        return prev.filter((x) => x !== grp);
                      });
                    }}
                  />
                  {grp}
                </label>
              ))}
            </div>
      </div>
    </div>
  </div>
)}

    {/* Generate Button - only show after columns are loaded */}
    <div className="mt-8">
      <button onClick={handleGenerateReport} className="bg-red-600 text-white px-4 py-2 rounded disabled:bg-gray-400" disabled={loadingReport}>
        {loadingReport ? 'Generating...' : 'Generate Product Growth Report'}
      </button>
    </div>

    {/* Results Table */}
    {groupResults?.results &&
  Object.entries(groupResults.results).map(([group, data]) => (
    <div key={group} className="mb-10 border rounded p-4 shadow">
      {/* Quantity Table */}
      <h4 className="text-lg font-semibold mb-2 text-blue-700">{group} - Quantity Growth (Qty in Mt)</h4>
      <div className="overflow-auto">
        {Array.isArray(data.qty_df) && data.qty_df.length > 0 ? (
          <table className="table-auto w-full border text-sm mb-6">
            <thead>
              <tr>
                {["PRODUCT NAME", "LY_QTY", "BUDGET_QTY", "CY_QTY", "ACHIEVEMENT %"].map((col) => (
                    <th key={col} className="border px-2 py-1 bg-gray-100">{col}</th>
                  ))}
              </tr>
            </thead>
            <tbody>
              {data.qty_df.map((row, idx) => (
                <tr key={idx}>
                  {["PRODUCT NAME", "LY_QTY", "BUDGET_QTY", "CY_QTY", "ACHIEVEMENT %"].map((col) => (
                    <td key={col} className="border px-2 py-1">{row[col]}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        ) : (
          <p className="text-sm text-gray-500">No quantity data available.</p>
        )}
      </div>

      {/* Value Table */}
      <h4 className="text-lg font-semibold mb-2 text-green-700">{group} - Value Growth (Value in Lakhs)</h4>
      <div className="overflow-auto">
        {Array.isArray(data.value_df) && data.value_df.length > 0 ? (
          <table className="table-auto w-full border text-sm">
            <thead>
              <tr>
                {["PRODUCT NAME", "LY_VALUE", "BUDGET_VALUE", "CY_VALUE", "ACHIEVEMENT %"].map((col) => (
                    <th key={col} className="border px-2 py-1 bg-gray-100">{col}</th>
                  ))}
              </tr>
            </thead>
            <tbody>
              {data.value_df.map((row, idx) => (
                <tr key={idx}>
                  {["PRODUCT NAME", "LY_VALUE", "BUDGET_VALUE", "CY_VALUE", "ACHIEVEMENT %"].map((col) => (
                    <td key={col} className="border px-2 py-1">{row[col]}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        ) : (
          <p className="text-sm text-gray-500">No value data available.</p>
        )}
      </div>
    </div>
  ))}

    {/* Download PPT - only show after generation completion */}
    {groupResults?.results && Object.keys(groupResults.results).length > 0 && (
      <button
        onClick={handleDownloadPPT}
        className="mt-6 bg-green-600 hover:bg-green-700 text-white font-bold py-2 px-4 rounded"
      >
        Download Product Growth PPT
      </button>
    )}
  </>
    )}
    </div>
);
}
export default ProductGrowth;
