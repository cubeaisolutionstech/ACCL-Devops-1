// components/OdTargetVsCollection.jsx
import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { useExcelData } from '../context/ExcelDataContext';
import { addReportToStorage } from '../utils/consolidatedStorage';


const OdTargetVsCollection = () => {
const { selectedFiles } = useExcelData();

const [sheets, setSheets] = useState({
osPrev: [], osCurr: [], sales: []
});
const [selectedSheet, setSelectedSheet] = useState({
osPrev: 'Sheet1', osCurr: 'Sheet1', sales: 'Sheet1'
});
const [headers, setHeaders] = useState({
osPrev: 1, osCurr: 1, sales: 1
});
const [columns, setColumns] = useState({
osPrev: [], osCurr: [], sales: []
});
const [mappings, setMappings] = useState({
osPrev: {}, osCurr: {}, sales: {}
});

const [monthOptions, setMonthOptions] = useState([]);
const [selectedMonth, setSelectedMonth] = useState('');

const [execList, setExecList] = useState([]);
const [branchList, setBranchList] = useState([]);
const [regionList, setRegionList] = useState([]);

const [filters, setFilters] = useState({
selectedExecs: [], selectedBranches: [], selectedRegions: []
});

const [regionMap, setRegionMap] = useState({});
const [results, setResults] = useState({ branch_summary: [], regional_summary: [] });
const [activeResultTab, setActiveResultTab] = useState('branch');

const [loadingColumns, setLoadingColumns] = useState(false); // Loader for Load Columns & Auto Map
const [loadingReport, setLoadingReport] = useState(false); // Loader for Generate OD Target vs Collection Report

// Fetch sheet names for all 3 files
const fetchSheetNames = async () => {
const load = async (filename) => {
const res = await axios.post('http://localhost:5000/api/branch/sheets', { filename });
return res.data.sheets || [];
};
const [osPrevSheets, osCurrSheets, salesSheets] = await Promise.all([
load(selectedFiles.osPrevFile), load(selectedFiles.osCurrFile), load(selectedFiles.salesFile)
]);
setSheets({ osPrev: osPrevSheets, osCurr: osCurrSheets, sales: salesSheets });
};

useEffect(() => {
if (selectedFiles.osPrevFile && selectedFiles.osCurrFile && selectedFiles.salesFile) {
fetchSheetNames();
}

if (branchList.length > 0 || execList.length > 0 || regionList.length > 0) {
    setFilters(prev => ({
      ...prev,
      selectedBranches: branchList,
      selectedExecs: execList,
      selectedRegions: regionList
    }));
  }
}, [branchList, execList, regionList, selectedFiles]);



// Load columns + auto mapping
const fetchColumnsAndMap = async () => {
  setLoadingColumns(true);
  try{
const getCols = async (filename, sheet, header) => {
const res = await axios.post('http://localhost:5000/api/branch/get_columns', {
filename, sheet_name: sheet, header
});
return res.data.columns || [];
};
const [osPrevCols, osCurrCols, salesCols] = await Promise.all([
  getCols(selectedFiles.osPrevFile, selectedSheet.osPrev, headers.osPrev),
  getCols(selectedFiles.osCurrFile, selectedSheet.osCurr, headers.osCurr),
  getCols(selectedFiles.salesFile, selectedSheet.sales, headers.sales)
]);

setColumns({ osPrev: osPrevCols, osCurr: osCurrCols, sales: salesCols });

const res = await axios.post('http://localhost:5000/api/branch/get_od_columns', {
  os_prev_columns: osPrevCols,
  os_curr_columns: osCurrCols,
  sales_columns: salesCols
});

const fetchFilters = async (mappingData) => {
  const res = await axios.post('http://localhost:5000/api/branch/get_od_filter_options', {
    os_prev_filename: selectedFiles.osPrevFile,
    os_curr_filename: selectedFiles.osCurrFile,
    sales_filename: selectedFiles.salesFile,
    os_prev_sheet: selectedSheet.osPrev,
    os_curr_sheet: selectedSheet.osCurr,
    sales_sheet: selectedSheet.sales,
    os_prev_header: headers.osPrev,
    os_curr_header: headers.osCurr,
    sales_header: headers.sales,
    os_prev_mapping: mappingData.os_jan_mapping,
    os_curr_mapping: mappingData.os_feb_mapping,
    sales_mapping: mappingData.sales_mapping
  });

  setExecList(res.data.executives || []);
  setBranchList(res.data.branches || []);
  setRegionList(res.data.regions || []);
};


setMappings(res.data);
await fetchFilters(res.data);
fetchMonths(res.data.sales_mapping?.bill_date);
  }catch(err){
    console.error('Error in fetchColumnsAndMap:', err);
  }finally{
    setLoadingColumns(false);
  }
};


const fetchMonths = async (dateCol) => {
  try{
const res = await axios.post('http://localhost:5000/api/branch/extract_months', {
sales_filename: selectedFiles.salesFile,
sales_sheet: selectedSheet.sales,
sales_header: headers.sales,
sales_date_col: dateCol
});
const months = res.data.months || [];

setMonthOptions(months);
if (months.length > 0) {
      setSelectedMonth(months[0]); // Or use a logic to pick current month if needed
    }
  } catch (error) {
    console.error('Error fetching months:', error);
    setMonthOptions([]);
  }
};

const addBranchODReportsToStorage = (resultsData, selectedMonth) => {
  try {
    const branchODReports = [];

    // 🟦 Add Branch Summary table
    if (resultsData.branch_summary?.length) {
      branchODReports.push({
        df: resultsData.branch_summary,
        title: `OD TARGET vs COLLECTION (Branch-wise) - ${selectedMonth}`,
        percent_cols: [3, 6],
        columns: branchColumns
      });
    }

    // 🟪 Add Regional Summary table
    if (resultsData.regional_summary?.length) {
      branchODReports.push({
        df: resultsData.regional_summary,
        title: `OD TARGET vs COLLECTION (Region-wise) - ${selectedMonth}`,
        percent_cols: [3, 6],
        columns: regionColumns
      });
    }

    if (branchODReports.length > 0) {
      addReportToStorage(branchODReports, 'branch_od_vs_results');
      console.log(`✅ Added ${branchODReports.length} Branch OD reports to consolidated storage`);
    }
  } catch (error) {
    console.error('❌ Error saving Branch OD reports to storage:', error);
  }
};


const handleCalculate = async () => {
  setLoadingReport(true);
  try {
    const payload = {
      os_prev_filename: selectedFiles.osPrevFile,
      os_curr_filename: selectedFiles.osCurrFile,
      sales_filename: selectedFiles.salesFile,
      os_prev_sheet: selectedSheet.osPrev,
      os_curr_sheet: selectedSheet.osCurr,
      sales_sheet: selectedSheet.sales,
      os_prev_header: headers.osPrev,
      os_curr_header: headers.osCurr,
      sales_header: headers.sales,
      selected_month: selectedMonth,
      os_prev_mapping: mappings.os_jan_mapping,
      os_curr_mapping: mappings.os_feb_mapping,
      sales_mapping: mappings.sales_mapping,
      selected_executives:[],// filters.selectedExecs,
      selected_branches:[],// filters.selectedBranches,
      selected_regions: [],//filters.selectedRegions,
    };

    console.log("📤 OD API Payload", payload);

    const res = await axios.post('http://localhost:5000/api/branch/calculate_od_target', payload);

    console.log("✅ OD API Response:", res.data);

    if (!res.data.branch_summary.length) {
      alert("⚠️ No data returned. Your filters may be excluding all rows. Try selecting fewer filters.");
      return;
    }

    setResults({
      branch_summary: res.data.branch_summary || [],
      regional_summary: res.data.regional_summary || [],
    });
    addBranchODReportsToStorage(res.data, selectedMonth);

    setRegionMap(res.data.region_mapping || {});
  } catch (error) {
    console.error("❌ Error calculating OD Target:", error);

    // Try to extract backend message
    const errorMsg = error?.response?.data?.error || "OD Target calculation failed. Please check console.";
    alert("❌ " + errorMsg);
  } finally{
    setLoadingReport(false);
  }
};

const branchColumns = [
  "Branch",
  "Due Target",
  "Collection Achieved",  
  "Overall % Achieved",
  "For the month Overdue",
  "For the month Collection",
  "% Achieved (Selected Month)"
];

const regionColumns = [
  "Region",
  "Due Target",
  "Collection Achieved",  
  "Overall % Achieved",
  "For the month Overdue",
  "For the month Collection",
  "% Achieved (Selected Month)"  
]

return (
<div className="p-4">
<h2 className="text-xl font-bold text-blue-800 mb-4">OD Target vs Collection</h2>
  {/* Sheet selection */}
  <div className="grid grid-cols-3 gap-4 mb-6">
    {['osPrev', 'osCurr', 'sales'].map((type) => (
      <div key={type}>
        <label className="font-semibold block mb-1">{type === 'osPrev' ? 'OS Previous' : type === 'osCurr' ? 'OS Current' : 'Sales'} Sheet</label>
        <select className="w-full p-2 border" value={selectedSheet[type]} onChange={e => setSelectedSheet(prev => ({ ...prev, [type]: e.target.value }))}>
          <option value="">Select</option>
          {sheets[type].map(s => <option key={s}>{s}</option>)}
        </select>
        <label className="block mt-2">Header Row</label>
        <input type="number" min={1} className="w-full p-2 border" value={headers[type]} onChange={e => setHeaders(prev => ({ ...prev, [type]: Number(e.target.value) }))} />
      </div>
    ))}
  </div>

  {/* Buttons at the top: only Load Columns & Auto Map visible before columns are loaded */}
  {Object.keys(mappings.sales_mapping || {}).length === 0 && (
    <div className="mb-4">
      <button className="bg-blue-600 text-white px-4 py-2 rounded disabled:bg-gray-400" onClick={fetchColumnsAndMap} disabled={loadingColumns}>
        {loadingColumns ? 'Loading...' : 'Load Columns & Auto Map'}
      </button>
    </div>
  )}

  {Object.keys(mappings.sales_mapping || {}).length > 0 && (
    <>
    <div className="mt-6">
      {/* You will render manual mapping dropdowns here (in next step) */}
      <h3 className="text-blue-700 font-semibold mb-2">Column Mapping loaded successfully!</h3>

      {/* OS Previous Month Mapping */}
    <div className="mb-6">
      <h4 className="text-blue-600 font-semibold mb-2">OS - Previous Month</h4>
      <div className="grid grid-cols-3 gap-4">
        {['due_date', 'ref_date', 'branch', 'net_value', 'executive', 'region'].map(key => (
          <div key={key}>
            <label className="block font-semibold capitalize mb-1">{key.replace('_', ' ')}</label>
            <select
              className="w-full p-2 border"
              value={mappings.os_jan_mapping?.[key] || ''}
              onChange={e =>
                setMappings(prev => ({
                  ...prev,
                  os_jan_mapping: { ...prev.os_jan_mapping, [key]: e.target.value }
                }))
              }
            >
              <option value="">Select</option>
              {columns.osPrev.map(col => (
                <option key={col} value={col}>{col}</option>
              ))}
            </select>
          </div>
        ))}
      </div>
    </div>
    {/* OS Current Month Mapping */}
    <div className="mb-6">
      <h4 className="text-blue-600 font-semibold mb-2">OS - Current Month</h4>
      <div className="grid grid-cols-3 gap-4">
        {['due_date', 'ref_date', 'branch', 'net_value', 'executive', 'region'].map(key => (
          <div key={key}>
            <label className="block font-semibold capitalize mb-1">{key.replace('_', ' ')}</label>
            <select
              className="w-full p-2 border"
              value={mappings.os_feb_mapping?.[key] || ''}
              onChange={e =>
                setMappings(prev => ({
                  ...prev,
                  os_feb_mapping: { ...prev.os_feb_mapping, [key]: e.target.value }
                }))
              }
            >
              <option value="">Select</option>
              {columns.osCurr.map(col => (
                <option key={col} value={col}>{col}</option>
              ))}
            </select>
          </div>
        ))}
      </div>
    </div>
    {/* Sales Mapping */}
    <div className="mb-6">
      <h4 className="text-blue-600 font-semibold mb-2">Sales</h4>
      <div className="grid grid-cols-3 gap-4">
        {['bill_date', 'due_date', 'branch', 'value', 'executive', 'region'].map(key => (
          <div key={key}>
            <label className="block font-semibold capitalize mb-1">{key.replace('_', ' ')}</label>
            <select
              className="w-full p-2 border"
              value={mappings.sales_mapping?.[key] || ''}
              onChange={e =>
                setMappings(prev => ({
                  ...prev,
                  sales_mapping: { ...prev.sales_mapping, [key]: e.target.value }
                }))
              }
            >
              <option value="">Select</option>
              {columns.sales.map(col => (
                <option key={col} value={col}>{col}</option>
              ))}
            </select>
          </div>
        ))}
      </div>
    </div>
    </div>
  {monthOptions.length > 0 && (
    <div className="mt-6">
      <label className="block font-semibold mb-1">Select Sales Month</label>
      <select className="w-full p-2 border" value={selectedMonth} onChange={e => setSelectedMonth(e.target.value)}>
        <option value="">Select</option>
        {monthOptions.map(month => <option key={month} value={month}>{month}</option>)}
      </select>
    </div>
  )}

  {branchList.length > 0 && (
  <div className="mt-8">
    <h3 className="font-semibold text-blue-800 mb-2">Filter Options</h3>

    <div className="grid grid-cols-3 gap-6">
      {/* Branch Filter */}
      <div>
        <label className="font-medium">Branches</label>
        <div className="mt-1">
          <label className="inline-flex items-center">
            <input
              type="checkbox"
              checked={filters.selectedBranches.length === branchList.length}
              onChange={(e) => {
                setFilters(prev => ({
                  ...prev,
                  selectedBranches: e.target.checked ? branchList : []
                }));
              }}
            />
            <span className="ml-2">Select All</span>
          </label>
        </div>
        <div className="max-h-48 overflow-y-auto border p-2 rounded">
          {branchList.map(b => (
            <label key={b} className="block text-sm">
              <input
                type="checkbox"
                checked={filters.selectedBranches.includes(b)}
                onChange={(e) => {
                  const updated = e.target.checked
                    ? [...filters.selectedBranches, b]
                    : filters.selectedBranches.filter(x => x !== b);
                  setFilters(prev => ({ ...prev, selectedBranches: updated }));
                }}
              />
              <span className="ml-2">{b}</span>
            </label>
          ))}
        </div>
      </div>
      {/* Executive Filter */}
      <div>
        <label className="font-medium">Executives</label>
        <div className="mt-1">
          <label className="inline-flex items-center">
            <input
              type="checkbox"
              checked={filters.selectedExecs.length === execList.length}
              onChange={(e) => {
                setFilters(prev => ({
                  ...prev,
                  selectedExecs: e.target.checked ? execList : []
                }));
              }}
            />
            <span className="ml-2">Select All</span>
          </label>
        </div>
        <div className="max-h-48 overflow-y-auto border p-2 rounded">
          {execList.map(exec => (
            <label key={exec} className="block text-sm">
              <input
                type="checkbox"
                checked={filters.selectedExecs.includes(exec)}
                onChange={(e) => {
                  const updated = e.target.checked
                    ? [...filters.selectedExecs, exec]
                    : filters.selectedExecs.filter(x => x !== exec);
                  setFilters(prev => ({ ...prev, selectedExecs: updated }));
                }}
              />
              <span className="ml-2">{exec}</span>
            </label>
          ))}
        </div>
      </div>

      {/* Region Filter */}
      <div>
        <label className="font-medium">Regions</label>
        <div className="mt-1">
          <label className="inline-flex items-center">
            <input
              type="checkbox"
              checked={filters.selectedRegions.length === regionList.length}
              onChange={(e) => {
                setFilters(prev => ({
                  ...prev,
                  selectedRegions: e.target.checked ? regionList : []
                }));
              }}
            />
            <span className="ml-2">Select All</span>
          </label>
        </div>
        <div className="max-h-48 overflow-y-auto border p-2 rounded">
          {regionList.map(region => (
            <label key={region} className="block text-sm">
              <input
                type="checkbox"
                checked={filters.selectedRegions.includes(region)}
                onChange={(e) => {
                  const updated = e.target.checked
                    ? [...filters.selectedRegions, region]
                    : filters.selectedRegions.filter(x => x !== region);
                  setFilters(prev => ({ ...prev, selectedRegions: updated }));
                }}
              />
              <span className="ml-2">{region}</span>
            </label>
          ))}
        </div>
      </div>
    </div>
  </div>
)}

  {/* Place the Generate button immediately after filters, before results */}
    <div className="mt-8">
      <button className="bg-red-600 text-white px-4 py-2 rounded disabled:bg-gray-400" onClick={handleCalculate} disabled={loadingReport}>
        {loadingReport ? 'Generating...' : 'Generate OD Target vs Collection Report'}
      </button>
    </div>
    </>
)}

  {/* === Results Display in 3 Tabs === */}
{results.branch_summary.length > 0 && (
  <div className="mt-8">
    <h3 className="text-lg font-bold mb-3 text-blue-700">OD Target vs Collection Results</h3>

    <div className="flex space-x-4 mb-4">
      <button
        onClick={() => setActiveResultTab('branch')}
        className={`px-4 py-2 rounded ${activeResultTab === 'branch' ? 'bg-blue-600 text-white' : 'bg-blue-100 text-black'}`}
      >
        Branch-wise
      </button>
      {results.regional_summary.length > 0 && (
        <button
          onClick={() => setActiveResultTab('regional')}
          className={`px-4 py-2 rounded ${activeResultTab === 'regional' ? 'bg-blue-600 text-white' : 'bg-blue-100 text-black'}`}
        >
          Regional Summary
        </button>
      )}
      <button
        onClick={() => setActiveResultTab('mapping')}
        className={`px-4 py-2 rounded ${activeResultTab === 'mapping' ? 'bg-blue-600 text-white' : 'bg-blue-100 text-black'}`}
      >
        Region Mapping
      </button>
    </div>

    {/* Branch Table */}
    {activeResultTab === 'branch' && (
      <div>
        <h4 className="font-semibold mb-2 text-blue-800">Branch-wise Performance</h4>
        <table className="table-auto w-full border text-sm">
  <thead>
    <tr>
      {branchColumns.map(col => (
        <th key={col} className="border px-2 py-1">{col}</th>
      ))}
    </tr>
  </thead>
  <tbody>
    {results.branch_summary.map((row, i) => (
      <tr key={i}>
        {branchColumns.map(col => (
          <td key={col} className="border px-2 py-1">{row[col]}</td>
        ))}
      </tr>
    ))}
  </tbody>
</table>
      </div>
    )}

    {/* Regional Table */}
    {activeResultTab === 'regional' && results.regional_summary.length > 0 && (
      <div>
        <h4 className="font-semibold mb-2 text-blue-800">Regional Summary</h4>
        <table className="table-auto w-full border text-sm">
          <thead>
            <tr>
              {[
                "Region",
                "Due Target",
                "Collection Achieved",
                "Overall % Achieved",
                "For the month Overdue",
                "For the month Collection",
                "% Achieved (Selected Month)"
              ].map(col => (
                <th key={col} className="border px-2 py-1">{col}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {results.regional_summary.map((row, i) => (
              <tr key={i}>
                {[
                  "Region",
                  "Due Target",
                  "Collection Achieved",
                  "Overall % Achieved",
                  "For the month Overdue",
                  "For the month Collection",
                  "% Achieved (Selected Month)"
                ].map(col => (
                  <td key={col} className="border px-2 py-1">{row[col]}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    )}

    {/* Region Mapping Table */}
    {activeResultTab === 'mapping' && (
      <div>
        <h4 className="font-semibold mb-2 text-blue-800">Region-Branch Mapping</h4>
        <table className="w-full border text-sm">
          <thead className="bg-gray-100">
            <tr>
              <th className="border px-2 py-1">Region</th>
              <th className="border px-2 py-1">Branches</th>
            </tr>
          </thead>
          <tbody>
            {Object.entries(regionMap).map(([region, branches]) => (
              <tr key={region}>
                <td className="border px-2 py-1">{region}</td>
                <td className="border px-2 py-1">{branches.join(', ')}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    )}
  </div>
)}
{results.branch_summary.length > 0 && (
  <div className="mt-6">
    <button
      onClick={async () => {
        try {
          const payload = {
            branch_summary: {
              data: results.branch_summary,
              columns: branchColumns, // use same array as table
            },
            regional_summary: {
              data: results.regional_summary,
              columns: [
                "Region",
                "Due Target",
                "Collection Achieved",
                "Overall % Achieved",
                "For the month Overdue",
                "For the month Collection",
                "% Achieved (Selected Month)"
              ]},
            title: `OD Target vs Collection - ${selectedMonth}`
          };

          const res = await axios.post(
            'http://localhost:5000/api/branch/download_od_ppt',
            payload,
            { responseType: 'blob' } // required to receive binary blob
          );

          const blob = new Blob([res.data], {
            type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
          });

          const url = window.URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = `OD_Target_vs_Collection_${selectedMonth}.pptx`;
          a.click();
          window.URL.revokeObjectURL(url);
        } catch (err) {
          console.error('❌ PPT download failed:', err);
          alert('Failed to download PPT. See console for details.');
        }
      }}
      className="bg-green-600 text-white px-5 py-2 rounded hover:bg-green-700"
    >
      Download OD Target PPT
    </button>
  </div>
)}
</div>
);
};

export default OdTargetVsCollection;