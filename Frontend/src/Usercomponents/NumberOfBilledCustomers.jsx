import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { useExcelData } from '../context/ExcelDataContext';
import OdTargetSubTab from './OdTargetSubTab';
import { addReportToStorage } from '../utils/consolidatedStorage';


const NumberOfBilledCustomers = () => {
  const { selectedFiles } = useExcelData();
  const [activeTab, setActiveTab] = useState('nbc');

  const [salesSheets, setSalesSheets] = useState([]);
  const [salesSheet, setSalesSheet] = useState('');
  const [salesHeader, setSalesHeader] = useState(1);

  const [nbcColumns, setNbcColumns] = useState({});
  const [allSalesCols, setAllSalesCols] = useState([]);
  const [filters, setFilters] = useState({ branches: [], executives: [] });
  const [selectedBranches, setSelectedBranches] = useState([]);
  const [selectedExecutives, setSelectedExecutives] = useState([]);
  const [selectAllBranches, setSelectAllBranches] = useState(true);
  const [selectAllExecutives, setSelectAllExecutives] = useState(true);

  const [nbcResults, setNbcResults] = useState({});
  const [loading, setLoading] = useState(false);
  const [loadingAutoMap, setLoadingAutoMap] = useState(false); // Loader for Load Columns & Auto Map
  const [loadingReport, setLoadingReport] = useState(false); // Loader for Generate Billed Customers Report


  const fetchSheetNames = async () => {
    try {
      const res = await axios.post('http://localhost:5000/api/branch/sheets', {
        filename: selectedFiles.salesFile
      });
      setSalesSheets(res.data.sheets || []);
      if (res.data.sheets?.length > 0) {
        setSalesSheet(res.data.sheets[0]);
      }
    } catch (err) {
      console.error("Error fetching sheet names", err);
    }
  };

  const fetchAutoMap = async () => {
    setLoadingAutoMap(true);
  try {
    const res = await axios.post('http://localhost:5000/api/branch/get_columns', {
      filename: selectedFiles.salesFile,
      sheet_name: salesSheet,
      header: salesHeader
    });
    setAllSalesCols(res.data.columns || []);
    console.log("fetched column", res.data.columns);

    const mapRes = await axios.post('http://localhost:5000/api/branch/get_nbc_columns', {
      filename: selectedFiles.salesFile,
      sheet_name: salesSheet,
      header: salesHeader
    });
    setNbcColumns(mapRes.data.mapping || {});
    console.log("map column", mapRes.data.mapping);

   // Auto-fetch filters after column mapping is complete
    if (mapRes.data.mapping) {
      const filterRes = await axios.post('http://localhost:5000/api/branch/get_nbc_filters', {
        filename: selectedFiles.salesFile,
        sheet_name: salesSheet,
        header: salesHeader,
        date_col: mapRes.data.mapping.date,
        branch_col: mapRes.data.mapping.branch,
        executive_col: mapRes.data.mapping.executive
      });
      const { branches, executives } = filterRes.data;
      setFilters({ branches, executives });
      setSelectedBranches(branches);
      setSelectedExecutives(executives);
    }
  } catch (err) {
    console.error("Error in auto mapping NBC columns", err);
  } finally {
    setLoadingAutoMap(false);
  }
};

  const addBranchNBCReportsToStorage = (resultsData) => {
  try {
    const customerReports = [];

    // Loop over each FY (or branch group) in results
    Object.entries(resultsData).forEach(([fyOrGroup, data]) => {
      customerReports.push({
        df: data.data || [],
        title: `BRANCH NBC REPORT - ${fyOrGroup}`,
        percent_cols: [], // No % columns
        columns: ["S.No", "Mapped_Branch", ...(data.months || [])]
      });
    });

    if (customerReports.length > 0) {
      addReportToStorage(customerReports, 'branch_nbc_results');
      console.log(`✅ Stored ${customerReports.length} Branch NBC reports to consolidated storage`);
    }
  } catch (error) {
    console.error("❌ Error storing Branch NBC reports:", error);
  }
};

  const handleGenerateReport = async () => {
    setLoadingReport(true);
    try {
      const payload = {
        filename: selectedFiles.salesFile,
        sheet_name: salesSheet,
        header: salesHeader,
        date_col: nbcColumns.date,
        branch_col: nbcColumns.branch,
        customer_id_col: nbcColumns.customer_id,
        executive_col: nbcColumns.executive,
        selected_branches: selectedBranches,
        selected_executives: selectedExecutives
      };
      const res = await axios.post('http://localhost:5000/api/branch/calculate_nbc_table', payload);
      if (res.data && res.data.results) {
      setNbcResults(res.data.results);
      addBranchNBCReportsToStorage(res.data.results);  // 🔥 Add to storage
    } else {
      console.warn("⚠️ No NBC results to store");
    }
    } catch (err) {
      console.error("Error generating NBC report", err);
    } finally {
      setLoadingReport(false);
    }
  };

  const handleDownloadPPT = async (fyKey) => {
  try {
    const fyData = nbcResults[fyKey];
    const payload = {
      data: fyData.data,
      title: `NUMBER OF BILLED CUSTOMERS - FY ${fyKey}`,
      months: fyData.months,
      financial_year: fyKey,
      logo_file: selectedFiles.logoFile  // optional, if available
    };

    const res = await axios.post("http://localhost:5000/api/branch/download_nbc_ppt", payload, {
      responseType: "blob"
    });

    const blob = new Blob([res.data], {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    });

    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `NBC_Report_FY_${fyKey}.pptx`;
    a.click();
    window.URL.revokeObjectURL(url);
  } catch (err) {
    console.error("❌ Failed to download NBC PPT", err);
    alert("NBC PPT generation failed. See console.");
  }
};

  useEffect(() => {
    if (selectedFiles.salesFile) {
      fetchSheetNames();
    }
  }, [selectedFiles.salesFile]);

  return (
    <div className="p-4">
      <h2 className="text-xl font-bold text-blue-900 mb-4">Number of Billed Customers & OD Target</h2>

      <div className="flex space-x-4 mb-4">
        <button
          onClick={() => setActiveTab('nbc')}
          className={`px-4 py-2 rounded ${activeTab === 'nbc' ? 'bg-blue-700 text-white' : 'bg-gray-200 text-black'}`}
        >
          Number of Billed Customers
        </button>
        <button
          onClick={() => setActiveTab('od')}
          className={`px-4 py-2 rounded ${activeTab === 'od' ? 'bg-blue-700 text-white' : 'bg-gray-200 text-black'}`}
        >
          OD Target
        </button>
      </div>

      {activeTab === 'nbc' && (
        <div>
          {/* Sheet & Header Selection */}
          <div className="grid grid-cols-2 gap-4 mb-4">
            <div>
              <label className="block font-medium">Sales Sheet</label>
              <select className="w-full p-2 border" value={salesSheet} onChange={(e) => setSalesSheet(e.target.value)}>
                <option value="">Select Sheet</option>
                {salesSheets.map(sheet => <option key={sheet}>{sheet}</option>)}
              </select>
            </div>
            <div>
              <label className="block font-medium">Header Row</label>
              <input
                type="number"
                min={1}
                value={salesHeader}
                onChange={(e) => setSalesHeader(Number(e.target.value))}
                className="w-full p-2 border"
              />
            </div>
          </div>

          <button onClick={fetchAutoMap} className="bg-blue-600 text-white px-4 py-2 rounded disabled:bg-gray-400" disabled={loadingAutoMap}>
            {loadingAutoMap ? 'Loading...' : 'Load Columns & Auto Map'}
          </button>

          {Object.keys(nbcColumns).length > 0 && (
            <div className="mt-6">
              <h3 className="font-bold text-blue-800 mb-2">Column Mapping</h3>
              <div className="grid grid-cols-2 gap-4">
                {['date', 'branch', 'customer_id', 'executive'].map((key) => (
                  <div key={key}>
                    <label className="block font-semibold mb-1 capitalize">{key.replace(/_/g, ' ')}</label>
                    <select
                      className="w-full p-2 border"
                      value={nbcColumns[key] || ''}
                      onChange={(e) => setNbcColumns(prev => ({ ...prev, [key]: e.target.value }))}
                    >
                      <option value="">Select Column</option>
                      {allSalesCols.map(col => (
                        <option key={col} value={col}>{col}</option>
                      ))}
                    </select>
                  </div>
                ))}
              </div>
            </div>
          )}

          {Object.keys(nbcColumns).length > 0 && (
            <>
              <div className="grid grid-cols-2 gap-6 mt-4">
                {/* Branch Filter */}
                <div>
                  <label className="block font-medium">Branches</label>
                  <input
                    type="checkbox"
                    checked={selectAllBranches}
                    onChange={() => {
                      const newVal = !selectAllBranches;
                      setSelectAllBranches(newVal);
                      setSelectedBranches(newVal ? filters.branches : []);
                    }}
                  /> Select All
                  <div className="max-h-40 overflow-y-auto border p-2 mt-2 rounded">
                    {filters.branches.map(branch => (
                      <label key={branch} className="flex items-center mb-1">
                        <input
                          type="checkbox"
                          className="mr-2"
                          value={branch}
                          checked={selectedBranches.includes(branch)}
                          onChange={(e) => {
                            const checked = e.target.checked;
                            setSelectedBranches((prev) => {
                              if (checked) return [...prev, branch];
                              return prev.filter((x) => x !== branch);
                            });
                          }}
                        />
                        {branch}
                      </label>
                    ))}
                  </div>
                </div>

                {/* Executive Filter */}
                <div>
                  <label className="block font-medium">Executives</label>
                  <input
                    type="checkbox"
                    checked={selectAllExecutives}
                    onChange={() => {
                      const newVal = !selectAllExecutives;
                      setSelectAllExecutives(newVal);
                      setSelectedExecutives(newVal ? filters.executives : []);
                    }}
                  /> Select All
                  <div className="max-h-40 overflow-y-auto border p-2 mt-2 rounded">
                    {filters.executives.map(exec => (
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
              </div>

              <button onClick={handleGenerateReport} className="mt-6 bg-red-600 text-white px-6 py-2 rounded disabled:bg-gray-400" disabled={loadingReport}>
                {loadingReport ? 'Generating...' : 'Generate Billed Customers Report'}
              </button>
              {Object.entries(nbcResults).map(([fy, result]) => (
  <div key={fy} className="mb-8">
    <h3 className="text-lg font-bold text-blue-700 mb-2">FY {fy}</h3>
    <div className="overflow-auto border rounded">
      <table className="table-auto w-full border text-sm">
        <thead>
          <tr>
            <th className="border px-3 py-1 bg-gray-100">S.No</th>
            <th className="border px-3 py-1 bg-gray-100">Branch Name</th>
            {result.months.map((month) => (
              <th key={month} className="border px-3 py-1 bg-gray-100">{month}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {result.data.map((row, idx) => (
            <tr key={idx}>
              <td className="border px-3 py-1 text-center">{row["S.No"]}</td>
              <td className="border px-3 py-1">{row["Mapped_Branch"]}</td>
              {result.months.map((month) => (
                <td key={month} className="border px-3 py-1 text-right">{row[month] ?? 0}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  </div>
))}
{Object.keys(nbcResults).map((fyKey) => (
  <button
    key={fyKey}
    className="bg-green-600 text-white px-4 py-2 rounded my-2"
    onClick={() => handleDownloadPPT(fyKey)}
  >
     Download NBC Report (FY {fyKey})
  </button>
))}

            </>
          )}
        </div>
      )}

      {activeTab === 'od' && (
  <div className="mt-10">
    <OdTargetSubTab />
  </div>
)}
    </div>
  );
};

export default NumberOfBilledCustomers;
