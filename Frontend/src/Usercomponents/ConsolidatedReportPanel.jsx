import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { getAllReportsForPPT, clearConsolidatedReports } from '../utils/consolidatedStorage';

const ConsolidatedReportPanel = () => {
  const [consolidatedReports, setConsolidatedReports] = useState({});
  const [reportTitle, setReportTitle] = useState('Consolidated Report');
  const [downloading, setDownloading] = useState(false);
  const [error, setError] = useState(null);
  const [refreshKey, setRefreshKey] = useState(0);

  // Load reports on component mount and when refreshKey changes
  useEffect(() => {
    const loadReports = () => {
      const reports = getAllReportsForPPT();
      setConsolidatedReports(reports);
      
    };

    loadReports();
    
    // Listen for storage changes from other tabs/components
    const handleStorageChange = () => {
      loadReports();
    };
    
    window.addEventListener('storage', handleStorageChange);
    return () => window.removeEventListener('storage', handleStorageChange);
  }, [refreshKey]);

  // Calculate total report count
  const totalReports = Array.isArray(consolidatedReports) ? consolidatedReports.length : 0;

  // Estimate total slide count
  const estimateTotalSlides = () => {
    let slideCount = 1; // Title slide
    
    if (Array.isArray(consolidatedReports)) {
      consolidatedReports.forEach(report => {
        if (!report.df || report.df.length === 0) {
          return; // Skip empty reports
        }
        
        // Check if this report has Executive column
        const hasExecutiveCol = report.df.length > 0 && 'Executive' in report.df[0];
        
        if (hasExecutiveCol) {
          // Count non-ACCLP, non-TOTAL executives
          const executiveCount = report.df.filter(row => 
            row.Executive && 
            row.Executive !== 'ACCLP' && 
            row.Executive !== 'TOTAL'
          ).length;
          
          // Use same split logic as backend (threshold: 20)
          if (executiveCount <= 20) {
            slideCount += 1; // Single slide
          } else {
            slideCount += 3; // Part 1, Part 2, Grand Total
          }
        } else {
          // Non-executive reports get 1 slide each
          slideCount += 1;
        }
      });
    }
    
    return slideCount;
  };

  const handleRefresh = () => {
    setRefreshKey(prev => prev + 1);
  };

  const handleClearAll = () => {
    if (window.confirm('Are you sure you want to clear all collected reports?')) {
      clearConsolidatedReports();
      setConsolidatedReports([]);
      setRefreshKey(prev => prev + 1);
    }
  };

  const handleDownloadPowerPoint = async () => {
    try {
      setDownloading(true);
      setError(null);

      // Get all reports from consolidated storage
      const allReportsFlattened = getAllReportsForPPT();
      
      if (!allReportsFlattened || allReportsFlattened.length === 0) {
        setError('No reports available for consolidated PPT generation');
        return;
      }

      console.log(`📊 Generating consolidated PPT with ${allReportsFlattened.length} reports`);
      
      // Log category breakdown
      const categoryBreakdown = {};
      allReportsFlattened.forEach(report => {
        const category = report.category || 'unknown';
        categoryBreakdown[category] = (categoryBreakdown[category] || 0) + 1;
      });
      console.log('📈 Report breakdown by category:', categoryBreakdown);

      // Transform data to backend format - remove category field
      const relaysForBackend = allReportsFlattened.map(report => ({
        df: report.df,
        title: report.title,
        percent_cols: report.percent_cols || [],
        columns: report.columns || []
      }));

      // Generate consolidated PPT
      const response = await axios.post('http://localhost:5000/api/executive/generate_consolidated_ppt', {
        reports_data: relaysForBackend,
        title: reportTitle,
        logo_file: null
      }, {
        responseType: 'blob',
        headers: {
          'Content-Type': 'application/json',
        },
      });

      // Download the generated PPT
      const blob = new Blob([response.data], {
        type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
      });
      
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `Consolidated_Report_${new Date().toISOString().slice(0, 10)}.pptx`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);

      console.log('✅ Consolidated PPT generated and downloaded successfully');

    } catch (error) {
      console.error('❌ Error generating consolidated PPT:', error);
      if (error.response?.data?.error) {
        setError(error.response.data.error);
      } else {
        setError('Failed to generate consolidated PowerPoint presentation');
      }
    } finally {
      setDownloading(false);
    }
  };

  const getReportCounts = () => {
    const counts = {
      budget: 0,
      od_collection: 0,
      product: 0,
      customer: 0,
      od_target: 0
    };
    
    if (Array.isArray(consolidatedReports)) {
      consolidatedReports.forEach(report => {
        const category = report.category || 'unknown';
        if (category === 'budget_results' || category === 'branch_budget_results') {
  counts.budget += 1;
}
        else if (category === 'od_vs_results' || category === 'branch_od_vs_results'){
           counts.od_collection += 1;}
        else if (category === 'product_results' || category === 'branch_product_results'){
          counts.product += 1;}
        else if (category === 'customers_results' || category === 'branch_nbc_results'){
           counts.customer += 1;}
        else if (category === 'od_results' || category === 'branch_od_results_previous' && 'branch_od_results_current') {counts.od_target += 1;}
      });
      console.log(consolidatedReports)
    }
    return counts;
  };
  

  const counts = getReportCounts();

  return (
    <div className="bg-white border rounded-lg shadow-sm p-4 mb-4">
      <div className="flex items-center justify-between mb-4">
        <h3 className="text-lg font-semibold text-gray-900">Summary Report Generator</h3>
        <span className="bg-blue-100 text-blue-800 text-xs font-medium px-2.5 py-0.5 rounded">
          {totalReports} Reports | ~{estimateTotalSlides()} Slides
        </span>
      </div>

      {error && (
        <div className="mb-4 text-red-600 text-sm">{error}</div>
      )}

      {totalReports > 0 ? (
        <>
          {/* Report Counts - Compact Grid */}
          <div className="grid grid-cols-5 gap-2 mb-4 text-xs">
            <div className="text-center p-2 bg-blue-50 rounded">
              <div className="font-medium text-blue-700">{counts.budget}</div>
              <div className="text-blue-600">Budget</div>
            </div>
            <div className="text-center p-2 bg-purple-50 rounded">
              <div className="font-medium text-purple-700">{counts.od_collection}</div>
              <div className="text-purple-600">OD Collection</div>
            </div>
            <div className="text-center p-2 bg-green-50 rounded">
              <div className="font-medium text-green-700">{counts.product}</div>
              <div className="text-green-600">Product</div>
            </div>
            <div className="text-center p-2 bg-orange-50 rounded">
              <div className="font-medium text-orange-700">{counts.customer}</div>
              <div className="text-orange-600">Customer</div>
            </div>
            <div className="text-center p-2 bg-red-50 rounded">
              <div className="font-medium text-red-700">{counts.od_target}</div>
              <div className="text-red-600">OD Target</div>
            </div>
          </div>

          {/* Title Input - Compact */}
          <div className="mb-3">
            <input
              type="text"
              value={reportTitle}
              onChange={(e) => setReportTitle(e.target.value)}
              className="w-full p-2 text-sm border border-gray-300 rounded focus:ring-1 focus:ring-blue-500 focus:border-blue-500"
              placeholder="Report Title"
            />
          </div>

          {/* Action Buttons - Compact Row */}
          <div className="flex gap-2">
            <button
              onClick={handleDownloadPowerPoint}
              disabled={downloading}
              className="flex-1 bg-blue-600 text-white px-3 py-2 text-sm rounded hover:bg-blue-700 disabled:bg-gray-400 transition-colors"
            >
              {downloading ? 'Generating...' : 'Download PPT'}
            </button>
            <button
              onClick={handleRefresh}
              className="px-3 py-2 text-sm bg-gray-500 text-white rounded hover:bg-gray-600 transition-colors"
            >
              Refresh
            </button>
            <button
              onClick={handleClearAll}
              className="px-3 py-2 text-sm bg-red-500 text-white rounded hover:bg-red-600 transition-colors"
            >
              Clear All
            </button>
          </div>
        </>
      ) : (
        <div className="text-center py-6">
          <div className="text-gray-400 mb-2">
            <svg className="mx-auto h-8 w-8" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
            </svg>
          </div>
          <p className="text-sm text-gray-600 mb-3">No reports generated yet</p>
          <button
            onClick={handleRefresh}
            className="px-4 py-2 text-sm bg-blue-600 text-white rounded hover:bg-blue-700"
          >
            Refresh Reports
          </button>
        </div>
      )}
    </div>
  );
};

export default ConsolidatedReportPanel;