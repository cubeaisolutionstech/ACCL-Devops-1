import React, { useState, useEffect, useCallback } from 'react';
import { 
  FileSpreadsheet, 
  Download, 
  AlertCircle, 
  CheckCircle, 
  Info, 
  RefreshCw,
  Search,
  X,
  Database,
  Save,
  FileCheck
} from 'lucide-react';

const API_BASE_URL = process.env.REACT_APP_API_URL || 'http://localhost:5000/api';

// Inline styles for the component (keeping original styles)
const styles = {
  container: {
    padding: '24px',
    maxWidth: '100%',
    overflowX: 'auto'
  },
  sectionHeader: {
    marginBottom: '24px'
  },
  title: {
    color: '#1f2937',
    marginBottom: '8px',
    fontSize: '1.5rem',
    fontWeight: '600'
  },
  description: {
    color: '#6b7280',
    margin: '0'
  },
  actionSection: {
    display: 'flex',
    gap: '12px',
    marginBottom: '24px',
    flexWrap: 'wrap'
  },
  btn: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '8px',
    padding: '10px 16px',
    border: 'none',
    borderRadius: '6px',
    fontSize: '14px',
    fontWeight: '500',
    cursor: 'pointer',
    transition: 'all 0.2s ease',
    textDecoration: 'none'
  },
  btnPrimary: {
    backgroundColor: '#3b82f6',
    color: 'white'
  },
  btnIntegrated: {
    backgroundColor: '#10b981',
    color: 'white'
  },
  btnSecondary: {
    backgroundColor: '#6b7280',
    color: 'white'
  },
  btnSuccess: {
    backgroundColor: '#059669',
    color: 'white'
  },
  btnDisabled: {
    opacity: '0.6',
    cursor: 'not-allowed'
  },
  // Auto-export control styles
  autoExportControl: {
    margin: '20px 0',
    padding: '16px',
    backgroundColor: '#f0f9ff',
    border: '1px solid #0ea5e9',
    borderRadius: '8px'
  },
  toggleLabel: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    cursor: 'pointer',
    fontSize: '14px',
    fontWeight: '500',
    color: '#0c4a6e'
  },
  toggleInput: {
    display: 'none'
  },
  toggleSwitch: {
    position: 'relative',
    width: '44px',
    height: '24px',
    backgroundColor: '#cbd5e1',
    borderRadius: '12px',
    transition: 'background-color 0.3s ease'
  },
  toggleSwitchChecked: {
    backgroundColor: '#0ea5e9'
  },
  toggleSwitchBefore: {
    content: '""',
    position: 'absolute',
    top: '2px',
    left: '2px',
    width: '20px',
    height: '20px',
    backgroundColor: 'white',
    borderRadius: '50%',
    transition: 'transform 0.3s ease'
  },
  toggleSwitchBeforeChecked: {
    transform: 'translateX(20px)'
  },
  toggleHelp: {
    display: 'block',
    marginTop: '8px',
    color: '#64748b',
    fontSize: '12px',
    fontStyle: 'italic',
    lineHeight: '1.4'
  },
  // Storage section styles
  storageSection: {
    backgroundColor: '#f0f9ff',
    border: '1px solid #0ea5e9',
    borderRadius: '8px',
    padding: '16px',
    marginBottom: '24px'
  },
  storageTitle: {
    margin: '0 0 12px 0',
    color: '#0c4a6e',
    fontSize: '16px',
    fontWeight: '600',
    display: 'flex',
    alignItems: 'center',
    gap: '8px'
  },
  storedFileItem: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: '12px',
    backgroundColor: 'white',
    border: '1px solid #e2e8f0',
    borderRadius: '6px',
    marginBottom: '8px'
  },
  storedFileInfo: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px'
  },
  storedFileName: {
    fontWeight: '500',
    color: '#1e40af'
  },
  storedFileSize: {
    fontSize: '12px',
    color: '#64748b'
  },
  storedFileActions: {
    display: 'flex',
    gap: '8px'
  },
  requirementsSection: {
    backgroundColor: '#f9fafb',
    border: '1px solid #e5e7eb',
    borderRadius: '8px',
    padding: '16px',
    marginBottom: '24px'
  },
  requirementsTitle: {
    margin: '0 0 12px 0',
    color: '#374151',
    fontSize: '16px',
    fontWeight: '600'
  },
  requirementsList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px'
  },
  requirement: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    fontSize: '14px'
  },
  requirementMet: {
    color: '#059669'
  },
  requirementUnmet: {
    color: '#dc2626'
  },
  requirementRequired: {
    color: '#dc2626',
    fontWeight: 'bold'
  },
  sessionSection: {
    backgroundColor: '#fef3c7',
    border: '1px solid #f59e0b',
    borderRadius: '8px',
    padding: '16px',
    marginBottom: '24px'
  },
  sessionSectionSuccess: {
    backgroundColor: '#d1fae5',
    border: '1px solid #10b981'
  },
  sessionTitle: {
    margin: '0 0 12px 0',
    color: '#92400e',
    fontSize: '16px',
    fontWeight: '600'
  },
  sessionTitleSuccess: {
    color: '#065f46'
  },
  sessionText: {
    margin: '0 0 12px 0',
    color: '#92400e',
    fontSize: '14px',
    lineHeight: '1.5'
  },
  sessionTextSuccess: {
    color: '#065f46'
  },
  searchSection: {
    marginBottom: '24px'
  },
  searchInputWrapper: {
    position: 'relative',
    maxWidth: '400px'
  },
  searchInput: {
    width: '100%',
    padding: '10px 12px 10px 36px',
    border: '1px solid #d1d5db',
    borderRadius: '6px',
    fontSize: '14px'
  },
  searchIcon: {
    position: 'absolute',
    left: '12px',
    top: '50%',
    transform: 'translateY(-50%)',
    color: '#9ca3af'
  },
  clearSearch: {
    position: 'absolute',
    right: '8px',
    top: '50%',
    transform: 'translateY(-50%)',
    background: 'none',
    border: 'none',
    color: '#9ca3af',
    cursor: 'pointer',
    padding: '4px',
    borderRadius: '4px'
  },
  tablesSection: {
    display: 'flex',
    flexDirection: 'column',
    gap: '32px'
  },
  dataTableContainer: {
    backgroundColor: 'white',
    border: '1px solid #e5e7eb',
    borderRadius: '8px',
    overflow: 'hidden'
  },
  tableHeaderSection: {
    padding: '16px',
    backgroundColor: '#f9fafb',
    borderBottom: '1px solid #e5e7eb',
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    flexWrap: 'wrap',
    gap: '12px'
  },
  tableTitle: {
    margin: '0',
    color: '#111827',
    fontSize: '18px',
    fontWeight: '600'
  },
  tableInfo: {
    display: 'flex',
    gap: '16px',
    fontSize: '14px',
    color: '#6b7280',
    flexWrap: 'wrap'
  },
  tableWrapper: {
    overflowX: 'auto',
    maxHeight: '600px',
    overflowY: 'auto',
    border: '1px solid #e5e7eb',
    borderRadius: '4px'
  },
  dataTable: {
    width: 'max-content',
    minWidth: '100%',
    borderCollapse: 'collapse',
    fontSize: '13px'
  },
  tableHeader: {
    backgroundColor: '#4472c4',
    color: '#ffffff',
    padding: '8px 6px',
    textAlign: 'center',
    fontWeight: '600',
    border: '1px solid #3b5998',
    position: 'sticky',
    top: '0',
    zIndex: '10',
    whiteSpace: 'nowrap',
    fontSize: '10px',
    minWidth: '100px',
    maxWidth: '180px',
    overflow: 'hidden',
    textOverflow: 'ellipsis'
  },
  tableHeaderFirst: {
    textAlign: 'left',
    maxWidth: '200px',
    fontSize: '12px',
    backgroundColor: '#2c5aa0',
    color: '#ffffff',
    minWidth: '150px'
  },
  tableCell: {
    padding: '6px 4px',
    border: '1px solid #e5e7eb',
    whiteSpace: 'nowrap',
    fontSize: '12px'
  },
  textCell: {
    textAlign: 'left',
    maxWidth: '200px',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    fontWeight: '500'
  },
  numberCell: {
    textAlign: 'right',
    fontFamily: 'Courier New, monospace',
    minWidth: '100px',
    maxWidth: '180px',
    overflow: 'hidden',
    textOverflow: 'ellipsis'
  },
  sessionIntegratedRow: {
    backgroundColor: '#d4edda',
    fontWeight: 'bold'
  },
  sessionIntegratedRowCell: {
    borderColor: '#28a745'
  },
  tableNote: {
    padding: '12px 16px',
    backgroundColor: '#fef3c7',
    borderTop: '1px solid #f59e0b',
    color: '#92400e',
    fontSize: '14px',
    textAlign: 'center'
  },
  infoSection: {
    marginTop: '24px'
  },
  infoCard: {
    display: 'flex',
    gap: '12px',
    padding: '16px',
    backgroundColor: '#eff6ff',
    border: '1px solid #bfdbfe',
    borderRadius: '8px',
    color: '#1e40af'
  },
  infoCardTitle: {
    margin: '0 0 4px 0',
    fontSize: '16px',
    fontWeight: '600'
  },
  infoCardText: {
    margin: '0',
    fontSize: '14px',
    lineHeight: '1.5'
  },
  emptyState: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '48px 24px',
    textAlign: 'center',
    color: '#6b7280'
  },
  emptyStateIcon: {
    color: '#d1d5db',
    marginBottom: '16px'
  },
  emptyStateTitle: {
    margin: '0 0 8px 0',
    color: '#374151',
    fontSize: '18px',
    fontWeight: '600'
  },
  emptyStateText: {
    margin: '0',
    fontSize: '14px',
    lineHeight: '1.5'
  }
};

const SalesAnalysisMonthwise = ({ 
  uploadedFiles, 
  selectedSheets, 
  addMessage, 
  loading, 
  setLoading,
  sessionTotals = null,
  setSessionTotals = null,
  // File storage props - only using global storage now
  storedFiles = [],
  setStoredFiles = null,
  onFileAdd = null,
  onFileRemove = null
}) => {
  const [salesMtData, setSalesMtData] = useState(null);
  const [salesValueData, setSalesValueData] = useState(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [processing, setProcessing] = useState(false);
  const [hasCheckedSession, setHasCheckedSession] = useState(false);
  
  // Auto-export state
  const [autoExportEnabled, setAutoExportEnabled] = useState(true);

  // Session storage management
  const saveToSessionStorage = (key, data) => {
    try {
      sessionStorage.setItem(key, JSON.stringify(data));
    } catch (error) {
      console.warn('Failed to save to session storage:', error);
    }
  };

  const loadFromSessionStorage = (key) => {
    try {
      const data = sessionStorage.getItem(key);
      return data ? JSON.parse(data) : null;
    } catch (error) {
      console.warn('Failed to load from session storage:', error);
      return null;
    }
  };

  // Load data from session storage on component mount
  useEffect(() => {
    const savedSalesMtData = loadFromSessionStorage('salesMtData');
    const savedSalesValueData = loadFromSessionStorage('salesValueData');
    
    if (savedSalesMtData) {
      setSalesMtData(savedSalesMtData);
    }
    if (savedSalesValueData) {
      setSalesValueData(savedSalesValueData);
    }
  }, []);

  // Save data to session storage whenever data changes
  useEffect(() => {
    if (salesMtData) {
      saveToSessionStorage('salesMtData', salesMtData);
    }
  }, [salesMtData]);

  useEffect(() => {
    if (salesValueData) {
      saveToSessionStorage('salesValueData', salesValueData);
    }
  }, [salesValueData]);

  // Store file in global storage - KEEP ONLY LATEST FILE
  const storeFileInSession = useCallback(async (fileBlob, fileName, fileDescription) => {
    try {
      const timestamp = new Date().toISOString();
      const fileUrl = URL.createObjectURL(fileBlob);
      
      const storedFileData = {
        id: `sales_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
        name: fileName,
        blob: fileBlob,
        size: fileBlob.size,
        url: fileUrl,
        createdAt: timestamp,
        type: 'sales-analysis-excel',
        source: 'Sales Analysis Monthwise',
        description: fileDescription,
        metadata: {
          mtRows: salesMtData?.data?.length || 0,
          valueRows: salesValueData?.data?.length || 0,
          hasSessionTotals: !!sessionTotals,
          auditorFile: uploadedFiles.auditor?.filename,
          auditorSheet: selectedSheets.auditor,
          autoGenerated: true,
          columnOrdered: true
        },
        tags: ['sales', 'monthwise', 'analysis', 'excel']
      };

      console.log('🔄 Storing Sales Analysis file (replacing previous):', {
        fileName,
        fileSize: fileBlob.size
      });

      // FIRST: Remove any existing Sales Analysis files to keep only the latest
      if (typeof setStoredFiles === 'function') {
        setStoredFiles(prev => {
          // Remove old Sales Analysis files and clean up their URLs
          const oldSalesFiles = prev.filter(f => f.type === 'sales-analysis-excel');
          oldSalesFiles.forEach(file => {
            if (file.url) {
              URL.revokeObjectURL(file.url);
            }
          });
          
          // Keep only non-sales files and add the new one
          const otherFiles = prev.filter(f => f.type !== 'sales-analysis-excel');
          return [storedFileData, ...otherFiles];
        });
        
        console.log('✅ Sales Analysis file replaced (keeping only latest)');
        addMessage(`💾 ${fileName} stored (previous Sales Analysis file replaced)`, 'success');
        return storedFileData;
      }
      
      // Fallback: try onFileAdd if setStoredFiles not available
      if (typeof onFileAdd === 'function') {
        onFileAdd(storedFileData);
        console.log('✅ Sales Analysis file added via onFileAdd callback');
        addMessage(`💾 ${fileName} stored in Combined Excel Manager`, 'success');
        return storedFileData;
      }

      addMessage('⚠️ File generated but could not be stored', 'warning');
      return null;
    } catch (error) {
      console.error('❌ Error storing Sales Analysis file:', error);
      addMessage(`❌ Error storing file: ${error.message}`, 'error');
      return null;
    }
  }, [salesMtData, salesValueData, sessionTotals, uploadedFiles, selectedSheets, addMessage, onFileAdd, setStoredFiles]);

  // Auto-generate Excel file
  const autoGenerateSalesReport = useCallback(async () => {
    if (!salesMtData && !salesValueData) return;

    console.log('🤖 Auto-generating Sales Analysis Excel file...');

    try {
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const fileName = `sales_analysis_auto_${timestamp}.xlsx`;

      // Ensure columns are ordered before sending to backend
      const orderedMtData = orderColumns(salesMtData);
      const orderedValueData = orderColumns(salesValueData);

      const response = await fetch(`${API_BASE_URL}/download-sales-monthwise-excel`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          sales_mt_data: orderedMtData,
          sales_value_data: orderedValueData,
          auto_generated: true
        })
      });

      if (response.ok) {
        const blob = await response.blob();
        
        const fileDescription = `Auto-generated Sales Analysis with MT (${orderedMtData?.data?.length || 0} records) and Value (${orderedValueData?.data?.length || 0} records) with session totals integration`;
        await storeFileInSession(blob, fileName, fileDescription);
        
        addMessage(`🤖 Auto-generated: ${fileName}`, 'success');
        return { success: true, fileName };
      } else {
        const errorData = await response.json();
        throw new Error(errorData.error || 'Failed to generate Excel file');
      }
    } catch (error) {
      console.error('❌ Auto-generation error:', error);
      addMessage(`❌ Auto-generation error: ${error.message}`, 'error');
    }
  }, [salesMtData, salesValueData, storeFileInSession, addMessage]);

  // Auto-generate report when data is available
  useEffect(() => {
    if ((salesMtData || salesValueData) && autoExportEnabled && sessionTotals) {
      const timer = setTimeout(() => {
        autoGenerateSalesReport();
      }, 2000); // 2 second delay after processing completes
      
      return () => clearTimeout(timer);
    }
  }, [salesMtData, salesValueData, autoExportEnabled, sessionTotals, autoGenerateSalesReport]);

  // Auto-process when tab is opened and requirements are met
  useEffect(() => {
    const autoProcess = async () => {
      if (uploadedFiles.auditor && 
          selectedSheets.auditor && 
          sessionTotals && 
          !hasCheckedSession && 
          !processing) {
        
        setHasCheckedSession(true);
        await processSalesAnalysisWithSession();
      }
    };

    autoProcess();
  }, [uploadedFiles.auditor, selectedSheets.auditor, sessionTotals, hasCheckedSession, processing]); // eslint-disable-line react-hooks/exhaustive-deps

  // Process the sales analysis data using session totals
  const processSalesAnalysisWithSession = async () => {
    if (!uploadedFiles.auditor) {
      addMessage('Please upload an auditor file first', 'error');
      return;
    }

    if (!selectedSheets.auditor) {
      addMessage('Please select an auditor sheet first', 'error');
      return;
    }

    if (!sessionTotals) {
      addMessage('Session totals required. Please run Product Analysis first.', 'error');
      return;
    }

    setProcessing(true);
    setLoading(true);

    try {
      addMessage('Processing Sales Analysis with session totals...', 'info');

      const response = await fetch(`${API_BASE_URL}/process-sales-monthwise-with-session`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          filepath: uploadedFiles.auditor.filepath,
          sheet_name: selectedSheets.auditor,
          session_totals: sessionTotals
        })
      });

      const result = await response.json();

      if (result.success) {
        // Ensure columns are in the correct order
        const orderedMtData = orderColumns(result.sales_mt_table);
        const orderedValueData = orderColumns(result.sales_value_table);
        
        setSalesMtData(orderedMtData);
        setSalesValueData(orderedValueData);
        addMessage('✅ Sales Analysis processed with session totals integrated into ACCLLP rows!', 'success');
      } else {
        addMessage(result.error || 'Failed to process sales analysis data', 'error');
        setSalesMtData(null);
        setSalesValueData(null);
      }
    } catch (error) {
      addMessage(`Error processing sales analysis: ${error.message}`, 'error');
      setSalesMtData(null);
      setSalesValueData(null);
    } finally {
      setProcessing(false);
      setLoading(false);
    }
  };

  // Helper function to order columns consistently
  const orderColumns = (tableData) => {
    if (!tableData || !tableData.columns || !tableData.data) return tableData;
    
    // Define the desired column order (adjust as needed)
    const desiredOrder = [
      'Product', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 
      'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar', 'YTD'
    ];
    
    // Filter and order columns based on desired order
    const orderedColumns = desiredOrder.filter(col => 
      tableData.columns.includes(col)
    );
    
    // Add any remaining columns not in the desired order
    const remainingColumns = tableData.columns.filter(col => 
      !desiredOrder.includes(col)
    );
    
    const finalColumns = [...orderedColumns, ...remainingColumns];
    
    // Reorder the data rows
    const orderedData = tableData.data.map(row => {
      const newRow = {};
      finalColumns.forEach(col => {
        if (row.hasOwnProperty(col)) {
          newRow[col] = row[col];
        }
      });
      return newRow;
    });
    
    return {
      ...tableData,
      columns: finalColumns,
      data: orderedData
    };
  };

  // Helper function for file size formatting
  const formatFileSize = (bytes) => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2) + ' ' + sizes[i]);
  };

  // Download and store Excel file
  const downloadAndStoreExcel = async (shouldStore = true) => {
    if (!salesMtData && !salesValueData) {
      addMessage('No data available to export', 'error');
      return;
    }

    try {
      // Ensure columns are ordered before sending to backend
      const orderedMtData = orderColumns(salesMtData);
      const orderedValueData = orderColumns(salesValueData);

      const response = await fetch(`${API_BASE_URL}/download-sales-monthwise-excel`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          sales_mt_data: orderedMtData,
          sales_value_data: orderedValueData
        })
      });

      if (response.ok) {
        const blob = await response.blob();
        
        // Create file info
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const filename = `sales_analysis_manual_${timestamp}.xlsx`;
        
        // Store in global storage if requested
        if (shouldStore) {
          const fileDescription = `Manual Sales Analysis with MT (${orderedMtData?.data?.length || 0} records) and Value (${orderedValueData?.data?.length || 0} records) with session totals integration`;
          await storeFileInSession(blob, filename, fileDescription);
        }
        
        // Also trigger immediate download
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
        
        addMessage(`✅ Excel file ${shouldStore ? 'generated, downloaded, and stored' : 'downloaded'}: ${filename}`, 'success');
      } else {
        const errorData = await response.json();
        addMessage(errorData.error || 'Failed to download Excel file', 'error');
      }
    } catch (error) {
      addMessage(`Error downloading Excel file: ${error.message}`, 'error');
    }
  };

  // Download stored file
  const downloadStoredFile = (storedFile) => {
    if (storedFile.url) {
      const a = document.createElement('a');
      a.href = storedFile.url;
      a.download = storedFile.name;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      addMessage(`📥 Downloaded: ${storedFile.name}`, 'info');
    } else {
      addMessage('File URL not available', 'error');
    }
  };

  // Clear session data
  const clearSessionData = () => {
    try {
      sessionStorage.removeItem('salesMtData');
      sessionStorage.removeItem('salesValueData');
      setSalesMtData(null);
      setSalesValueData(null);
      addMessage('Session data cleared successfully', 'info');
    } catch (error) {
      addMessage('Failed to clear session data', 'error');
    }
  };

  // Filter data based on search term
  const filterData = (data) => {
    if (!data || !data.data || !searchTerm) return data.data;
    
    return data.data.filter(row => 
      Object.values(row).some(value => 
        String(value).toLowerCase().includes(searchTerm.toLowerCase())
      )
    );
  };

  // Format number for display
  const formatNumber = (value) => {
    if (value === "" || value === null || value === undefined || value === 0) return "";
    const num = parseFloat(value);
    if (isNaN(num)) return value;
    return num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  };

  // Get Sales Analysis specific files from global storage
  const salesAnalysisFiles = storedFiles.filter(f => f.type === 'sales-analysis-excel');

  // Check requirements
  const hasAuditorFile = uploadedFiles.auditor;
  const hasAuditorSheet = selectedSheets.auditor;
  const hasSessionData = sessionTotals && Object.keys(sessionTotals).length > 0;
  const canProcess = hasAuditorFile && hasAuditorSheet && hasSessionData;

  // Auto-export toggle component
  const AutoExportToggle = () => (
    <div style={styles.autoExportControl}>
      <label style={styles.toggleLabel}>
        <input
          type="checkbox"
          checked={autoExportEnabled}
          onChange={(e) => setAutoExportEnabled(e.target.checked)}
          style={styles.toggleInput}
        />
        <span style={{
          ...styles.toggleSwitch,
          ...(autoExportEnabled ? styles.toggleSwitchChecked : {}),
          position: 'relative'
        }}>
          <span style={{
            ...styles.toggleSwitchBefore,
            ...(autoExportEnabled ? styles.toggleSwitchBeforeChecked : {}),
            position: 'absolute'
          }}></span>
        </span>
        Auto-generate Excel files for Combined Data Manager
      </label>
      <small style={styles.toggleHelp}>
        When enabled, Excel files are automatically generated and stored when analysis completes
      </small>
    </div>
  );

  // Data table component
  const DataTable = ({ data, title, tableType }) => {
    if (!data || !data.data || data.data.length === 0) {
      return (
        <div style={styles.dataTableContainer}>
          <h3 style={styles.tableTitle}>{title}</h3>
          <div style={styles.emptyState}>
            <AlertCircle size={48} style={styles.emptyStateIcon} />
            <p>No data available for {title}</p>
          </div>
        </div>
      );
    }

    const filteredData = filterData(data);
    const firstCol = data.columns[0];

    return (
      <div style={styles.dataTableContainer}>
        <div style={styles.tableHeaderSection}>
          <h3 style={styles.tableTitle}>{title}</h3>
          <div style={styles.tableInfo}>
            <span>Shape: {data.shape[0]} rows × {data.shape[1]} columns</span>
            {searchTerm && <span>Filtered: {filteredData.length} of {data.data.length} rows</span>}
          </div>
        </div>
        
        <div style={styles.tableWrapper}>
          <table style={styles.dataTable}>
            <thead>
              <tr>
                {data.columns.map((col, index) => (
                  <th 
                    key={index} 
                    title={col}
                    style={{
                      ...styles.tableHeader,
                      ...(index === 0 ? styles.tableHeaderFirst : {}),
                      ...(col.includes('YTD') ? { 
                        minWidth: '160px', 
                        maxWidth: '200px', 
                        fontSize: '9px',
                        whiteSpace: 'normal',
                        wordWrap: 'break-word',
                        lineHeight: '1.2',
                        height: '60px'
                      } : {})
                    }}
                  >
                    {col.includes('YTD') ? col : (col.length > 15 ? `${col.substring(0, 15)}...` : col)}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filteredData.slice(0, 100).map((row, rowIndex) => {
                const isAccllp = String(row[firstCol]).toUpperCase() === 'ACCLLP' || 
                                String(row[firstCol]).toUpperCase() === 'TOTAL SALES';
                
                return (
                  <tr key={rowIndex} style={isAccllp ? styles.sessionIntegratedRow : {}}>
                    {data.columns.map((col, colIndex) => (
                      <td 
                        key={colIndex} 
                        style={{
                          ...styles.tableCell,
                          ...(colIndex === 0 ? styles.textCell : styles.numberCell),
                          ...(isAccllp ? styles.sessionIntegratedRowCell : {}),
                          ...(data.columns[colIndex].includes('YTD') ? { 
                            minWidth: '160px', 
                            maxWidth: '200px' 
                          } : {})
                        }}
                      >
                        {colIndex === 0 ? row[col] : formatNumber(row[col])}
                      </td>
                    ))}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        
        {filteredData.length > 100 && (
          <div style={styles.tableNote}>
            Showing first 100 rows of {filteredData.length} filtered rows
          </div>
        )}
      </div>
    );
  };

  return (
    <div style={styles.container}>
      <style>
        {`
          @keyframes spin {
            from { transform: rotate(0deg); }
            to { transform: rotate(360deg); }
          }
          .spinning {
            animation: spin 1s linear infinite;
          }
        `}
      </style>
      
      <div style={styles.sectionHeader}>
        <h2 style={styles.title}>📊 Sales Analysis Month-wise</h2>
        <p style={styles.description}>
          Automatically processes month-wise sales data with product analysis totals integration
        </p>
      </div>

      {/* Auto-export control */}
      <AutoExportToggle />

      {/* Requirements */}
      <div style={styles.requirementsSection}>
        <h3 style={styles.requirementsTitle}>Requirements:</h3>
        <div style={styles.requirementsList}>
          <div style={{
            ...styles.requirement, 
            ...(hasAuditorFile ? styles.requirementMet : styles.requirementUnmet)
          }}>
            {hasAuditorFile ? <CheckCircle size={16} /> : <AlertCircle size={16} />}
            <span>Auditor file uploaded</span>
          </div>
          <div style={{
            ...styles.requirement, 
            ...(hasAuditorSheet ? styles.requirementMet : styles.requirementUnmet)
          }}>
            {hasAuditorSheet ? <CheckCircle size={16} /> : <AlertCircle size={16} />}
            <span>Auditor sheet selected</span>
          </div>
          <div style={{
            ...styles.requirement, 
            ...(hasSessionData ? styles.requirementMet : styles.requirementRequired)
          }}>
            {hasSessionData ? <CheckCircle size={16} /> : <AlertCircle size={16} />}
            <span><strong>Session totals data (REQUIRED for integration)</strong></span>
          </div>
        </div>
      </div>

      {/* Action Buttons */}
      {canProcess && (
        <div style={styles.actionSection}>
          <button
            onClick={processSalesAnalysisWithSession}
            disabled={processing}
            style={{
              ...styles.btn,
              ...styles.btnIntegrated,
              ...(processing ? styles.btnDisabled : {})
            }}
          >
            {processing ? (
              <>
                <RefreshCw className="spinning" size={16} />
                Processing with Session Totals...
              </>
            ) : (
              <>
                <Database size={16} />
                🔄 Reprocess with Session Totals
              </>
            )}
          </button>

          {(salesMtData || salesValueData) && (
            <>
              <button
                onClick={() => downloadAndStoreExcel(true)}
                style={{...styles.btn, ...styles.btnSuccess}}
                title="Generate, download and store in Combined Data Manager"
              >
                <Save size={16} />
                Generate & Store Excel File
              </button>
              
              <button
                onClick={() => downloadAndStoreExcel(false)}
                style={{...styles.btn, ...styles.btnPrimary}}
                title="Download only (don't store)"
              >
                <Download size={16} />
                Download Only
              </button>
            </>
          )}
        </div>
      )}

      {/* Search */}
      {(salesMtData || salesValueData) && (
        <div style={styles.searchSection}>
          <div style={styles.searchInputWrapper}>
            <Search size={16} style={styles.searchIcon} />
            <input
              type="text"
              placeholder="Search in tables..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              style={styles.searchInput}
            />
            {searchTerm && (
              <button
                onClick={() => setSearchTerm('')}
                style={styles.clearSearch}
              >
                <X size={16} />
              </button>
            )}
          </div>
        </div>
      )}

      {/* Data Tables */}
      <div style={styles.tablesSection}>
        {salesMtData && (
          <DataTable 
            data={salesMtData} 
            title="SALES in MT (with Session Totals)" 
            tableType="MT"
          />
        )}

        {salesValueData && (
          <DataTable 
            data={salesValueData} 
            title="SALES in Value (with Session Totals)" 
            tableType="Value"
          />
        )}
      </div>

      {/* Loading/Empty States */}
      {!canProcess && (
        <div style={styles.emptyState}>
          <FileSpreadsheet size={48} style={styles.emptyStateIcon} />
          <h3 style={styles.emptyStateTitle}>
            {!hasSessionData ? 'Session Totals Required' : 'Ready to Process'}
          </h3>
          <p style={styles.emptyStateText}>
            {!hasSessionData 
              ? 'Please run Product Analysis first to generate session totals, then return to this tab.'
              : 'Upload auditor file and select sheet to automatically process sales analysis.'}
          </p>
        </div>
      )}

      {processing && (
        <div style={styles.emptyState}>
          <RefreshCw className="spinning" size={48} style={styles.emptyStateIcon} />
          <h3 style={styles.emptyStateTitle}>Processing Sales Analysis</h3>
          <p style={styles.emptyStateText}>
            Integrating session totals into ACCLLP rows...
          </p>
        </div>
      )}

      {/* Integration Help */}
      {(!onFileAdd || !setStoredFiles) && (
        <div style={{
          marginTop: '24px',
          padding: '16px',
          backgroundColor: '#fef3c7',
          border: '1px solid #f59e0b',
          borderRadius: '8px',
          color: '#92400e'
        }}>
          <div style={{
            display: 'flex',
            alignItems: 'center',
            gap: '8px',
            marginBottom: '8px',
            fontWeight: '600'
          }}>
            <AlertCircle size={16} />
            Limited Storage Integration
          </div>
          <div style={{ fontSize: '14px', lineHeight: '1.4' }}>
            Some storage functions are not available. Files can still be generated and downloaded, 
            but automatic storage in Combined Data Manager may be limited.
          </div>
        </div>
      )}
    </div>
  );
};

export default SalesAnalysisMonthwise;