import React, { useState, useCallback, useEffect } from 'react';
import { 
  Download, 
  File, 
  Trash2, 
  Eye, 
  FileSpreadsheet,
  Database,
  Layers,
  BarChart3,
  RefreshCw,
  MapPin,
  TrendingUp,
  Package,
  Merge,
  Activity
} from 'lucide-react';

const CombinedExcelManager = ({ 
  regionData, 
  fiscalInfo, 
  addMessage, 
  loading, 
  setLoading,
  storedFiles = [],
  setStoredFiles = () => {},
  onFilesClear = () => {}
}) => {
  // State Management
  const [filePreview, setFilePreview] = useState(null);
  const [processing, setProcessing] = useState(false);
  const [selectedFiles, setSelectedFiles] = useState([]);
  const [combiningCategory, setCombiningCategory] = useState(null);

  const API_BASE_URL = process.env.REACT_APP_API_URL || 'http://localhost:5000/api';

  // Region Sheet Column Definitions with formatting specifications
  const regionMtColumns = [
    { header: "Region", key: "region", type: "text" },
    { header: "Product", key: "product", type: "text" },
    { header: "Month", key: "month", type: "text" },
    { header: "Fiscal Year", key: "fiscalYear", type: "text" },
    { header: "Quantity (MT)", key: "quantity", type: "number", format: "0.00" },
    { header: "Growth %", key: "growthPercentage", type: "number", format: "0.00" },
    { header: "Market Share %", key: "marketShare", type: "number", format: "0.00" },
    { header: "Avg Price", key: "avgPrice", type: "number", format: "0.00" },
    { header: "Previous Year Qty", key: "prevYearQty", type: "number", format: "0.00" },
    { header: "YTD Qty", key: "ytdQty", type: "number", format: "0.00" },
    { header: "YTD Growth %", key: "ytdGrowth", type: "number", format: "0.00" },
    { header: "Notes", key: "notes", type: "text" }
  ];

  const regionValueColumns = [
    { header: "Region", key: "region", type: "text" },
    { header: "Product", key: "product", type: "text" },
    { header: "Month", key: "month", type: "text" },
    { header: "Fiscal Year", key: "fiscalYear", type: "text" },
    { header: "Value (₹)", key: "value", type: "number", format: "0.00" },
    { header: "Growth %", key: "growthPercentage", type: "number", format: "0.00" },
    { header: "Market Share %", key: "marketShare", type: "number", format: "0.00" },
    { header: "Avg Price", key: "avgPrice", type: "number", format: "0.00" },
    { header: "Previous Year Value", key: "prevYearValue", type: "number", format: "0.00" },
    { header: "YTD Value", key: "ytdValue", type: "number", format: "0.00" },
    { header: "YTD Growth %", key: "ytdGrowth", type: "number", format: "0.00" },
    { header: "Notes", key: "notes", type: "text" }
  ];

  // Enhanced file categorization function
  const categorizeFiles = useCallback((files) => {
    const categories = {
      sales: [],
      region: [],
      product: [],
      ts_pw: [],
      ero_pw: [],
      combined: []
    };

    files.forEach(file => {
      const fileType = file.type?.toLowerCase() || '';
      const fileName = file.name?.toLowerCase() || '';
      const source = file.source?.toLowerCase() || '';

      if (fileType.includes('sales') || fileName.includes('sales') || source.includes('sales') || fileType === 'sales-analysis-excel') {
        categories.sales.push(file);
      } else if (fileType.includes('region') || fileName.includes('region') || source.includes('region')) {
        categories.region.push(file);
      } else if (fileType.includes('product') || fileName.includes('product') || source.includes('product') || file.metadata?.analysisType === 'product') {
        categories.product.push(file);
      } else if (fileType.includes('tspw') || fileType.includes('ts-pw') || fileName.includes('ts')) {
        categories.ts_pw.push(file);
      } else if (fileType.includes('eropw') || fileType.includes('ero-pw') || fileName.includes('ero')) {
        categories.ero_pw.push(file);
      } else if (['combined', 'master-combined', 'selected-combined', 'master-combined-enhanced', 'selected-combined-enhanced', 'category-based-combined'].includes(fileType)) {
        categories.combined.push(file);
      }
    });

    return categories;
  }, []);

  // Utility functions
  const formatFileSize = (bytes) => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  const formatDate = (dateString) => {
    return new Date(dateString).toLocaleString('en-IN', {
      year: 'numeric',
      month: 'short',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    });
  };

  const getFileTypeIcon = (fileType) => {
    switch (fileType) {
      case 'region':
      case 'region-mt':
      case 'region-value':
        return <MapPin size={20} />;
      case 'combined':
      case 'master-combined':
      case 'selected-combined':
      case 'master-combined-enhanced':
      case 'selected-combined-enhanced':
      case 'category-based-combined':
        return <Merge size={20} />;
      case 'product':
        return <Package size={20} />;
      case 'sales-analysis-excel':
      case 'sales':
        return <TrendingUp size={20} />;
      case 'eropw':
        return <BarChart3 size={20} />;
      case 'tspw':
        return <Activity size={20} />;
      default:
        return <FileSpreadsheet size={20} />;
    }
  };

  const ensureUniqueId = (file, index, categoryPrefix) => {
    return {
      ...file,
      uniqueKey: `${categoryPrefix}_${file.id}_${index}_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`
    };
  };

  // Computed values
  const fileGroups = categorizeFiles(storedFiles);
  const hasData = (regionData.mt && regionData.mt.data && regionData.mt.data.length > 0) || 
                  (regionData.value && regionData.value.data && regionData.value.data.length > 0);

  const categorySummary = Object.entries(fileGroups).map(([category, files]) => ({
    name: category,
    count: files.length,
    icon: {
      sales: '📊',
      region: '🌍',
      product: '📦',
      ts_pw: '⚡',
      ero_pw: '📈',
      combined: '🔗'
    }[category] || '📄'
  })).filter(cat => cat.count > 0);

  // Enhanced downloadCategoryExcel function
  const downloadCategoryExcel = useCallback(async (categoryName, categoryFiles) => {
    setCombiningCategory(categoryName);
    setProcessing(true);
    try {
      if (categoryFiles.length === 0) {
        throw new Error('No files in this category');
      }

      // If only one file, download it directly
      if (categoryFiles.length === 1) {
        const file = categoryFiles[0];
        if (!file.blob) {
          throw new Error('File content not available');
        }

        const url = URL.createObjectURL(file.blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${categoryName.replace('_', '-')}_${fiscalInfo.fiscal_year_str || 'report'}_${new Date().toISOString().slice(0,10)}.xlsx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        addMessage(`✅ Downloaded ${categoryName.replace('_', ' ')} Excel file`, 'success');
        return;
      }

      // Multiple files - combine all sheets from all files in this category
      addMessage(`🔄 Combining ${categoryFiles.length} files from ${categoryName.replace('_', ' ')} category with centered titles and formatting...`, 'info');

      const filesToCombine = [];
      
      // Prepare files for combining
      for (const file of categoryFiles) {
        if (file.blob) {
          try {
            const arrayBuffer = await file.blob.arrayBuffer();
            const base64String = btoa(
              new Uint8Array(arrayBuffer).reduce(
                (data, byte) => data + String.fromCharCode(byte), ''
              )
            );
            
            filesToCombine.push({
              name: `${categoryName}_${file.name.replace('.xlsx', '')}`,
              content: base64String,
              originalName: file.name,
              fileType: file.type,
              source: file.source,
              metadata: {
                ...file.metadata,
                formattingApplied: file.metadata?.formattingApplied || true,
                decimalPlaces: file.metadata?.decimalPlaces || 2,
                numberFormat: file.metadata?.numberFormat || '0.00',
                centerTitles: true
              }
            });
          } catch (error) {
            console.warn(`Failed to process file ${file.name}:`, error);
            addMessage(`⚠️ Skipped file ${file.name} due to processing error`, 'warning');
          }
        }
      }

      if (filesToCombine.length === 0) {
        throw new Error('No valid files to combine in this category');
      }

      const response = await fetch(`${API_BASE_URL}/combined/combine-category-excel-files`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          category_name: categoryName,
          files: filesToCombine,
          fiscal_year: fiscalInfo.fiscal_year_str || '',
          region_mt_columns: regionMtColumns,
          region_value_columns: regionValueColumns,
          combination_type: 'category_based',
          excel_formatting: {
            number_format: {
              decimal_places: 2,
              apply_to_all_numbers: true,
              format_pattern: '0.00',
              force_format: true
            },
            column_formatting: {
              apply_column_types: true,
              number_columns_format: '0.00',
              percentage_columns_format: '0.00',
              currency_columns_format: '0.00'
            },
            title_formatting: {
              center_titles: true,
              title_font_size: 14,
              title_font_bold: true,
              title_background_color: '#E3F2FD',
              title_font_color: '#1565C0',
              merge_title_cells: true,
              add_borders: true
            }
          }
        })
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const result = await response.json();
      
      if (result.success) {
        // Convert base64 back to blob
        const byteCharacters = atob(result.file_data);
        const byteNumbers = new Array(byteCharacters.length);
        for (let i = 0; i < byteCharacters.length; i++) {
          byteNumbers[i] = byteCharacters.charCodeAt(i);
        }
        const byteArray = new Uint8Array(byteNumbers);
        const blob = new Blob([byteArray], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
        
        // Generate filename
        const timestamp = new Date().toISOString().slice(0, 19).replace(/[:-]/g, '');
        const fileName = result.file_name || `${categoryName.replace('_', '-')}_combined_${fiscalInfo.fiscal_year_str || 'report'}_${timestamp}.xlsx`;
        
        // Download combined file
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        // Add to stored files
        const newFile = {
          id: Date.now(),
          name: fileName,
          blob: blob,
          size: blob.size,
          createdAt: new Date().toISOString(),
          fiscalYear: fiscalInfo.fiscal_year_str || 'N/A',
          type: `${categoryName}-combined`,
          source: `Combined ${categoryName.replace('_', ' ')} Category`,
          url: URL.createObjectURL(blob),
          includedFiles: categoryFiles.length,
          sheets: result.metadata?.sheets_created || [],
          description: `Combined Excel from ${categoryFiles.length} ${categoryName.replace('_', ' ')} files with ${result.metadata?.sheets_created?.length || 0} sheets (Centered titles + 2-decimal formatting)`,
          metadata: {
            ...result.metadata,
            category: categoryName,
            source_files: categoryFiles.map(f => ({
              name: f.name,
              type: f.type,
              size: f.size
            })),
            region_mt_columns: regionMtColumns,
            region_value_columns: regionValueColumns,
            combination_type: 'category_based',
            formattingApplied: true,
            decimalPlaces: 2,
            numberFormat: '0.00',
            centerTitles: true,
            titleFormatting: {
              centered: true,
              bold: true,
              fontSize: 14,
              backgroundColor: '#E3F2FD',
              fontColor: '#1565C0'
            }
          }
        };
        
        setStoredFiles(prev => [newFile, ...prev]);
        
        addMessage(`✅ Successfully combined and downloaded ${categoryFiles.length} files from ${categoryName.replace('_', ' ')} category with centered titles and formatting`, 'success');
        addMessage(`📄 Created ${result.metadata?.sheets_created?.length || 0} sheets with centered titles`, 'info');
        
        if (result.metadata?.sheets_created?.length > 0) {
          addMessage(`📋 Sheets: ${result.metadata.sheets_created.join(', ')}`, 'info');
        }
      } else {
        throw new Error(result.error || 'Failed to combine category files');
      }
    } catch (error) {
      addMessage(`❌ Failed to combine ${categoryName.replace('_', ' ')} category files: ${error.message}`, 'error');
    } finally {
      setProcessing(false);
      setCombiningCategory(null);
    }
  }, [fiscalInfo, addMessage, API_BASE_URL, regionMtColumns, regionValueColumns, setStoredFiles]);

  // Get category statistics
  const getCategoryStats = useCallback((categoryFiles) => {
    const totalSheets = categoryFiles.reduce((sum, file) => 
      sum + (file.sheets?.length || 1), 0
    );
    const totalSize = categoryFiles.reduce((sum, file) => sum + file.size, 0);
    const totalRecords = categoryFiles.reduce((sum, file) => 
      sum + (file.mtRecords || 0) + (file.valueRecords || 0), 0
    );

    return { totalSheets, totalSize, totalRecords };
  }, []);

  // Enhanced Combine all Excel files into one master file with centered titles
  const combineAndDownloadExcel = useCallback(async () => {
    setProcessing(true);
    try {
      // Get all category files
      const categories = categorizeFiles(storedFiles);
      const filesToCombine = [];
      
      // Prepare files for combining
      for (const [category, files] of Object.entries(categories)) {
        if (files.length > 0) {
          const mostRecentFile = files.reduce((prev, current) => 
            (new Date(prev.createdAt) > new Date(current.createdAt)) ? prev : current
          );
          
          if (mostRecentFile.blob) {
            const arrayBuffer = await mostRecentFile.blob.arrayBuffer();
            const base64String = btoa(
              new Uint8Array(arrayBuffer).reduce(
                (data, byte) => data + String.fromCharCode(byte), ''
              )
            );
            
            filesToCombine.push({
              name: category,
              content: base64String,
              metadata: {
                ...mostRecentFile.metadata,
                decimalPlaces: mostRecentFile.metadata?.decimalPlaces || 2,
                category: category,
                sourceFile: mostRecentFile.name,
                centerTitles: true
              }
            });
          }
        }
      }

      if (filesToCombine.length === 0) {
        throw new Error('No files available to combine');
      }

      addMessage(`🔄 Combining all category files with centered titles and comprehensive formatting...`, 'info');
      
      const response = await fetch(`${API_BASE_URL}/combined/combine-excel-files`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          files: filesToCombine,
          region_mt_columns: regionMtColumns,
          region_value_columns: regionValueColumns,
          excel_formatting: {
            number_format: {
              decimal_places: 2,
              apply_to_all_numbers: true,
              format_pattern: '0.00',
              force_format: true,
              apply_to_sheets: 'all'
            },
            column_formatting: {
              apply_column_types: true,
              number_columns_format: '0.00',
              percentage_columns_format: '0.00',
              currency_columns_format: '0.00'
            },
            title_formatting: {
              center_all_titles: true,
              title_font_size: 16,
              title_font_bold: true,
              title_background_color: '#E3F2FD',
              title_font_color: '#1565C0',
              merge_title_cells: true,
              add_title_borders: true,
              title_row_height: 30,
              apply_to_all_sheets: true,
              sheet_title_style: {
                font_size: 18,
                font_bold: true,
                font_color: '#1565C0',
                background_color: '#F3E5F5',
                center_horizontally: true,
                center_vertically: true,
                add_border: true,
                border_color: '#9C27B0'
              }
            },
            header_formatting: {
              center_headers: true,
              header_font_bold: true,
              header_background_color: '#F5F5F5',
              header_font_color: '#333333',
              add_header_borders: true,
              header_row_height: 25
            }
          }
        })
      });

      const result = await response.json();
      
      if (result.success) {
        // Convert base64 back to blob
        const byteCharacters = atob(result.file_data);
        const byteNumbers = new Array(byteCharacters.length);
        for (let i = 0; i < byteCharacters.length; i++) {
          byteNumbers[i] = byteCharacters.charCodeAt(i);
        }
        const byteArray = new Uint8Array(byteNumbers);
        const blob = new Blob([byteArray], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
        
        // Download combined file
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = result.file_name || `combined_centered_titles_report_${new Date().toISOString().slice(0,10)}.xlsx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        // Add to stored files
        const newFile = {
          id: Date.now(),
          name: result.file_name,
          blob: blob,
          size: blob.size,
          createdAt: new Date().toISOString(),
          type: 'master-combined-centered-titles',
          source: 'Master Combined from all categories with centered titles and comprehensive formatting',
          url: URL.createObjectURL(blob),
          metadata: {
            ...result.metadata,
            region_mt_columns: regionMtColumns,
            region_value_columns: regionValueColumns,
            formattingApplied: true,
            decimalPlaces: 2,
            numberFormat: '0.00',
            totalSourceFiles: filesToCombine.length,
            categoriesIncluded: Object.keys(categories).filter(key => categories[key].length > 0),
            hasOverviewSheet: true,
            overviewSheetName: 'Master_Overview',
            generationType: 'master_combined_centered_titles',
            titleFormatting: {
              allTitlesCentered: true,
              titleFontSize: 16,
              titleFontBold: true,
              titleBackgroundColor: '#E3F2FD',
              titleFontColor: '#1565C0',
              mergeTitleCells: true,
              addTitleBorders: true,
              titleRowHeight: 30
            },
            headerFormatting: {
              headersCentered: true,
              headerFontBold: true,
              headerBackgroundColor: '#F5F5F5',
              headerFontColor: '#333333',
              addHeaderBorders: true,
              headerRowHeight: 25
            }
          }
        };
        
        setStoredFiles(prev => [newFile, ...prev]);
        addMessage('✅ Successfully combined all category files with centered titles and comprehensive formatting', 'success');
        addMessage(`📊 Created ${result.metadata?.total_sheets || 'multiple'} sheets with centered titles`, 'success');
        
        if (result.metadata?.category_breakdown) {
          const categories = Object.keys(result.metadata.category_breakdown);
          addMessage(`📂 Categories combined: ${categories.join(', ')}`, 'info');
        }
        
        // Additional success message about formatting
        addMessage('🎨 Applied formatting: Centered titles, bold headers, colored backgrounds, and 2-decimal number format', 'info');
      } else {
        throw new Error(result.error || 'Failed to combine files with centered titles');
      }
    } catch (error) {
      addMessage(`❌ Error combining files with centered titles: ${error.message}`, 'error');
    } finally {
      setProcessing(false);
    }
  }, [storedFiles, addMessage, setStoredFiles, API_BASE_URL, categorizeFiles, regionMtColumns, regionValueColumns]);

  // File management functions
  const downloadStoredFile = useCallback((fileData) => {
    const a = document.createElement('a');
    a.href = fileData.url;
    a.download = fileData.name;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    
    addMessage(`📁 Downloaded: ${fileData.name}`, 'success');
  }, [addMessage]);

  const deleteStoredFile = useCallback((fileId) => {
    const fileToDelete = storedFiles.find(file => file.id === fileId);
    if (fileToDelete && fileToDelete.url) {
      URL.revokeObjectURL(fileToDelete.url);
    }
    
    setStoredFiles(prev => prev.filter(file => file.id !== fileId));
    setSelectedFiles(prev => prev.filter(id => id !== fileId));
    
    addMessage('🗑️ File removed from storage', 'info');
  }, [storedFiles, setStoredFiles, addMessage]);

  const previewFile = useCallback((fileData) => {
    setFilePreview(fileData);
  }, []);

  const toggleFileSelection = useCallback((fileId) => {
    setSelectedFiles(prev => 
      prev.includes(fileId) 
        ? prev.filter(id => id !== fileId)
        : [...prev, fileId]
    );
  }, []);

  const selectAllFiles = useCallback(() => {
    setSelectedFiles(storedFiles.map(f => f.id));
  }, [storedFiles]);

  const clearSelection = useCallback(() => {
    setSelectedFiles([]);
  }, []);

  return (
    <div className="combined-excel-manager">
      {/* Header Section */}
      <div className="manager-header">
        <h3>🔗 Auditor File</h3>
        <p>Combine and manage Excel files with centered titles and professional formatting</p>
      </div>

      {/* Category Overview */}
      {categorySummary.length > 0 && (
        <div className="category-overview">
          <h4>📂 File Categories Overview</h4>
          <div className="category-grid">
            {categorySummary.map(category => (
              <div key={category.name} className={`category-card ${category.name}`}>
                <div className="category-icon">{category.icon}</div>
                <div className="category-content">
                  <span className="category-name">{category.name.replace('_', ' ').toUpperCase()}</span>
                  <span className="category-count">{category.count} files</span>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Action Groups */}
      <div className="manager-actions enhanced">
        <div className="action-group primary">
          <button
            onClick={combineAndDownloadExcel}
            className="btn btn-primary btn-large"
            disabled={loading || processing || storedFiles.length === 0}
          >
            {processing ? <RefreshCw size={16} className="spinning" /> : <Merge size={16} />}
            {processing ? 'Combining with Centered Titles...' : 'Auditor Format File'}
          </button>
          <p className="action-description">
            🎨 Combine all category Excel files into one master file with centered titles, professional formatting, and separate sheets
          </p>
        </div>

        {storedFiles.length > 0 && (
          <div className="action-group secondary">
            <button
              onClick={onFilesClear}
              className="btn btn-danger"
              disabled={loading || processing}
            >
              <Trash2 size={16} />
              Clear All Files ({storedFiles.length})
            </button>
          </div>
        )}
      </div>

      {/* Enhanced Files Display Section */}
      {storedFiles.length === 0 ? (
        <div className="empty-storage enhanced">
          <FileSpreadsheet size={48} />
          <h4>No Files Stored Yet</h4>
          <p>Generate files from analysis modules to see them organized by categories</p>
        </div>
      ) : (
        <div className="stored-files-section enhanced">
          {/* Enhanced Category Files Display */}
          <div className="enhanced-category-files-display">
            {Object.entries(fileGroups).map(([categoryName, categoryFiles]) => {
              if (categoryFiles.length === 0) return null;
              
              const categoryConfig = {
                sales: { 
                  icon: TrendingUp, 
                  title: '📊 Sales Analysis Files', 
                  color: 'sales', 
                  description: 'Sales performance and analysis files with month-wise data'
                },
                region: { 
                  icon: MapPin, 
                  title: '🌍 Region Analysis Files', 
                  color: 'region',
                  description: 'Regional performance data with MT and Value metrics'
                },
                product: { 
                  icon: Package, 
                  title: '📦 Product Analysis Files', 
                  color: 'product',
                  description: 'Product-wise performance and analysis data'
                },
                ts_pw: { 
                  icon: Activity, 
                  title: '⚡ TS-PW Analysis Files', 
                  color: 'ts-pw',
                  description: 'TS Product-wise analysis and metrics'
                },
                ero_pw: { 
                  icon: BarChart3, 
                  title: '📈 ERO-PW Analysis Files', 
                  color: 'ero-pw',
                  description: 'ERO Product-wise analysis for western region'
                },
                combined: { 
                  icon: Merge, 
                  title: '🔗 Combined Analysis Files', 
                  color: 'combined',
                  description: 'Previously combined analysis files from multiple sources'
                }
              };

              const config = categoryConfig[categoryName];
              if (!config) return null;

              const stats = getCategoryStats(categoryFiles);
              const isCurrentlyCombining = combiningCategory === categoryName;

              return (
                <div key={categoryName} className={`enhanced-file-category ${config.color}`}>
                  {/* Enhanced Category Header */}
                  <div className="enhanced-category-header">
                    <div className="category-title-section">
                      <div className="category-icon-wrapper">
                        <config.icon size={24} />
                      </div>
                      <div className="category-info">
                        <h5 className="category-title">{config.title}</h5>
                        <p className="category-description">{config.description}</p>
                        <div className="category-metrics">
                          <span className="metric">
                            <File size={14} />
                            {categoryFiles.length} files
                          </span>
                          <span className="metric">
                            <Layers size={14} />
                            {stats.totalSheets} sheets
                          </span>
                          <span className="metric">
                            <Database size={14} />
                            {formatFileSize(stats.totalSize)}
                          </span>
                          {stats.totalRecords > 0 && (
                            <span className="metric">
                              <BarChart3 size={14} />
                              {stats.totalRecords.toLocaleString()} records
                            </span>
                          )}
                        </div>
                      </div>
                    </div>

                    <div className="category-actions">
                      <button 
                        onClick={() => downloadCategoryExcel(categoryName, categoryFiles)}
                        className={`btn btn-primary btn-category-download ${isCurrentlyCombining ? 'combining' : ''}`}
                        disabled={loading || processing || categoryFiles.length === 0 || isCurrentlyCombining}
                        title={categoryFiles.length === 1 ? 
                          'Download single file' : 
                          `Combine all ${categoryFiles.length} files into one Excel with centered titles and ${stats.totalSheets} sheets`
                        }
                      >
                        {isCurrentlyCombining ? (
                          <>
                            <RefreshCw size={16} className="spinning" />
                            Combining with Titles...
                          </>
                        ) : (
                          <>
                            {categoryFiles.length === 1 ? (
                              <>
                                <Download size={16} />
                                Download File
                              </>
                            ) : (
                              <>
                                <Merge size={16} />
                                Combine All ({categoryFiles.length}) + Center Titles
                              </>
                            )}
                          </>
                        )}
                      </button>

                      <div className="files-count-display">
                        <span className="count-number">{categoryFiles.length}</span>
                        <span className="count-label">file{categoryFiles.length !== 1 ? 's' : ''}</span>
                      </div>
                    </div>
                  </div>

                  {/* Enhanced Combination Preview for multiple files */}
                  {categoryFiles.length > 1 && (
                    <div className="combination-preview">
                      <div className="preview-header">
                        <Merge size={16} />
                        <span>Combination Preview with Centered Titles</span>
                      </div>
                      <div className="preview-content">
                        <div className="preview-stats">
                          <div className="stat-item">
                            <span className="stat-value">{categoryFiles.length}</span>
                            <span className="stat-label">Source Files</span>
                          </div>
                          <div className="stat-item">
                            <span className="stat-value">{stats.totalSheets}</span>
                            <span className="stat-label">Total Sheets</span>
                          </div>
                          <div className="stat-item">
                            <span className="stat-value">1</span>
                            <span className="stat-label">Combined File</span>
                          </div>
                          <div className="stat-item formatting-preview">
                            <span className="stat-value">🎨</span>
                            <span className="stat-label">Centered Titles</span>
                          </div>
                        </div>
                        <div className="preview-description">
                          <p>All sheets from {categoryFiles.length} files will be combined into one Excel file with professionally centered titles, bold headers, colored backgrounds, and comprehensive formatting.</p>
                          <div className="formatting-features">
                            <span className="feature-badge">📐 Centered Titles</span>
                            <span className="feature-badge">🔢 2-Decimal Format</span>
                            <span className="feature-badge">🎨 Colored Headers</span>
                            <span className="feature-badge">📊 Professional Layout</span>
                          </div>
                        </div>
                      </div>
                    </div>
                  )}

                  {/* Files Grid */}
                  <div className="enhanced-files-grid">
                    {categoryFiles.map((file, index) => {
                      const uniqueFile = ensureUniqueId(file, index, categoryName);
                      return (
                        <FileCard 
                          key={uniqueFile.uniqueKey}
                          file={file} 
                          isSelected={selectedFiles.includes(file.id)}
                          onToggleSelect={toggleFileSelection}
                          onDownload={downloadStoredFile}
                          onPreview={previewFile}
                          onDelete={deleteStoredFile}
                          formatFileSize={formatFileSize}
                          formatDate={formatDate}
                          getFileTypeIcon={getFileTypeIcon}
                          categoryType={config.color}
                          regionMtColumns={regionMtColumns}
                          regionValueColumns={regionValueColumns}
                        />
                      );
                    })}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* File Preview Modal */}
      {filePreview && (
        <FilePreviewModal 
          file={filePreview}
          onClose={() => setFilePreview(null)}
          onDownload={downloadStoredFile}
          formatFileSize={formatFileSize}
          formatDate={formatDate}
          regionMtColumns={regionMtColumns}
          regionValueColumns={regionValueColumns}
        />
      )}

      {/* Enhanced Styles */}
      <style jsx>{`
        .combined-excel-manager {
          background: white;
          border-radius: 12px;
          padding: 24px;
          box-shadow: 0 4px 12px rgba(0,0,0,0.1);
          border: 1px solid #e0e0e0;
        }

        .manager-header {
          text-align: center;
          margin-bottom: 24px;
          padding-bottom: 20px;
          border-bottom: 2px solid #f0f0f0;
        }

        .manager-header h3 {
          margin: 0 0 8px 0;
          color: #333;
          font-size: 22px;
          font-weight: 700;
        }

        .manager-header p {
          margin: 0 0 16px 0;
          color: #666;
          font-size: 15px;
          line-height: 1.5;
        }

        .category-overview {
          background: linear-gradient(135deg, #f0f8ff 0%, #e6f3ff 100%);
          border-radius: 8px;
          padding: 20px;
          margin-bottom: 24px;
          border: 1px solid #2196f3;
        }

        .category-overview h4 {
          margin: 0 0 16px 0;
          color: #1565c0;
          font-size: 16px;
        }

        .category-grid {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
          gap: 12px;
        }

        .category-card {
          background: white;
          border-radius: 8px;
          padding: 16px;
          display: flex;
          align-items: center;
          gap: 12px;
          border: 1px solid #e9ecef;
          transition: all 0.2s ease;
        }

        .category-card:hover {
          box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        }

        .category-card.sales {
          border-left: 4px solid #28a745;
        }

        .category-card.region {
          border-left: 4px solid #007bff;
        }

        .category-card.product {
          border-left: 4px solid #dc3545;
        }

        .category-card.ts_pw {
          border-left: 4px solid #6c757d;
        }

        .category-card.ero_pw {
          border-left: 4px solid #17a2b8;
        }

        .category-card.combined {
          border-left: 4px solid #ffc107;
        }

        .category-icon {
          font-size: 24px;
          width: 40px;
          text-align: center;
        }

        .category-content {
          flex: 1;
        }

        .category-name {
          display: block;
          font-weight: 600;
          color: #333;
          font-size: 13px;
        }

        .category-count {
          display: block;
          color: #666;
          font-size: 12px;
        }

        .manager-actions.enhanced {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(320px, 1fr));
          gap: 24px;
          margin-bottom: 24px;
        }

        .action-group {
          background: #f8f9fa;
          border-radius: 8px;
          padding: 20px;
          border: 1px solid #e9ecef;
        }

        .action-group.primary {
          border-left: 4px solid #007bff;
          background: linear-gradient(135deg, #f0f8ff 0%, #e6f3ff 100%);
        }

        .action-group.secondary {
          border-left: 4px solid #6c757d;
        }

        .action-buttons {
          display: flex;
          gap: 12px;
          flex-wrap: wrap;
        }

        .action-description {
          margin: 8px 0 0;
          font-size: 13px;
          color: #666;
          font-style: italic;
        }

        .btn {
          display: inline-flex;
          align-items: center;
          gap: 8px;
          padding: 10px 16px;
          border: none;
          border-radius: 6px;
          font-size: 14px;
          font-weight: 500;
          cursor: pointer;
          transition: all 0.2s ease;
          text-decoration: none;
          text-align: center;
          white-space: nowrap;
          position: relative;
        }

        .btn:disabled {
          opacity: 0.6;
          cursor: not-allowed;
          transform: none;
        }

        .btn:hover:not(:disabled) {
          transform: translateY(-1px);
          box-shadow: 0 4px 8px rgba(0,0,0,0.15);
        }

        .btn-small {
          padding: 6px 12px;
          font-size: 12px;
        }

        .btn-large {
          padding: 14px 28px;
          font-size: 16px;
          font-weight: 600;
        }

        .btn-primary {
          background: linear-gradient(135deg, #007bff, #0056b3);
          color: white;
          box-shadow: 0 2px 4px rgba(0,123,255,0.3);
        }

        .btn-secondary {
          background: linear-gradient(135deg, #6c757d, #5a6268);
          color: white;
          box-shadow: 0 2px 4px rgba(108,117,125,0.3);
        }

        .btn-danger {
          background: linear-gradient(135deg, #dc3545, #c82333);
          color: white;
          box-shadow: 0 2px 4px rgba(220,53,69,0.3);
        }

        .empty-storage.enhanced {
          text-align: center;
          padding: 60px 20px;
          color: #666;
          border: 2px dashed #dee2e6;
          border-radius: 12px;
          background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
        }

        .empty-storage.enhanced svg {
          color: #dee2e6;
          margin-bottom: 20px;
        }

        .empty-storage.enhanced h4 {
          margin: 0 0 12px 0;
          color: #495057;
          font-size: 18px;
        }

        .empty-storage.enhanced p {
          margin: 0 0 16px 0;
          font-size: 14px;
        }

        /* Enhanced Category Files Display Styles */
        .enhanced-category-files-display {
          display: flex;
          flex-direction: column;
          gap: 32px;
        }

        .enhanced-file-category {
          border: 1px solid #e9ecef;
          border-radius: 12px;
          overflow: hidden;
          background: white;
          box-shadow: 0 2px 8px rgba(0,0,0,0.05);
          transition: all 0.3s ease;
        }

        .enhanced-file-category:hover {
          box-shadow: 0 4px 16px rgba(0,0,0,0.1);
        }

        .enhanced-file-category.sales {
          border-left: 4px solid #28a745;
        }

        .enhanced-file-category.region {
          border-left: 4px solid #007bff;
        }

        .enhanced-file-category.combined {
          border-left: 4px solid #ffc107;
        }

        .enhanced-file-category.product {
          border-left: 4px solid #dc3545;
        }

        .enhanced-file-category.ts-pw {
          border-left: 4px solid #6c757d;
        }

        .enhanced-file-category.ero-pw {
          border-left: 4px solid #17a2b8;
        }

        .enhanced-category-header {
          display: flex;
          justify-content: space-between;
          align-items: flex-start;
          padding: 24px;
          background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
          border-bottom: 1px solid #dee2e6;
          gap: 20px;
        }

        .category-title-section {
          display: flex;
          gap: 16px;
          flex: 1;
          min-width: 0;
        }

        .category-icon-wrapper {
          background: white;
          padding: 12px;
          border-radius: 12px;
          box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          color: #495057;
          display: flex;
          align-items: center;
          justify-content: center;
          min-width: 48px;
          height: 48px;
        }

        .category-info {
          flex: 1;
          min-width: 0;
        }

        .category-title {
          margin: 0 0 4px 0;
          color: #333;
          font-size: 18px;
          font-weight: 700;
          line-height: 1.3;
        }

        .category-description {
          margin: 0 0 12px 0;
          color: #666;
          font-size: 14px;
          line-height: 1.4;
        }

        .category-metrics {
          display: flex;
          gap: 16px;
          flex-wrap: wrap;
        }

        .metric {
          display: flex;
          align-items: center;
          gap: 6px;
          font-size: 12px;
          color: #495057;
          font-weight: 500;
          background: white;
          padding: 4px 8px;
          border-radius: 6px;
          box-shadow: 0 1px 2px rgba(0,0,0,0.05);
        }

        .category-actions {
          display: flex;
          align-items: center;
          gap: 16px;
          flex-wrap: wrap;
        }

        .btn-category-download {
          position: relative;
          padding: 12px 24px;
          font-size: 14px;
          font-weight: 600;
          border-radius: 8px;
          transition: all 0.3s ease;
          min-width: 200px;
          box-shadow: 0 2px 4px rgba(0,123,255,0.3);
        }

        .btn-category-download.combining {
          background: linear-gradient(135deg, #6c757d, #5a6268);
          cursor: not-allowed;
        }

        .btn-category-download:hover:not(:disabled):not(.combining) {
          transform: translateY(-2px);
          box-shadow: 0 4px 12px rgba(0,123,255,0.4);
        }

        .files-count-display {
          display: flex;
          flex-direction: column;
          align-items: center;
          gap: 2px;
          min-width: 60px;
        }

        .count-number {
          font-size: 24px;
          font-weight: 700;
          color: #495057;
          line-height: 1;
        }

        .count-label {
          font-size: 11px;
          color: #6c757d;
          text-transform: uppercase;
          font-weight: 600;
        }

        .combination-preview {
          background: linear-gradient(135deg, #e8f5e8 0%, #c8e6c9 100%);
          border-bottom: 1px solid #4caf50;
          padding: 20px 24px;
        }

        .preview-header {
          display: flex;
          align-items: center;
          gap: 8px;
          margin-bottom: 16px;
          color: #2e7d32;
          font-weight: 600;
          font-size: 14px;
        }

        .preview-content {
          display: flex;
          gap: 24px;
          align-items: flex-start;
          flex-wrap: wrap;
        }

        .preview-stats {
          display: flex;
          gap: 16px;
          flex-wrap: wrap;
        }

        .stat-item {
          display: flex;
          flex-direction: column;
          align-items: center;
          gap: 4px;
          background: white;
          padding: 12px 16px;
          border-radius: 8px;
          box-shadow: 0 2px 4px rgba(0,0,0,0.05);
          min-width: 80px;
        }

        .stat-item.formatting-preview {
          background: linear-gradient(135deg, #fff3e0 0%, #ffe0b2 100%);
          border: 1px solid #ff9800;
        }

        .stat-value {
          font-size: 20px;
          font-weight: 700;
          color: #2e7d32;
          line-height: 1;
        }

        .stat-label {
          font-size: 11px;
          color: #455a64;
          text-transform: uppercase;
          font-weight: 600;
          text-align: center;
        }

        .preview-description {
          flex: 1;
          min-width: 200px;
        }

        .preview-description p {
          margin: 0 0 12px 0;
          color: #2e7d32;
          font-size: 13px;
          line-height: 1.4;
          font-style: italic;
        }

        .formatting-features {
          display: flex;
          flex-wrap: wrap;
          gap: 6px;
        }

        .feature-badge {
          background: white;
          color: #2e7d32;
          padding: 4px 8px;
          border-radius: 12px;
          font-size: 11px;
          font-weight: 600;
          border: 1px solid #4caf50;
        }

        .enhanced-files-grid {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(420px, 1fr));
          gap: 20px;
          padding: 24px;
        }

        .spinning {
          animation: spin 1s linear infinite;
        }

        @keyframes spin {
          from { transform: rotate(0deg); }
          to { transform: rotate(360deg); }
        }

        @media (max-width: 1200px) {
          .enhanced-files-grid {
            grid-template-columns: repeat(auto-fill, minmax(380px, 1fr));
          }
        }

        @media (max-width: 768px) {
          .combined-excel-manager {
            padding: 16px;
          }

          .manager-actions.enhanced {
            grid-template-columns: 1fr;
            gap: 16px;
          }

          .action-buttons {
            flex-direction: column;
            align-items: stretch;
          }

          .enhanced-category-header {
            flex-direction: column;
            align-items: stretch;
            gap: 16px;
          }

          .category-title-section {
            flex-direction: column;
            align-items: center;
            text-align: center;
            gap: 12px;
          }

          .category-metrics {
            justify-content: center;
          }

          .category-actions {
            justify-content: center;
            flex-direction: column;
            gap: 12px;
          }

          .preview-content {
            flex-direction: column;
            align-items: stretch;
            gap: 16px;
          }

          .preview-stats {
            justify-content: center;
          }

          .enhanced-files-grid {
            grid-template-columns: 1fr;
            gap: 16px;
            padding: 16px;
          }

          .btn-category-download {
            min-width: auto;
            width: 100%;
          }

          .category-grid {
            grid-template-columns: 1fr;
          }
        }

        @media (max-width: 480px) {
          .enhanced-category-header {
            padding: 16px;
          }

          .combination-preview {
            padding: 16px;
          }

          .category-metrics {
            gap: 8px;
          }

          .metric {
            font-size: 11px;
            padding: 3px 6px;
          }

          .preview-stats {
            gap: 12px;
          }

          .stat-item {
            padding: 8px 12px;
            min-width: 70px;
          }

          .stat-value {
            font-size: 18px;
          }
        }
      `}</style>
    </div>
  );
};

// File Card Component
const FileCard = ({ 
  file, 
  isSelected, 
  onToggleSelect, 
  onDownload, 
  onPreview, 
  onDelete, 
  formatFileSize, 
  formatDate, 
  getFileTypeIcon,
  categoryType = 'default',
  regionMtColumns,
  regionValueColumns
}) => (
  <div className={`enhanced-file-card ${isSelected ? 'selected' : ''} ${categoryType}`}>
    <div className="file-header">
      <div className="file-selection">
        <input
          type="checkbox"
          checked={isSelected}
          onChange={() => onToggleSelect(file.id)}
          className="file-checkbox"
        />
      </div>
      <div className="file-icon">
        {getFileTypeIcon(file.type)}
        <span className="file-type-badge">{file.type}</span>
      </div>
      <div className="file-info">
        <h5 className="file-name">{file.name}</h5>
        <div className="file-meta">
          <span>{formatFileSize(file.size)}</span>
          <span>•</span>
          <span>{formatDate(file.createdAt || file.timestamp)}</span>
        </div>
      </div>
    </div>

    <div className="file-details">
      <div className="detail-section">
        <div className="detail-row">
          <span className="detail-label">Source:</span>
          <span className="detail-value">{file.source}</span>
        </div>
        <div className="detail-row">
          <span className="detail-label">Fiscal Year:</span>
          <span className="detail-value">{file.fiscalYear}</span>
        </div>
        {(file.mtRecords > 0 || file.metadata?.mtRows > 0) && (
          <div className="detail-row">
            <span className="detail-label">MT Records:</span>
            <span className="detail-value">{(file.mtRecords || file.metadata?.mtRows || 0).toLocaleString()}</span>
          </div>
        )}
        {(file.valueRecords > 0 || file.metadata?.valueRows > 0) && (
          <div className="detail-row">
            <span className="detail-label">Value Records:</span>
            <span className="detail-value">{(file.valueRecords || file.metadata?.valueRows || 0).toLocaleString()}</span>
          </div>
        )}
        
        {/* Show formatting information if available */}
        {file.metadata?.centerTitles && (
          <div className="detail-row formatting-info">
            <span className="detail-label">Formatting:</span>
            <span className="detail-value formatting-badge">🎨 Centered Titles Applied</span>
          </div>
        )}
      </div>
      
      {(file.type.includes('region') && (file.type.includes('mt') || file.type.includes('value'))) && (
        <div className="detail-section">
          <div className="columns-list">
            {(file.type.includes('mt') ? regionMtColumns : regionValueColumns).map((col, index) => (
              <span key={index} className="column-tag">{col.header}</span>
            ))}
          </div>
        </div>
      )}

      {file.sheets?.length > 0 && (
        <div className="detail-section">
          <div className="sheets-list">
            {file.sheets.map((sheet, index) => (
              <span key={index} className="sheet-tag">{sheet}</span>
            ))}
          </div>
        </div>
      )}

      {file.description && (
        <div className="detail-section">
          <p className="file-description">{file.description}</p>
        </div>
      )}
    </div>

    <div className="file-actions">
      <button
        onClick={() => onDownload(file)}
        className="btn btn-primary btn-small"
        title="Download File"
      >
        <Download size={14} />
        Download
      </button>
      <button
        onClick={() => onPreview(file)}
        className="btn btn-info btn-small"
        title="View Details"
      >
        <Eye size={14} />
        Details
      </button>
      <button
        onClick={() => onDelete(file.id)}
        className="btn btn-danger btn-small"
        title="Delete File"
      >
        <Trash2 size={14} />
      </button>
    </div>

    <style jsx>{`
      .enhanced-file-card {
        border: 2px solid #e0e0e0;
        border-radius: 12px;
        padding: 20px;
        background: white;
        transition: all 0.3s ease;
        position: relative;
        overflow: hidden;
      }

      .enhanced-file-card::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        height: 4px;
        background: #e0e0e0;
        transition: background 0.3s ease;
      }

      .enhanced-file-card.sales::before {
        background: linear-gradient(90deg, #28a745, #20c997);
      }

      .enhanced-file-card.region::before {
        background: linear-gradient(90deg, #007bff, #6610f2);
      }

      .enhanced-file-card.combined::before {
        background: linear-gradient(90deg, #ffc107, #fd7e14);
      }

      .enhanced-file-card.product::before {
        background: linear-gradient(90deg, #dc3545, #e83e8c);
      }

      .enhanced-file-card.ts-pw::before {
        background: linear-gradient(90deg, #6c757d, #495057);
      }

      .enhanced-file-card.ero-pw::before {
        background: linear-gradient(90deg, #17a2b8, #6f42c1);
      }

      .enhanced-file-card:hover {
        box-shadow: 0 8px 16px rgba(0,0,0,0.1);
        border-color: #007bff;
        transform: translateY(-2px);
      }

      .enhanced-file-card.selected {
        border-color: #28a745;
        background: linear-gradient(135deg, #ffffff 0%, #f0fff4 100%);
        box-shadow: 0 4px 12px rgba(40,167,69,0.2);
      }

      .file-header {
        display: flex;
        align-items: flex-start;
        gap: 16px;
        margin-bottom: 16px;
      }

      .file-checkbox {
        width: 18px;
        height: 18px;
        cursor: pointer;
        margin-top: 8px;
      }

      .file-icon {
        background: #e3f2fd;
        padding: 12px;
        border-radius: 8px;
        color: #1976d2;
        position: relative;
      }

      .file-type-badge {
        position: absolute;
        top: -8px;
        right: -8px;
        background: #ff9800;
        color: white;
        font-size: 10px;
        padding: 2px 6px;
        border-radius: 8px;
        font-weight: 600;
        text-transform: uppercase;
      }

      .file-info {
        flex: 1;
      }

      .file-name {
        margin: 0 0 6px 0;
        font-size: 15px;
        font-weight: 600;
        color: #333;
        word-break: break-word;
        line-height: 1.3;
      }

      .file-meta {
        display: flex;
        gap: 8px;
        font-size: 12px;
        color: #666;
      }

      .file-details {
        margin-bottom: 16px;
        background: #f8f9fa;
        border-radius: 8px;
        padding: 16px;
      }

      .detail-section {
        margin-bottom: 16px;
      }

         .detail-section:last-child {
        margin-bottom: 0;
      }

      .detail-row {
        display: flex;
        justify-content: space-between;
        padding: 6px 0;
        font-size: 13px;
        border-bottom: 1px solid #f0f0f0;
      }

      .detail-row:last-child {
        border-bottom: none;
      }

      .detail-label {
        color: #666;
        font-weight: 500;
      }

      .detail-value {
        color: #333;
        font-weight: 600;
        text-align: right;
        max-width: 60%;
        word-break: break-word;
      }

      .columns-list {
        display: flex;
        flex-wrap: wrap;
        gap: 6px;
      }

      .column-tag {
        background: #e3f2fd;
        color: #1976d2;
        padding: 4px 8px;
        border-radius: 4px;
        font-size: 11px;
      }

      .sheets-list {
        display: flex;
        flex-wrap: wrap;
        gap: 6px;
      }

      .sheet-tag {
        background: #e3f2fd;
        color: #1976d2;
        padding: 4px 8px;
        border-radius: 12px;
        font-size: 11px;
        font-weight: 500;
      }

      .file-description {
        margin: 0;
        font-size: 13px;
        color: #666;
        line-height: 1.4;
      }

      .file-actions {
        display: flex;
        gap: 8px;
        flex-wrap: wrap;
      }

      .btn {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        padding: 8px 12px;
        border: none;
        border-radius: 6px;
        font-size: 12px;
        font-weight: 500;
        cursor: pointer;
        transition: all 0.2s ease;
      }

      .btn-primary {
        background: #007bff;
        color: white;
      }

      .btn-info {
        background: #17a2b8;
        color: white;
      }

      .btn-danger {
        background: #dc3545;
        color: white;
      }

      .btn:hover {
        transform: translateY(-1px);
        box-shadow: 0 2px 4px rgba(0,0,0,0.15);
      }

      @media (max-width: 480px) {
        .file-header {
          flex-direction: column;
          align-items: center;
          text-align: center;
          gap: 12px;
        }

        .file-info {
          text-align: center;
        }

        .detail-row {
          flex-direction: column;
          gap: 4px;
          text-align: center;
        }

        .detail-value {
          text-align: center;
          max-width: 100%;
        }

        .file-actions {
          justify-content: center;
        }
      }
    `}</style>
  </div>
);

// File Preview Modal Component
const FilePreviewModal = ({ 
  file, 
  onClose, 
  onDownload, 
  formatFileSize, 
  formatDate,
  regionMtColumns,
  regionValueColumns 
}) => (
  <div className="file-preview-modal">
    <div className="modal-backdrop" onClick={onClose} />
    <div className="modal-content">
      <div className="modal-header">
        <h4>📄 File Details: {file.name}</h4>
        <button onClick={onClose} className="modal-close">×</button>
      </div>
      
      <div className="modal-body">
        <div className="preview-section">
          <div className="info-grid">
            <div className="info-item">
              <span className="info-label">File Name:</span>
              <span className="info-value">{file.name}</span>
            </div>
            <div className="info-item">
              <span className="info-label">File Size:</span>
              <span className="info-value">{formatFileSize(file.size)}</span>
            </div>
            <div className="info-item">
              <span className="info-label">Created:</span>
              <span className="info-value">{formatDate(file.createdAt || file.timestamp)}</span>
            </div>
            <div className="info-item">
              <span className="info-label">Type:</span>
              <span className="info-value">{file.type}</span>
            </div>
            <div className="info-item">
              <span className="info-label">Source:</span>
              <span className="info-value">{file.source}</span>
            </div>
            <div className="info-item">
              <span className="info-label">Fiscal Year:</span>
              <span className="info-value">{file.fiscalYear}</span>
            </div>
          </div>
        </div>

        <div className="preview-section">
          <div className="data-summary">
            {(file.mtRecords > 0 || file.metadata?.mtRows > 0) && (
              <div className="summary-card">
                <div className="summary-icon mt-icon">MT</div>
                <div className="summary-content">
                  <span className="summary-number">{(file.mtRecords || file.metadata?.mtRows || 0).toLocaleString()}</span>
                  <span className="summary-label">MT Records</span>
                </div>
              </div>
            )}
            {(file.valueRecords > 0 || file.metadata?.valueRows > 0) && (
              <div className="summary-card">
                <div className="summary-icon value-icon">₹</div>
                <div className="summary-content">
                  <span className="summary-number">{(file.valueRecords || file.metadata?.valueRows || 0).toLocaleString()}</span>
                  <span className="summary-label">Value Records</span>
                </div>
              </div>
            )}
            {file.includedFiles && (
              <div className="summary-card">
                <div className="summary-icon files-icon">📁</div>
                <div className="summary-content">
                  <span className="summary-number">{file.includedFiles}</span>
                  <span className="summary-label">Included Files</span>
                </div>
              </div>
            )}
          </div>
        </div>

        {file.sheets?.length > 0 && (
          <div className="preview-section">
            <div className="sheets-preview">
              {file.sheets.map((sheet, index) => (
                <div key={index} className="sheet-preview-item">
                  <FileSpreadsheet size={16} />
                  <span>{sheet}</span>
                </div>
              ))}
            </div>
          </div>
        )}

        {file.description && (
          <div className="preview-section">
            <p className="file-description">{file.description}</p>
          </div>
        )}
      </div>

      <div className="modal-footer">
        <button onClick={() => onDownload(file)} className="btn btn-primary">
          <Download size={16} />
          Download File
        </button>
        <button onClick={onClose} className="btn btn-secondary">
          Close
        </button>
      </div>
    </div>

    <style jsx>{`
      .file-preview-modal {
        position: fixed;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        z-index: 1000;
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 20px;
      }

      .modal-backdrop {
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: rgba(0,0,0,0.6);
      }

      .modal-content {
        background: white;
        border-radius: 12px;
        width: 100%;
        max-width: 700px;
        position: relative;
        box-shadow: 0 20px 40px rgba(0,0,0,0.3);
        max-height: 90vh;
        overflow-y: auto;
      }

      .modal-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 24px;
        border-bottom: 1px solid #e0e0e0;
        background: white;
        border-radius: 12px 12px 0 0;
      }

      .modal-header h4 {
        margin: 0;
        color: #333;
        font-size: 18px;
        word-break: break-word;
      }

      .modal-close {
        background: none;
        border: none;
        font-size: 28px;
        cursor: pointer;
        color: #666;
        width: 36px;
        height: 36px;
        display: flex;
        align-items: center;
        justify-content: center;
        border-radius: 6px;
        transition: all 0.2s ease;
      }

      .modal-close:hover {
        background: #f0f0f0;
        color: #333;
      }

      .modal-body {
        padding: 24px;
      }

      .preview-section {
        margin-bottom: 24px;
      }

      .preview-section:last-child {
        margin-bottom: 0;
      }

      .info-grid {
        display: grid;
        gap: 12px;
      }

      .info-item {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 12px 16px;
        background: #f8f9fa;
        border-radius: 6px;
        border: 1px solid #e9ecef;
      }

      .info-label {
        font-weight: 500;
        color: #666;
        font-size: 14px;
      }

      .info-value {
        font-weight: 600;
        color: #333;
        font-size: 14px;
        text-align: right;
        max-width: 60%;
        word-break: break-word;
      }

      .data-summary {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 16px;
      }

      .summary-card {
        display: flex;
        align-items: center;
        gap: 16px;
        padding: 20px;
        background: #f8f9fa;
        border-radius: 8px;
        border-left: 4px solid #007bff;
      }

      .summary-icon {
        width: 45px;
        height: 45px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: bold;
        font-size: 16px;
        color: white;
      }

      .mt-icon {
        background: linear-gradient(135deg, #28a745, #20c997);
      }

      .value-icon {
        background: linear-gradient(135deg, #17a2b8, #6f42c1);
      }

      .files-icon {
        background: linear-gradient(135deg, #fd7e14, #e83e8c);
      }

      .summary-content {
        display: flex;
        flex-direction: column;
      }

      .summary-number {
        font-size: 20px;
        font-weight: bold;
        color: #333;
        line-height: 1.2;
      }

      .summary-label {
        font-size: 12px;
        color: #666;
        text-transform: uppercase;
        margin-top: 2px;
      }

      .sheets-preview {
        display: flex;
        flex-direction: column;
        gap: 8px;
      }

      .sheet-preview-item {
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 8px 12px;
        background: #f8f9fa;
        border-radius: 6px;
        border: 1px solid #e9ecef;
        font-size: 14px;
        color: #333;
      }

      .file-description {
        margin: 0;
        padding: 16px;
        background: #f8f9fa;
        border-radius: 8px;
        border: 1px solid #e9ecef;
        color: #666;
        line-height: 1.5;
      }

      .modal-footer {
        display: flex;
        gap: 12px;
        justify-content: flex-end;
        padding: 24px;
        border-top: 1px solid #e0e0e0;
        background: #f8f9fa;
        border-radius: 0 0 12px 12px;
      }

      .btn {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        padding: 10px 20px;
        border: none;
        border-radius: 6px;
        font-size: 14px;
        font-weight: 500;
        cursor: pointer;
        transition: all 0.2s ease;
      }

      .btn-primary {
        background: #007bff;
        color: white;
      }

      .btn-secondary {
        background: #6c757d;
        color: white;
      }

      .btn:hover {
        transform: translateY(-1px);
        box-shadow: 0 2px 4px rgba(0,0,0,0.15);
      }

      @media (max-width: 768px) {
        .modal-content {
          margin: 10px;
          max-height: 95vh;
        }

        .modal-footer {
          flex-direction: column-reverse;
        }

        .data-summary {
          grid-template-columns: 1fr;
        }

        .info-item {
          flex-direction: column;
          gap: 8px;
          text-align: center;
        }

        .info-value {
          text-align: center;
          max-width: 100%;
        }

        .modal-header h4 {
          font-size: 16px;
        }
      }
    `}</style>
  </div>
);

export default CombinedExcelManager;