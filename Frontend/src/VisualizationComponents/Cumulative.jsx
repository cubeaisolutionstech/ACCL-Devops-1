// Cumulative.js
import React, { useState } from 'react';

const months = [
  "April", "May", "June", "July", 
  "August", "September", "October", "November",
  "December", "January", "February", "March"
];

const Cumulative = () => {
  const [files, setFiles] = useState({});
  const [skipFirstRow, setSkipFirstRow] = useState(false);
  const [preview, setPreview] = useState(null);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState(null);
  const [previewGenerated, setPreviewGenerated] = useState(false);

  const handleFileChange = (month, e) => {
    if (e.target.files[0]) {
      const newFiles = { ...files, [month]: e.target.files[0] };
      setFiles(newFiles);
      setPreviewGenerated(false);
    }
  };

  const handleProcess = async () => {
    setLoading(true);
    setMessage(null);
    setPreview(null);

    try {
      const formData = new FormData();
      Object.entries(files).forEach(([month, file]) => {
        formData.append(month, file);
      });
      formData.append('skipFirstRow', skipFirstRow);

      const response = await fetch('http://127.0.0.1:5000/api/process', {
        method: 'POST',
        body: formData,
        credentials: 'include'
      });

      const rawText = await response.text();

      const safeJsonParse = (text) => {
        try {
          return JSON.parse(
            text
              .replace(/"Part No":\s*NaN/g, '"Part No": null')
              .replace(/:\s*NaN/g, ': null')
              .replace(/Infinity/g, 'null')
              .replace(/-Infinity/g, 'null')
          );
        } catch (e) {
          console.error('JSON parse error:', e);
          throw new Error(`Data format error: ${e.message}`);
        }
      };

      const data = safeJsonParse(rawText);
      if (!response.ok || !data.success) {
        throw new Error(data.message || 'Processing failed');
      }

      const deepCleanData = (obj) => {
        if (obj === null || obj === undefined) return '';
        if (typeof obj === 'number' && isNaN(obj)) return '';
        if (Array.isArray(obj)) return obj.map(deepCleanData);
        if (typeof obj === 'object') {
          return Object.fromEntries(Object.entries(obj).map(([k, v]) => [k, deepCleanData(v)]));
        }
        return obj;
      };

      const cleanPreview = Array.isArray(data.preview) 
        ? data.preview.map(deepCleanData) 
        : [];

      setPreview(cleanPreview);
      setMessage(data.message || 'Files processed successfully');
      setPreviewGenerated(true);

      if (data.warnings?.length > 0) {
        setMessage(prev => `${prev} (with ${data.warnings.length} warning${data.warnings.length > 1 ? 's' : ''})`);
      }

    } catch (err) {
      const cleanErrorMessage = err.message
        .replace(/NaN/g, '[Invalid Number]')
        .replace(/undefined/g, '[Missing Data]');
      setMessage(`Error: ${cleanErrorMessage}`);
      setPreviewGenerated(false);
    } finally {
      setLoading(false);
    }
  };

  const handleDownload = async () => {
    setLoading(true);
    setMessage(null);

    try {
      const formData = new FormData();
      Object.entries(files).forEach(([month, file]) => {
        formData.append(month, file);
      });
      formData.append('skipFirstRow', skipFirstRow);

      const response = await fetch('http://127.0.0.1:5000/api/download', {
        method: 'POST',
        body: formData,
        credentials: 'include'
      });

      const contentType = response.headers.get('content-type');

      if (!response.ok) {
        const errorText = contentType?.includes('application/json') 
          ? (await response.json()).message 
          : await response.text();
        throw new Error(errorText || 'Download failed');
      }

      if (!contentType?.includes('spreadsheetml.sheet')) {
        throw new Error('Server returned unexpected file format');
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'Combined_Sales_Report.xlsx';
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);

    } catch (err) {
      setMessage(err.message || 'An unknown error occurred during download');
    } finally {
      setLoading(false);
    }
  };

  const isProcessingDisabled = loading || Object.keys(files).length === 0;
  const isDownloadDisabled = loading || !previewGenerated;

  return (
    <div style={{ minHeight: '100vh', padding: '20px', backgroundColor: '#f8f9fa', fontFamily: 'Arial, sans-serif' }}>
      <div style={{ maxWidth: '1200px', margin: '0 auto', backgroundColor: 'white', padding: '20px', borderRadius: '8px', boxShadow: '0 2px 4px rgba(0,0,0,0.1)' }}>
        <h1 style={{ textAlign: 'center', marginBottom: '30px', color: '#343a40', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '10px' }}>
          <span style={{ fontSize: '1.2em' }}>📊</span> Monthly Sales Data Processor
        </h1>

        <div style={{ marginBottom: '25px', padding: '15px', backgroundColor: '#f1f3f5', borderRadius: '6px' }}>
          <label style={{ display: 'flex', alignItems: 'center' }}>
            <input
              type="checkbox"
              checked={skipFirstRow}
              onChange={() => setSkipFirstRow(!skipFirstRow)}
              style={{ marginRight: '10px', width: '18px', height: '18px' }}
            />
            <span>Skip first row (use second row as headers)</span>
          </label>
        </div>

        <div style={{ marginBottom: '25px', padding: '15px', border: '1px solid #dee2e6', borderRadius: '6px' }}>
          <h2 style={{ marginBottom: '15px' }}>Upload New Files</h2>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(200px, 1fr))', gap: '15px' }}>
            {months.map(month => (
              <div key={month} style={{ padding: '15px', border: `1px solid ${files[month] ? '#40c057' : '#adb5bd'}`, borderRadius: '6px', backgroundColor: files[month] ? '#ebfbee' : 'white' }}>
                <div style={{ fontWeight: 'bold', marginBottom: '10px', color: '#495057' }}>{month}</div>
                <input type="file" accept=".xlsx,.xls" onChange={(e) => handleFileChange(month, e)} style={{ display: 'none' }} id={`file-${month}`} />
                <label htmlFor={`file-${month}`} style={{ display: 'block', padding: '8px 12px', backgroundColor: '#228be6', color: 'white', borderRadius: '4px', cursor: 'pointer', textAlign: 'center', fontSize: '14px' }}>
                  {files[month] ? 'Change File' : 'Select File'}
                </label>
                {files[month] && <div style={{ marginTop: '8px', fontSize: '12px', color: '#495057', wordBreak: 'break-word' }}>{files[month].name}</div>}
              </div>
            ))}
          </div>
        </div>

        <div style={{ display: 'flex', justifyContent: 'center', gap: '15px', marginBottom: '25px', flexWrap: 'wrap' }}>
          <button onClick={handleProcess} disabled={isProcessingDisabled} style={{
            padding: '10px 20px',
            backgroundColor: isProcessingDisabled ? '#adb5bd' : '#228be6',
            color: 'white',
            border: 'none',
            borderRadius: '6px',
            cursor: isProcessingDisabled ? 'not-allowed' : 'pointer',
            fontSize: '16px',
            fontWeight: 'bold',
            minWidth: '180px'
          }}>
            {loading ? 'Processing...' : 'Process Files'}
          </button>

          <button onClick={handleDownload} disabled={isDownloadDisabled} style={{
            padding: '10px 20px',
            backgroundColor: isDownloadDisabled ? '#adb5bd' : '#40c057',
            color: 'white',
            border: 'none',
            borderRadius: '6px',
            cursor: isDownloadDisabled ? 'not-allowed' : 'pointer',
            fontSize: '16px',
            fontWeight: 'bold',
            minWidth: '180px'
          }}>
            {loading ? 'Preparing...' : 'Download Excel'}
          </button>
        </div>

        {message && (
          <div style={{
            padding: '15px',
            marginBottom: '20px',
            backgroundColor: message.includes('Error') ? '#fff3bf' : '#ebfbee',
            color: message.includes('Error') ? '#e67700' : '#2b8a3e',
            borderRadius: '6px',
            border: `1px solid ${message.includes('Error') ? '#ffd43b' : '#40c057'}`,
            wordBreak: 'break-word'
          }}>
            {message}
          </div>
        )}

        {preview && preview.length > 0 && (
          <div style={{ padding: '15px', border: '1px solid #dee2e6', borderRadius: '6px', marginBottom: '20px' }}>
            <h2 style={{ marginBottom: '15px', display: 'flex', alignItems: 'center', gap: '10px' }}>
              <span>📋</span> Data Preview (First 5 Rows)
            </h2>
            <div style={{ overflowX: 'auto', border: '1px solid #dee2e6', borderRadius: '4px' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '14px' }}>
                <thead>
                  <tr style={{ backgroundColor: '#f1f3f5' }}>
                    {Object.keys(preview[0]).map(header => (
                      <th key={header} style={{ padding: '10px', textAlign: 'left', border: '1px solid #dee2e6', fontWeight: '600' }}>
                        {header}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {preview.map((row, rowIndex) => (
                    <tr key={rowIndex} style={{ backgroundColor: rowIndex % 2 === 0 ? 'white' : '#f8f9fa' }}>
                      {Object.values(row).map((cell, cellIndex) => (
                        <td key={cellIndex} style={{ padding: '8px 10px', border: '1px solid #dee2e6', whiteSpace: 'nowrap' }}>
                          {cell !== null && cell !== undefined ? cell.toString() : ''}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default Cumulative;
