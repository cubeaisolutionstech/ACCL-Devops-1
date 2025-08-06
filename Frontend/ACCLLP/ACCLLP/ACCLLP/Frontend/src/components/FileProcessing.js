import React, { useState, useRef } from 'react';
import '../dashboard.css';
import './CompanyProductMapping.css';

const FileProcessing = () => {
  const [activeTab, setActiveTab] = useState('budget');
  // File upload state
  const [selectedFile, setSelectedFile] = useState(null);
  const [uploadMessage, setUploadMessage] = useState('');
  const fileInputRef = useRef();
  const [dragActive, setDragActive] = useState(false);

  const [salesFile, setSalesFile] = useState(null);
  const [salesUploadMessage, setSalesUploadMessage] = useState('');
  const salesFileInputRef = useRef();
  const [salesDragActive, setSalesDragActive] = useState(false);

  const [osFile, setOsFile] = useState(null);
  const [osUploadMessage, setOsUploadMessage] = useState('');
  const osFileInputRef = useRef();
  const [osDragActive, setOsDragActive] = useState(false);

  const handleFileChange = (e) => {
    if (e.target.files && e.target.files[0]) {
      setSelectedFile(e.target.files[0]);
      setUploadMessage('');
    }
  };
  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      setSelectedFile(e.dataTransfer.files[0]);
      setUploadMessage('');
    }
  };
  const handleDragOver = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(true);
  };
  const handleDragLeave = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
  };
  const handleUpload = (e) => {
    e.preventDefault();
    setUploadMessage('File processed successfully!');
    setTimeout(() => setUploadMessage(''), 2000);
  };

  // Sales Processing handlers
  const handleSalesFileChange = (e) => {
    if (e.target.files && e.target.files[0]) {
      setSalesFile(e.target.files[0]);
      setSalesUploadMessage('');
    }
  };
  const handleSalesDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setSalesDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      setSalesFile(e.dataTransfer.files[0]);
      setSalesUploadMessage('');
    }
  };
  const handleSalesDragOver = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setSalesDragActive(true);
  };
  const handleSalesDragLeave = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setSalesDragActive(false);
  };
  const handleSalesUpload = (e) => {
    e.preventDefault();
    setSalesUploadMessage('File processed successfully!');
    setTimeout(() => setSalesUploadMessage(''), 2000);
  };

  // OS Processing handlers
  const handleOsFileChange = (e) => {
    if (e.target.files && e.target.files[0]) {
      setOsFile(e.target.files[0]);
      setOsUploadMessage('');
    }
  };
  const handleOsDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setOsDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      setOsFile(e.dataTransfer.files[0]);
      setOsUploadMessage('');
    }
  };
  const handleOsDragOver = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setOsDragActive(true);
  };
  const handleOsDragLeave = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setOsDragActive(false);
  };
  const handleOsUpload = (e) => {
    e.preventDefault();
    setOsUploadMessage('File processed successfully!');
    setTimeout(() => setOsUploadMessage(''), 2000);
  };

  return (
    <div className="management-section" style={{ maxWidth: 1200, margin: '32px auto' }}>
      <h2>File Processing</h2>
      <div className="management-tabs" style={{ marginBottom: 32 }}>
        <button
          className={`tab-btn ${activeTab === 'budget' ? 'active' : ''}`}
          onClick={() => setActiveTab('budget')}
        >
          Budget Processing
        </button>
        <button
          className={`tab-btn ${activeTab === 'sales' ? 'active' : ''}`}
          onClick={() => setActiveTab('sales')}
        >
          Sales Processing
        </button>
        <button
          className={`tab-btn ${activeTab === 'os' ? 'active' : ''}`}
          onClick={() => setActiveTab('os')}
        >
          OS Processing
        </button>
      </div>
      {activeTab === 'budget' && (
        <div>
          <h3>Process Budget File <span style={{ fontSize: 18, color: '#aaa', marginLeft: 6 }}>↔️</span></h3>
          <div style={{ marginBottom: 16, fontWeight: 500 }}>Upload Budget File</div>
          <div
            className="drop-zone"
            style={{ background: dragActive ? '#23242a' : undefined, borderColor: dragActive ? '#ff4c4c' : undefined, transition: 'background 0.2s, border 0.2s', maxWidth: 700 }}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
          >
            <input
              type="file"
              id="budget-file-upload"
              style={{ display: 'none' }}
              ref={fileInputRef}
              accept=".xlsx,.xls"
              onChange={handleFileChange}
            />
            <label htmlFor="budget-file-upload" className="drop-label" style={{ cursor: 'pointer' }}>
              <div className="drop-icon">
                <svg width="32" height="32" fill="none" stroke="#aaa" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" viewBox="0 0 24 24"><path d="M12 19V6M5 12l7-7 7 7"/></svg>
              </div>
              <p>Drag and drop file here</p>
              <p className="file-limit">Limit 200MB per file • XLSX, XLS</p>
              <button
                className="browse-btn"
                type="button"
                onClick={e => {
                  e.preventDefault();
                  if (fileInputRef.current) fileInputRef.current.click();
                }}
              >
                Browse files
              </button>
              {selectedFile && (
                <div style={{ marginTop: 10, color: '#fff' }}>
                  <b>Selected:</b> {selectedFile.name}
                  <button
                    className="browse-btn"
                    style={{ marginLeft: 12 }}
                    type="button"
                    onClick={handleUpload}
                  >
                    Upload
                  </button>
                </div>
              )}
              {uploadMessage && (
                <div style={{ marginTop: 8, color: '#4caf50', fontWeight: 600 }}>{uploadMessage}</div>
              )}
            </label>
          </div>
        </div>
      )}
      {activeTab === 'sales' && (
        <div>
          <h3>Process Sales File <span style={{ fontSize: 18, color: '#aaa', marginLeft: 6 }}>↔️</span></h3>
          <div style={{ marginBottom: 16, fontWeight: 500 }}>Upload Sales File</div>
          <div
            className="drop-zone"
            style={{ background: salesDragActive ? '#23242a' : undefined, borderColor: salesDragActive ? '#ff4c4c' : undefined, transition: 'background 0.2s, border 0.2s', maxWidth: 700 }}
            onDragOver={handleSalesDragOver}
            onDragLeave={handleSalesDragLeave}
            onDrop={handleSalesDrop}
          >
            <input
              type="file"
              id="sales-file-upload"
              style={{ display: 'none' }}
              ref={salesFileInputRef}
              accept=".xlsx,.xls"
              onChange={handleSalesFileChange}
            />
            <label htmlFor="sales-file-upload" className="drop-label" style={{ cursor: 'pointer' }}>
              <div className="drop-icon">
                <svg width="32" height="32" fill="none" stroke="#aaa" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" viewBox="0 0 24 24"><path d="M12 19V6M5 12l7-7 7 7"/></svg>
              </div>
              <p>Drag and drop file here</p>
              <p className="file-limit">Limit 200MB per file • XLSX, XLS</p>
              <button
                className="browse-btn"
                type="button"
                onClick={e => {
                  e.preventDefault();
                  if (salesFileInputRef.current) salesFileInputRef.current.click();
                }}
              >
                Browse files
              </button>
              {salesFile && (
                <div style={{ marginTop: 10, color: '#fff' }}>
                  <b>Selected:</b> {salesFile.name}
                  <button
                    className="browse-btn"
                    style={{ marginLeft: 12 }}
                    type="button"
                    onClick={handleSalesUpload}
                  >
                    Upload
                  </button>
                </div>
              )}
              {salesUploadMessage && (
                <div style={{ marginTop: 8, color: '#4caf50', fontWeight: 600 }}>{salesUploadMessage}</div>
              )}
            </label>
          </div>
        </div>
      )}
      {activeTab === 'os' && (
        <div>
          <h3>Process OS File <span style={{ fontSize: 18, color: '#aaa', marginLeft: 6 }}>↔️</span></h3>
          <div style={{ marginBottom: 16, fontWeight: 500 }}>Upload OS File</div>
          <div
            className="drop-zone"
            style={{ background: osDragActive ? '#23242a' : undefined, borderColor: osDragActive ? '#ff4c4c' : undefined, transition: 'background 0.2s, border 0.2s', maxWidth: 700 }}
            onDragOver={handleOsDragOver}
            onDragLeave={handleOsDragLeave}
            onDrop={handleOsDrop}
          >
            <input
              type="file"
              id="os-file-upload"
              style={{ display: 'none' }}
              ref={osFileInputRef}
              accept=".xlsx,.xls"
              onChange={handleOsFileChange}
            />
            <label htmlFor="os-file-upload" className="drop-label" style={{ cursor: 'pointer' }}>
              <div className="drop-icon">
                <svg width="32" height="32" fill="none" stroke="#aaa" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" viewBox="0 0 24 24"><path d="M12 19V6M5 12l7-7 7 7"/></svg>
              </div>
              <p>Drag and drop file here</p>
              <p className="file-limit">Limit 200MB per file • XLSX, XLS</p>
              <button
                className="browse-btn"
                type="button"
                onClick={e => {
                  e.preventDefault();
                  if (osFileInputRef.current) osFileInputRef.current.click();
                }}
              >
                Browse files
              </button>
              {osFile && (
                <div style={{ marginTop: 10, color: '#fff' }}>
                  <b>Selected:</b> {osFile.name}
                  <button
                    className="browse-btn"
                    style={{ marginLeft: 12 }}
                    type="button"
                    onClick={handleOsUpload}
                  >
                    Upload
                  </button>
                </div>
              )}
              {osUploadMessage && (
                <div style={{ marginTop: 8, color: '#4caf50', fontWeight: 600 }}>{osUploadMessage}</div>
              )}
            </label>
          </div>
        </div>
      )}
    </div>
  );
};

export default FileProcessing; 