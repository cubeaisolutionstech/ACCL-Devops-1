import React, { useRef, useState } from 'react';
import '../dashboard.css';
import './CompanyProductMapping.css';

const BackupRestore = () => {
  // Simulated counts
  const [counts] = useState({ branches: 11, regions: 3, companies: 5 });
  // Backup
  const handleBackup = () => {
    const data = JSON.stringify({ message: 'This is a dummy backup file.' }, null, 2);
    const blob = new Blob([data], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'backup.json';
    a.click();
    URL.revokeObjectURL(url);
  };
  // Restore
  const [selectedFile, setSelectedFile] = useState(null);
  const [uploadMessage, setUploadMessage] = useState('');
  const fileInputRef = useRef();
  const [dragActive, setDragActive] = useState(false);

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
    setUploadMessage('Backup restored successfully!');
    setTimeout(() => setUploadMessage(''), 2000);
  };

  return (
    <div className="management-section" style={{ maxWidth: 1200, margin: '32px auto' }}>
      <h2>Backup & Restore</h2>
      <div style={{ display: 'flex', gap: 40, alignItems: 'flex-start', marginTop: 32 }}>
        {/* Backup Section */}
        <div style={{ flex: 1 }}>
          <h3>Backup Mappings</h3>
          <p>Export branch, region, and company mappings:</p>
          <ul style={{ marginBottom: 24 }}>
            <li>Branches: {counts.branches}</li>
            <li>Regions: {counts.regions}</li>
            <li>Companies: {counts.companies}</li>
          </ul>
          <button className="add-btn" onClick={handleBackup}>Create Backup</button>
        </div>
        {/* Restore Section */}
        <div style={{ flex: 1 }}>
          <h3>Restore Mappings</h3>
          <p>Upload Backup File</p>
          <div
            className="drop-zone"
            style={{ background: dragActive ? '#23242a' : undefined, borderColor: dragActive ? '#ff4c4c' : undefined, transition: 'background 0.2s, border 0.2s' }}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
          >
            <input
              type="file"
              id="restore-file-upload"
              style={{ display: 'none' }}
              ref={fileInputRef}
              accept=".json"
              onChange={handleFileChange}
            />
            <label htmlFor="restore-file-upload" className="drop-label" style={{ cursor: 'pointer' }}>
              <div className="drop-icon">
                <svg width="32" height="32" fill="none" stroke="#aaa" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" viewBox="0 0 24 24"><path d="M12 19V6M5 12l7-7 7 7"/></svg>
              </div>
              <p>Drag and drop file here</p>
              <p className="file-limit">Limit 200MB per file • JSON</p>
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
      </div>
    </div>
  );
};

export default BackupRestore; 