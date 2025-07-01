import React, { useState, useRef, useEffect } from 'react';
import '../dashboard.css';
import './ExecutiveManagement.css';
import { FaEye, FaSearch, FaTimes, FaTrash } from 'react-icons/fa';
import { FiMaximize } from 'react-icons/fi';

const ExecutiveManagement = ({ onStatsChange = () => {} }) => {
  const [activeTab, setActiveTab] = useState('executiveCreation');
  const [executives, setExecutives] = useState([
    { name: 'ADDHISESHAN.R', code: '1212', customers: 0, branch: 'Chennai', region: 1 },
    { name: 'AFSHANA.A', code: '1231', customers: 0, branch: 'Madurai', region: 2 },
    { name: 'ALAGUMAIN B', code: '1183', customers: 0, branch: 'Trichy', region: 3 },
    { name: 'ANANDHA RAMAKRISHNAN M', code: '1292', customers: 76, branch: 'Coimbatore', region: 4 },
    { name: 'ARUNAGRILK', code: '1197', customers: 0, branch: 'Karur', region: 5 },
    { name: 'BALAJ.I.K', code: '1155', customers: 157, branch: 'Salem', region: 6 },
    { name: 'DEEPA R', code: '1138', customers: 39, branch: 'Erode', region: 7 }
  ]);

  const [newExecutive, setNewExecutive] = useState({
    name: '',
    code: ''
  });

  // Table view states
  const [showActions, setShowActions] = useState(false);
  const [showDropdown, setShowDropdown] = useState(false);
  const [showSearch, setShowSearch] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [showColumns, setShowColumns] = useState({
    name: true,
    code: true,
    customers: true,
    region: true
  });
  const [zoomed, setZoomed] = useState(false);
  const [showFull, setShowFull] = useState(true);

  // Remove Executive section
  const [removeExec, setRemoveExec] = useState('');
  const [removeWarning, setRemoveWarning] = useState(false);

  // Refs for dropdown and search modal
  const actionsRef = useRef();
  const dropdownRef = useRef();
  const searchModalRef = useRef();
  const fileInputRef = useRef();

  // Handle outside clicks for dropdown and search
  useEffect(() => {
    if (!showDropdown) return;
    function handleClick(e) {
      if (
        dropdownRef.current &&
        !dropdownRef.current.contains(e.target) &&
        actionsRef.current &&
        !actionsRef.current.contains(e.target)
      ) {
        setShowDropdown(false);
      }
    }
    function handleEsc(e) { if (e.key === 'Escape') setShowDropdown(false); }
    document.addEventListener('mousedown', handleClick);
    document.addEventListener('keydown', handleEsc);
    return () => {
      document.removeEventListener('mousedown', handleClick);
      document.removeEventListener('keydown', handleEsc);
    };
  }, [showDropdown]);

  useEffect(() => {
    if (!showSearch) return;
    function handleClick(e) {
      if (searchModalRef.current && !searchModalRef.current.contains(e.target)) setShowSearch(false);
    }
    function handleEsc(e) { if (e.key === 'Escape') setShowSearch(false); }
    document.addEventListener('mousedown', handleClick);
    document.addEventListener('keydown', handleEsc);
    return () => {
      document.removeEventListener('mousedown', handleClick);
      document.removeEventListener('keydown', handleEsc);
    };
  }, [showSearch]);

  const handleInputChange = (e) => {
    const { name, value } = e.target;
    setNewExecutive(prev => ({ ...prev, [name]: value }));
  };

  const handleAddExecutive = (e) => {
    e.preventDefault();
    if (newExecutive.name && newExecutive.code) {
      setExecutives([...executives, {
        name: newExecutive.name,
        code: newExecutive.code,
        customers: 0,
        branch: 'New Branch',
        region: 0
      }]);
      setNewExecutive({ name: '', code: '' });
    }
  };

  const handleRemoveExecutive = () => {
    setExecutives(executives.filter(exec => exec.name !== removeExec));
    setRemoveExec('');
    setRemoveWarning(false);
  };

  // Filter executives based on search term
  const filteredExecutives = executives.filter(exec => {
    const term = searchQuery.toLowerCase();
    return (
      (showColumns.name && exec.name.toLowerCase().includes(term)) ||
      (showColumns.code && exec.code.toLowerCase().includes(term)) ||
      (showColumns.customers && exec.customers.toString().includes(term)) ||
      (showColumns.region && exec.region.toString().includes(term))
    );
  });

  // Download CSV
  const handleDownload = () => {
    const header = ['Name', 'Code', 'Customers', 'Region'].join(',') + '\n';
    const rows = filteredExecutives.map(exec =>
      [exec.name, exec.code, exec.customers, exec.region].join(',')
    ).join('\n');
    const csv = header + rows;
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'executives.csv';
    a.click();
    URL.revokeObjectURL(url);
  };

  const tableContent = (
    <>
      <div className="em-mapping-header-row">
        <h3>Current Executives</h3>
        {!zoomed && (
          <div className={`em-mapping-actions${showActions ? ' visible' : ''}`} ref={actionsRef}>
            <button
              className={`em-mapping-action em-eye-action${showDropdown ? ' active' : ''}`}
              title="Show/Hide Columns"
              onClick={() => setShowDropdown(v => !v)}
              tabIndex={0}
            >
              <FaEye />
            </button>
            {showDropdown && (
              <div className="em-mapping-dropdown" ref={dropdownRef}>
                <div className="em-mapping-dropdown-row">
                  <input
                    type="checkbox"
                    checked={showColumns.name}
                    onChange={() => setShowColumns(sc => ({ ...sc, name: !sc.name }))}
                  />
                  <span>Name</span>
                </div>
                <div className="em-mapping-dropdown-row">
                  <input
                    type="checkbox"
                    checked={showColumns.code}
                    onChange={() => setShowColumns(sc => ({ ...sc, code: !sc.code }))}
                  />
                  <span>Code</span>
                </div>
                <div className="em-mapping-dropdown-row">
                  <input
                    type="checkbox"
                    checked={showColumns.customers}
                    onChange={() => setShowColumns(sc => ({ ...sc, customers: !sc.customers }))}
                  />
                  <span>Customers</span>
                </div>
                <div className="em-mapping-dropdown-row">
                  <input
                    type="checkbox"
                    checked={showColumns.region}
                    onChange={() => setShowColumns(sc => ({ ...sc, region: !sc.region }))}
                  />
                  <span>Region</span>
                </div>
              </div>
            )}
            <button className="em-mapping-action" title="Zoom" onClick={() => setZoomed(z => !z)}>
              <FiMaximize />
            </button>
            <button className="em-mapping-action" title="Search" onClick={() => setShowSearch(true)}>
              <FaSearch />
            </button>
            <button className="em-mapping-action" title="Download CSV" onClick={handleDownload}>
              <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#aaa" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
            </button>
          </div>
        )}
      </div>

      <table>
        <thead>
          <tr>
            {showColumns.name && <th>Executive</th>}
            {showColumns.code && <th>Code</th>}
            {showColumns.customers && <th>Customers</th>}
            {showColumns.region && <th>Region</th>}
          </tr>
        </thead>
        <tbody>
          {filteredExecutives.map((exec, index) => (
            <tr key={index}>
              {showColumns.name && <td>{exec.name}</td>}
              {showColumns.code && <td>{exec.code}</td>}
              {showColumns.customers && <td>{exec.customers}</td>}
              {showColumns.region && <td>{exec.region}</td>}
            </tr>
          ))}
        </tbody>
      </table>
    </>
  );

  // Remove Executive section below the table
  const selectedExec = executives.find(e => e.name === removeExec);

  useEffect(() => {
    if (onStatsChange) onStatsChange({ executives: executives.length });
  }, [executives, onStatsChange]);

  return (
    <div className="management-section">
      <h2>Executive Management</h2>
      
      <div className="management-tabs">
        <button 
          className={`tab-btn ${activeTab === 'executiveCreation' ? 'active' : ''}`}
          onClick={() => setActiveTab('executiveCreation')}
        >
          Executive Creation
        </button>
        <button 
          className={`tab-btn ${activeTab === 'customerCodeManagement' ? 'active' : ''}`}
          onClick={() => setActiveTab('customerCodeManagement')}
        >
          Customer Code Management
        </button>
      </div>

      {activeTab === 'executiveCreation' ? (
        <>
          <div className="add-executive">
            <h3>Add New Executive</h3>
            <form onSubmit={handleAddExecutive}>
              <div className="form-group">
                <label>Executive Name:</label>
                <input 
                  type="text" 
                  name="name"
                  value={newExecutive.name}
                  onChange={handleInputChange}
                />
              </div>
              <div className="form-group">
                <label>Executive Code:</label>
                <input 
                  type="text" 
                  name="code"
                  value={newExecutive.code}
                  onChange={handleInputChange}
                />
              </div>
              <button type="submit" className="add-btn">Add Executive</button>
            </form>
          </div>

          <div 
            className={`em-current-mappings${zoomed ? ' em-zoomed' : ''}`}
            onMouseEnter={() => setShowActions(true)}
            onMouseLeave={() => { setShowActions(false); setShowDropdown(false); }}
          >
            {zoomed && (
              <div className="em-zoom-modal">
                <div className="em-zoom-modal-content">
                  <button className="em-zoom-close" onClick={() => setZoomed(false)} title="Close">&times;</button>
                  {tableContent}
                </div>
              </div>
            )}
            {!zoomed && tableContent}
            {showSearch && (
              <div className="em-search-modal">
                <div className="em-search-modal-content" ref={searchModalRef}>
                  <input
                    className="em-mapping-search em-mapping-search-modal"
                    type="text"
                    placeholder="Search..."
                    value={searchQuery}
                    autoFocus
                    onChange={e => setSearchQuery(e.target.value)}
                  />
                  <button className="em-search-close" onClick={() => setShowSearch(false)} title="Close">&times;</button>
                </div>
              </div>
            )}
          </div>

          {/* Remove Executive Section Below Table */}
          <div className="remove-executive-section">
            <label htmlFor="remove-exec-select" style={{ fontWeight: 500, marginBottom: 6 }}>Remove Executive:</label>
            <select
              id="remove-exec-select"
              className="remove-exec-dropdown"
              value={removeExec}
              onChange={e => {
                setRemoveExec(e.target.value);
                setRemoveWarning(!!e.target.value);
              }}
              style={{ display: 'block', width: '100%', marginBottom: 12, marginTop: 2 }}
            >
              <option value="">Select Executive</option>
              {executives.map((exec, idx) => (
                <option key={idx} value={exec.name}>{exec.name}</option>
              ))}
            </select>
            {removeWarning && selectedExec && (
              <div style={{ background: '#59562b', color: '#fff', borderRadius: 8, padding: '12px 16px', marginBottom: 12, display: 'flex', alignItems: 'center', gap: 10 }}>
                <span style={{ color: '#ffcc00', fontSize: 18, marginRight: 8 }}>⚠️</span>
                <span>Removing '{selectedExec.name}' will affect:</span>
              </div>
            )}
            {removeWarning && selectedExec && (
              <ul style={{ margin: '0 0 12px 0', padding: 0, listStyle: 'disc inside' }}>
                <li><b>Branches:</b> {selectedExec.branch}</li>
                <li><b>Region:</b> {selectedExec.region}</li>
              </ul>
            )}
            <button
              className="remove-executive-btn remove-executive-btn-block"
              style={{ marginTop: 8, width: 'fit-content', display: 'flex', alignItems: 'center', gap: 8 }}
              disabled={!removeExec}
              onClick={handleRemoveExecutive}
            >
              <FaTrash style={{ marginRight: 4 }} /> Remove Executive
            </button>
          </div>
        </>
      ) : (
        <div className="customer-code-management">
          <div className="bulk-assignment">
            <h3>Bulk Customer Assignment</h3>
            <p>Upload Executive-Customer File</p>
            
            <div className="drop-zone">
              <input 
                type="file" 
                id="file-upload" 
                style={{ display: 'none' }}
                ref={fileInputRef}
              />
              <label htmlFor="file-upload" className="drop-label">
                <div className="drop-icon">[ ]</div>
                <p>Drag and drop file here</p>
                <p className="file-limit">Limit 200MB per file - XLS, XLSX</p>
                <button
                  className="browse-btn"
                  type="button"
                  onClick={e => { e.preventDefault(); fileInputRef.current && fileInputRef.current.click(); }}
                >
                  Browse files
                </button>
              </label>
            </div>
          </div>

          <div className="divider"></div>

          <div className="manual-management">
            <h3>Manual Customer Management</h3>
            
            <div className="form-group">
              <label>Select Executive:</label>
              <select>
                <option value="">Select an executive</option>
                {executives.map((exec, index) => (
                  <option key={index} value={exec.name}>
                    {exec.name}
                  </option>
                ))}
              </select>
            </div>

            <div className="customer-management-view">
              <h4>Manual Customer Management</h4>
              <p>Select Executive:</p>
              <div className="executive-info">
                <strong>No executive selected</strong>
              </div>

              <div className="divider"></div>

              <div className="assigned-customers">
                <h4>Assigned Customers (0)</h4>
                <p>No customers assigned</p>
              </div>

              <div className="divider"></div>

              <div className="actions">
                <h4>Actions</h4>
                <p>Add New Customer Codes (one per line):</p>
                <textarea rows={4} />
                <button className="add-codes-btn">
                  Add Codes
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default ExecutiveManagement; 