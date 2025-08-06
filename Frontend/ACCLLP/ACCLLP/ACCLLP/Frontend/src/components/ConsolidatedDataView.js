import React, { useState, useRef, useEffect } from 'react';
import '../dashboard.css';
import './CompanyProductMapping.css';
import { FaEye, FaSearch, FaTimes } from 'react-icons/fa';
import { FiMaximize } from 'react-icons/fi';

const dummyExecutives = [
  'YASMEEN MARTIN N',
  'SHESATHRI S',
  'SRINIVASAN T',
  'BALAJI K',
  'PREETHI G',
  'VIJAY RAO S',
  'ANANDHA RAMAKRISHNAN M',
];
const dummyBranches = [
  'Erode',
  'Coimbatore',
  'Salem',
  'Karur',
  'Chennai',
  'West',
];
const dummyData = [
  { code: '212008', name: 'CATALYST CHEMICAL INDUSTRIES', executive: 'YASMEEN MARTIN N', execCode: '1138', branch: 'Erode', region: 'West' },
  { code: '212032', name: 'LAKSHMI BALAJI AGENCY', executive: 'YASMEEN MARTIN N', execCode: '1138', branch: 'Erode', region: 'West' },
  { code: '212109', name: 'R.V.COLOURS', executive: 'YASMEEN MARTIN N', execCode: '1138', branch: 'Erode', region: 'West' },
  { code: '212117', name: 'SHRI VELA DYEING', executive: 'YASMEEN MARTIN N', execCode: '1138', branch: 'Erode', region: 'West' },
  { code: '210070', name: 'CAVINKARE PVT LTD', executive: 'SHESATHRI S', execCode: '1299', branch: 'Erode', region: 'West' },
  { code: '210334', name: 'RAJHAM REFINERIES,ANTHIYUR', executive: 'SHESATHRI S', execCode: '1299', branch: 'Erode', region: 'West' },
  // ... more rows as needed
];

const ConsolidatedDataView = () => {
  // Filters
  const [selectedExecutives, setSelectedExecutives] = useState(['All']);
  const [selectedBranches, setSelectedBranches] = useState(['All']);
  const [customerSearch, setCustomerSearch] = useState('');

  // Table controls
  const [showActions, setShowActions] = useState(false);
  const [showDropdown, setShowDropdown] = useState(false);
  const [showSearch, setShowSearch] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [zoomed, setZoomed] = useState(false);
  const [showColumns, setShowColumns] = useState({
    code: true,
    name: true,
    executive: true,
    execCode: true,
    branch: true,
    region: true,
  });
  const actionsRef = useRef();
  const dropdownRef = useRef();
  const searchModalRef = useRef();

  // Dropdown and search modal close
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

  // Multi-select logic for filters
  const handleExecutiveChange = (value) => {
    if (value === 'All') {
      setSelectedExecutives(['All']);
    } else {
      setSelectedExecutives((prev) => {
        if (prev.includes('All')) return [value];
        if (prev.includes(value)) return prev.filter((v) => v !== value);
        return [...prev, value];
      });
    }
  };
  const handleBranchChange = (value) => {
    if (value === 'All') {
      setSelectedBranches(['All']);
    } else {
      setSelectedBranches((prev) => {
        if (prev.includes('All')) return [value];
        if (prev.includes(value)) return prev.filter((v) => v !== value);
        return [...prev, value];
      });
    }
  };

  // Filtering data
  const filteredData = dummyData.filter(row => {
    const execMatch = selectedExecutives.includes('All') || selectedExecutives.includes(row.executive);
    const branchMatch = selectedBranches.includes('All') || selectedBranches.includes(row.branch);
    const customerMatch =
      row.name.toLowerCase().includes(customerSearch.toLowerCase()) ||
      row.code.toLowerCase().includes(customerSearch.toLowerCase());
    const tableSearchMatch =
      row.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
      row.code.toLowerCase().includes(searchQuery.toLowerCase()) ||
      row.executive.toLowerCase().includes(searchQuery.toLowerCase()) ||
      row.execCode.toLowerCase().includes(searchQuery.toLowerCase()) ||
      row.branch.toLowerCase().includes(searchQuery.toLowerCase()) ||
      row.region.toLowerCase().includes(searchQuery.toLowerCase());
    return execMatch && branchMatch && customerMatch && tableSearchMatch;
  });

  // Download CSV
  const handleDownload = () => {
    const header = [
      showColumns.code ? 'Customer Code' : null,
      showColumns.name ? 'Customer Name' : null,
      showColumns.executive ? 'Executive' : null,
      showColumns.execCode ? 'Executive Code' : null,
      showColumns.branch ? 'Branch' : null,
      showColumns.region ? 'Region' : null,
    ].filter(Boolean).join(',') + '\n';
    const rows = filteredData.map(row => [
      showColumns.code ? row.code : null,
      showColumns.name ? '"' + row.name + '"' : null,
      showColumns.executive ? row.executive : null,
      showColumns.execCode ? row.execCode : null,
      showColumns.branch ? row.branch : null,
      showColumns.region ? row.region : null,
    ].filter(Boolean).join(','));
    const csv = header + rows.join('\n');
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'consolidated_data.csv';
    a.click();
    URL.revokeObjectURL(url);
  };

  // Table UI
  const tableContent = (
    <>
      <div className="em-mapping-header-row">
        <h3 style={{ margin: 0 }}> </h3>
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
                    checked={showColumns.code}
                    onChange={() => setShowColumns(sc => ({ ...sc, code: !sc.code }))}
                  />
                  <span>Customer Code</span>
                </div>
                <div className="em-mapping-dropdown-row">
                  <input
                    type="checkbox"
                    checked={showColumns.name}
                    onChange={() => setShowColumns(sc => ({ ...sc, name: !sc.name }))}
                  />
                  <span>Customer Name</span>
                </div>
                <div className="em-mapping-dropdown-row">
                  <input
                    type="checkbox"
                    checked={showColumns.executive}
                    onChange={() => setShowColumns(sc => ({ ...sc, executive: !sc.executive }))}
                  />
                  <span>Executive</span>
                </div>
                <div className="em-mapping-dropdown-row">
                  <input
                    type="checkbox"
                    checked={showColumns.execCode}
                    onChange={() => setShowColumns(sc => ({ ...sc, execCode: !sc.execCode }))}
                  />
                  <span>Executive Code</span>
                </div>
                <div className="em-mapping-dropdown-row">
                  <input
                    type="checkbox"
                    checked={showColumns.branch}
                    onChange={() => setShowColumns(sc => ({ ...sc, branch: !sc.branch }))}
                  />
                  <span>Branch</span>
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
            {showColumns.code && <th>Customer Code</th>}
            {showColumns.name && <th>Customer Name</th>}
            {showColumns.executive && <th>Executive</th>}
            {showColumns.execCode && <th>Executive Code</th>}
            {showColumns.branch && <th>Branch</th>}
            {showColumns.region && <th>Region</th>}
          </tr>
        </thead>
        <tbody>
          {filteredData.map((row, idx) => (
            <tr key={idx}>
              {showColumns.code && <td>{row.code}</td>}
              {showColumns.name && <td>{row.name}</td>}
              {showColumns.executive && <td>{row.executive}</td>}
              {showColumns.execCode && <td>{row.execCode}</td>}
              {showColumns.branch && <td>{row.branch}</td>}
              {showColumns.region && <td>{row.region}</td>}
            </tr>
          ))}
        </tbody>
      </table>
    </>
  );

  return (
    <div className="management-section" style={{ maxWidth: 1400, margin: '32px auto' }}>
      <h2>Consolidated Data View</h2>
      <div style={{ display: 'flex', gap: 24, marginBottom: 24, alignItems: 'flex-end' }}>
        <div style={{ flex: 1 }}>
          <label style={{ fontWeight: 500 }}>Filter by Executive:</label>
          <div style={{ display: 'flex', gap: 8, marginTop: 6 }}>
            <div style={{ flex: 1, position: 'relative' }}>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                <div
                  style={{
                    background: '#ff4c4c', color: '#fff', borderRadius: 16, padding: '2px 14px', fontWeight: 600, display: 'flex', alignItems: 'center', gap: 4
                  }}
                >
                  All
                  <FaTimes style={{ marginLeft: 4, cursor: 'pointer' }} onClick={() => setSelectedExecutives([])} />
                </div>
                {dummyExecutives.map(exec =>
                  selectedExecutives.includes(exec) && (
                    <div
                      key={exec}
                      style={{ background: '#ff4c4c', color: '#fff', borderRadius: 16, padding: '2px 14px', fontWeight: 600, display: 'flex', alignItems: 'center', gap: 4 }}
                    >
                      {exec}
                      <FaTimes style={{ marginLeft: 4, cursor: 'pointer' }} onClick={() => handleExecutiveChange(exec)} />
                    </div>
                  )
                )}
              </div>
              <select
                style={{ width: '100%', marginTop: 6, height: 40, fontSize: '1rem', borderRadius: 6, border: '1px solid #333', background: '#191b20', color: '#fff' }}
                value={selectedExecutives[selectedExecutives.length - 1] || 'All'}
                onChange={e => handleExecutiveChange(e.target.value)}
              >
                <option value="All">All</option>
                {dummyExecutives.map(exec => (
                  <option key={exec} value={exec}>{exec}</option>
                ))}
              </select>
            </div>
          </div>
        </div>
        <div style={{ flex: 1 }}>
          <label style={{ fontWeight: 500 }}>Filter by Branch:</label>
          <div style={{ display: 'flex', gap: 8, marginTop: 6 }}>
            <div style={{ flex: 1, position: 'relative' }}>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                <div
                  style={{
                    background: '#ff4c4c', color: '#fff', borderRadius: 16, padding: '2px 14px', fontWeight: 600, display: 'flex', alignItems: 'center', gap: 4
                  }}
                >
                  All
                  <FaTimes style={{ marginLeft: 4, cursor: 'pointer' }} onClick={() => setSelectedBranches([])} />
                </div>
                {dummyBranches.map(branch =>
                  selectedBranches.includes(branch) && (
                    <div
                      key={branch}
                      style={{ background: '#ff4c4c', color: '#fff', borderRadius: 16, padding: '2px 14px', fontWeight: 600, display: 'flex', alignItems: 'center', gap: 4 }}
                    >
                      {branch}
                      <FaTimes style={{ marginLeft: 4, cursor: 'pointer' }} onClick={() => handleBranchChange(branch)} />
                    </div>
                  )
                )}
              </div>
              <select
                style={{ width: '100%', marginTop: 6, height: 40, fontSize: '1rem', borderRadius: 6, border: '1px solid #333', background: '#191b20', color: '#fff' }}
                value={selectedBranches[selectedBranches.length - 1] || 'All'}
                onChange={e => handleBranchChange(e.target.value)}
              >
                <option value="All">All</option>
                {dummyBranches.map(branch => (
                  <option key={branch} value={branch}>{branch}</option>
                ))}
              </select>
            </div>
          </div>
        </div>
        <div style={{ flex: 1, display: 'flex', flexDirection: 'column', justifyContent: 'flex-end' }}>
          <label style={{ fontWeight: 500 }}>Search Customer:</label>
          <input
            type="text"
            style={{ width: '100%', marginTop: 6, height: 40, fontSize: '1rem', borderRadius: 6, border: '1px solid #333', background: '#191b20', color: '#fff' }}
            placeholder="Search by name or code..."
            value={customerSearch}
            onChange={e => setCustomerSearch(e.target.value)}
          />
        </div>
      </div>
      <div style={{ marginBottom: 16, fontWeight: 500 }}>
        Results: {filteredData.length} records
      </div>
      <button className="add-btn" style={{ marginBottom: 18 }} onClick={handleDownload}>Download CSV</button>
      <div
        className="em-current-mappings"
        onMouseEnter={() => setShowActions(true)}
        onMouseLeave={() => { setShowActions(false); setShowDropdown(false); }}
        style={{ marginTop: 0 }}
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
    </div>
  );
};

export default ConsolidatedDataView; 