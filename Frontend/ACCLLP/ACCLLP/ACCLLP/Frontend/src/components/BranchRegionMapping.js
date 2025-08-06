import React, { useState, useRef, useEffect } from 'react';
import './BranchRegionMapping.css';
import { FaEye, FaSearch, FaDownload, FaTimes } from 'react-icons/fa';
import { FiMaximize } from 'react-icons/fi';

const branchesListDefault = [
  'BGLR', 'Chennai', 'Covai', 'Erode', 'Group', 'Karur', 'Madurai', 'Pondy', 'Poultry', 'Salem'
];
const regionsListDefault = ['Group Company', 'North', 'West'];
const executivesList = [
  'AADHISESHAN.R', 'AFSHANA.A', 'KRISHNAKUMAR ...', 'MANOJ KUMAR.P', 'PREETHI.G', 'RAMKUMAR.A', 'SHARMILA S', 'VIJAY RAO S', 'ANANDHA RAMAKRISHNAN M', 'BALAJI K', 'SRINIVASAN T'
];

const COLUMN_OPTIONS = [
  { key: 'branch', label: 'Branch' },
  { key: 'region', label: 'Region' },
  { key: 'executives', label: 'Executives' },
  { key: 'count', label: 'Count' },
];

const BranchRegionMapping = ({ onStatsChange = () => {} }) => {
  const [activeTab, setActiveTab] = useState('manual');
  const [branchName, setBranchName] = useState('');
  const [regionName, setRegionName] = useState('');
  const [removeBranch, setRemoveBranch] = useState('');
  const [removeRegion, setRemoveRegion] = useState('');
  const [selectedBranch, setSelectedBranch] = useState('Chennai');
  const [selectedRegion, setSelectedRegion] = useState('North');
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef();
  const [searchTerm, setSearchTerm] = useState('');
  const [showFull, setShowFull] = useState(true);
  const [showColumns, setShowColumns] = useState({ branch: true, region: true, executives: true, count: true });
  const [showDropdown, setShowDropdown] = useState(false);
  const [zoomed, setZoomed] = useState(false);
  const [showSearch, setShowSearch] = useState(false);
  const mappingActionsRef = useRef();
  const mappingDropdownRef = useRef();
  const searchModalRef = useRef();
  const [showActions, setShowActions] = useState(false);

  // State for current branches and regions
  const [currentBranches, setCurrentBranches] = useState([
    { branch: 'BGLR', executives: 1, inRegions: 1, executiveList: ['PAVAN.K'] },
    { branch: 'Chennai', executives: 3, inRegions: 1, executiveList: ['AADHISESHAN.R', 'AFSHANA.A', 'KRISHNAKUMAR ...'] },
    { branch: 'Covai', executives: 2, inRegions: 1, executiveList: ['KARTHIK.V', 'KUMARESAN.R'] },
    { branch: 'Erode', executives: 2, inRegions: 1, executiveList: ['DEEPA.R', 'ISWARYA.S'] },
    { branch: 'Group', executives: 1, inRegions: 1, executiveList: ['SIVAKUMAR.A'] },
    { branch: 'Karur', executives: 2, inRegions: 1, executiveList: ['ARUNAGIRI.K', 'H.S.PANDIAN'] },
    { branch: 'Madurai', executives: 3, inRegions: 1, executiveList: ['ALAGUMANI B', 'MARANADU K', 'SIVAJOTHI M'] },
    { branch: 'Pondy', executives: 2, inRegions: 1, executiveList: ['ELAMVAZHUTHY B', 'MAOTSETUNG R'] },
    { branch: 'Poultry', executives: 2, inRegions: 1, executiveList: ['BALAJI K', 'KAMALA KANNAN M'] },
    { branch: 'Salem', executives: 1, inRegions: 1, executiveList: ['SATHEESH KUMAR.M'] },
  ]);
  const [currentRegions, setCurrentRegions] = useState([
    { region: 'Group Company', branches: 1, branchList: ['Group'] },
    { region: 'North', branches: 3, branchList: ['Pondy', 'BGLR', 'Group'] },
    { region: 'West', branches: 7, branchList: ['Madurai', 'Karur', 'Poultry', 'Erode', 'Tirupur', 'Covai', 'Salem'] },
  ]);

  // For select options
  const branchesList = currentBranches.map(b => b.branch);
  const regionsList = currentRegions.map(r => r.region);

  // Map Executives to Branches state
  const [branchExecMap, setBranchExecMap] = useState(() => {
    const map = {};
    currentBranches.forEach(b => { map[b.branch] = b.executiveList ? [...b.executiveList] : []; });
    return map;
  });
  const [selectedExecutives, setSelectedExecutives] = useState(branchExecMap[selectedBranch] || []);

  // Map Branches to Regions state
  const [regionBranchMap, setRegionBranchMap] = useState(() => {
    const map = {};
    currentRegions.forEach(r => { map[r.region] = r.branchList ? [...r.branchList] : []; });
    return map;
  });
  const [selectedBranches, setSelectedBranches] = useState(regionBranchMap[selectedRegion] || []);

  // When selectedBranch or selectedRegion changes, update selectedExecutives/selectedBranches
  React.useEffect(() => {
    setSelectedExecutives(branchExecMap[selectedBranch] || []);
  }, [selectedBranch, branchExecMap]);
  React.useEffect(() => {
    setSelectedBranches(regionBranchMap[selectedRegion] || []);
  }, [selectedRegion, regionBranchMap]);

  // Ensure mapping state is always in sync with currentBranches/currentRegions
  React.useEffect(() => {
    // Sync branchExecMap with currentBranches
    setBranchExecMap(prev => {
      const map = { ...prev };
      currentBranches.forEach(b => {
        if (!map[b.branch]) map[b.branch] = [];
      });
      // Remove deleted branches
      Object.keys(map).forEach(branch => {
        if (!currentBranches.find(b => b.branch === branch)) delete map[branch];
      });
      return map;
    });
  }, [currentBranches]);
  React.useEffect(() => {
    // Sync regionBranchMap with currentRegions
    setRegionBranchMap(prev => {
      const map = { ...prev };
      currentRegions.forEach(r => {
        if (!map[r.region]) map[r.region] = [];
      });
      // Remove deleted regions
      Object.keys(map).forEach(region => {
        if (!currentRegions.find(r => r.region === region)) delete map[region];
      });
      return map;
    });
  }, [currentRegions]);

  // Build current mappings from regionBranchMap and branchExecMap
  const currentMappings = [];
  for (const region in regionBranchMap) {
    for (const branch of regionBranchMap[region]) {
      const execs = branchExecMap[branch] || [];
      currentMappings.push({
        branch,
        region,
        executives: execs.join(', '),
        count: execs.length
      });
    }
  }

  // Search filter
  const filteredMappings = currentMappings.filter(row => {
    const term = searchTerm.toLowerCase();
    return (
      (showColumns.branch && row.branch.toLowerCase().includes(term)) ||
      (showColumns.region && row.region.toLowerCase().includes(term)) ||
      (showColumns.executives && row.executives.toLowerCase().includes(term))
    );
  });

  // Download CSV
  const handleDownload = () => {
    const header = COLUMN_OPTIONS.filter(col => showColumns[col.key]).map(col => col.label).join(',') + '\n';
    const rows = filteredMappings.map(row =>
      COLUMN_OPTIONS.filter(col => showColumns[col.key]).map(col => col.key === 'executives' ? '"' + row[col.key] + '"' : row[col.key]).join(',')
    ).join('\n');
    const csv = header + rows;
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'current_mappings.csv';
    a.click();
    URL.revokeObjectURL(url);
  };

  // Dropdown close on outside click or ESC
  useEffect(() => {
    if (!showDropdown) return;
    function handleClick(e) {
      if (
        mappingDropdownRef.current &&
        !mappingDropdownRef.current.contains(e.target) &&
        mappingActionsRef.current &&
        !mappingActionsRef.current.contains(e.target)
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

  // Search modal close on outside click or ESC
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

  // Handlers for create/remove
  const handleCreateBranch = (e) => {
    e.preventDefault();
    if (branchName.trim() && !branchesList.includes(branchName.trim())) {
      setCurrentBranches(prev => [...prev, { branch: branchName.trim(), executives: 0, inRegions: 0, executiveList: [] }]);
      setBranchExecMap(prev => ({ ...prev, [branchName.trim()]: [] }));
      setBranchName('');
    }
  };
  const handleCreateRegion = (e) => {
    e.preventDefault();
    if (regionName.trim() && !regionsList.includes(regionName.trim())) {
      setCurrentRegions(prev => [...prev, { region: regionName.trim(), branches: 0, branchList: [] }]);
      setRegionBranchMap(prev => ({ ...prev, [regionName.trim()]: [] }));
      setRegionName('');
    }
  };
  const handleRemoveBranch = (e) => {
    e.preventDefault();
    if (removeBranch) {
      setCurrentBranches(prev => prev.filter(b => b.branch !== removeBranch));
      setBranchExecMap(prev => { const map = { ...prev }; delete map[removeBranch]; return map; });
      setRemoveBranch('');
    }
  };
  const handleRemoveRegion = (e) => {
    e.preventDefault();
    if (removeRegion) {
      setCurrentRegions(prev => prev.filter(r => r.region !== removeRegion));
      setRegionBranchMap(prev => { const map = { ...prev }; delete map[removeRegion]; return map; });
      setRemoveRegion('');
    }
  };

  // Map Executives to Branches logic
  const handleExecutiveSelect = (exec) => {
    setSelectedExecutives(prev => prev.includes(exec) ? prev.filter(e => e !== exec) : [...prev, exec]);
  };
  const handleRemoveExecutive = (exec) => {
    setSelectedExecutives(prev => prev.filter(e => e !== exec));
  };
  const handleUpdateBranchMapping = () => {
    setBranchExecMap(prev => ({ ...prev, [selectedBranch]: [...selectedExecutives] }));
    setCurrentBranches(prev => prev.map(b =>
      b.branch === selectedBranch
        ? { ...b, executives: selectedExecutives.length, executiveList: [...selectedExecutives] }
        : b
    ));
    // Force re-render by resetting selectedBranch
    setSelectedBranch('');
    setTimeout(() => setSelectedBranch(selectedBranch), 0);
  };

  // Map Branches to Regions logic
  const handleBranchSelect = (branch) => {
    setSelectedBranches(prev => prev.includes(branch) ? prev.filter(b => b !== branch) : [...prev, branch]);
  };
  const handleRemoveBranchFromRegion = (branch) => {
    setSelectedBranches(prev => prev.filter(b => b !== branch));
  };
  const handleUpdateRegionMapping = () => {
    setRegionBranchMap(prev => ({ ...prev, [selectedRegion]: [...selectedBranches] }));
    setCurrentRegions(prev => prev.map(r =>
      r.region === selectedRegion
        ? { ...r, branches: selectedBranches.length, branchList: [...selectedBranches] }
        : r
    ));
    // Force re-render by resetting selectedRegion
    setSelectedRegion('');
    setTimeout(() => setSelectedRegion(selectedRegion), 0);
  };

  // File upload handlers
  const handleDragOver = (e) => {
    e.preventDefault();
    setIsDragging(true);
  };
  const handleDragLeave = () => {
    setIsDragging(false);
  };
  const handleDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);
    // handle file upload logic here
  };
  const handleFileInput = (e) => {
    // handle file upload logic here
  };

  // Add state for Current Branches table
  const [showBranchesDropdown, setShowBranchesDropdown] = useState(false);
  const [showBranchesSearch, setShowBranchesSearch] = useState(false);
  const [branchesZoomed, setBranchesZoomed] = useState(false);
  const [branchesSearchTerm, setBranchesSearchTerm] = useState('');
  const [branchesShowColumns, setBranchesShowColumns] = useState({ branch: true, executives: true, inRegions: true });
  const branchesDropdownRef = useRef();
  const branchesActionsRef = useRef();
  const branchesSearchModalRef = useRef();

  // Add state for Current Regions table
  const [showRegionsDropdown, setShowRegionsDropdown] = useState(false);
  const [showRegionsSearch, setShowRegionsSearch] = useState(false);
  const [regionsZoomed, setRegionsZoomed] = useState(false);
  const [regionsSearchTerm, setRegionsSearchTerm] = useState('');
  const [regionsShowColumns, setRegionsShowColumns] = useState({ region: true, branches: true });
  const regionsDropdownRef = useRef();
  const regionsActionsRef = useRef();
  const regionsSearchModalRef = useRef();

  // Dropdown close on outside click or ESC for branches
  useEffect(() => {
    if (!showBranchesDropdown) return;
    function handleClick(e) {
      if (
        branchesDropdownRef.current &&
        !branchesDropdownRef.current.contains(e.target) &&
        branchesActionsRef.current &&
        !branchesActionsRef.current.contains(e.target)
      ) {
        setShowBranchesDropdown(false);
      }
    }
    function handleEsc(e) { if (e.key === 'Escape') setShowBranchesDropdown(false); }
    document.addEventListener('mousedown', handleClick);
    document.addEventListener('keydown', handleEsc);
    return () => {
      document.removeEventListener('mousedown', handleClick);
      document.removeEventListener('keydown', handleEsc);
    };
  }, [showBranchesDropdown]);
  // Dropdown close on outside click or ESC for regions
  useEffect(() => {
    if (!showRegionsDropdown) return;
    function handleClick(e) {
      if (
        regionsDropdownRef.current &&
        !regionsDropdownRef.current.contains(e.target) &&
        regionsActionsRef.current &&
        !regionsActionsRef.current.contains(e.target)
      ) {
        setShowRegionsDropdown(false);
      }
    }
    function handleEsc(e) { if (e.key === 'Escape') setShowRegionsDropdown(false); }
    document.addEventListener('mousedown', handleClick);
    document.addEventListener('keydown', handleEsc);
    return () => {
      document.removeEventListener('mousedown', handleClick);
      document.removeEventListener('keydown', handleEsc);
    };
  }, [showRegionsDropdown]);
  // Search modal close for branches
  useEffect(() => {
    if (!showBranchesSearch) return;
    function handleClick(e) {
      if (branchesSearchModalRef.current && !branchesSearchModalRef.current.contains(e.target)) setShowBranchesSearch(false);
    }
    function handleEsc(e) { if (e.key === 'Escape') setShowBranchesSearch(false); }
    document.addEventListener('mousedown', handleClick);
    document.addEventListener('keydown', handleEsc);
    return () => {
      document.removeEventListener('mousedown', handleClick);
      document.removeEventListener('keydown', handleEsc);
    };
  }, [showBranchesSearch]);
  // Search modal close for regions
  useEffect(() => {
    if (!showRegionsSearch) return;
    function handleClick(e) {
      if (regionsSearchModalRef.current && !regionsSearchModalRef.current.contains(e.target)) setShowRegionsSearch(false);
    }
    function handleEsc(e) { if (e.key === 'Escape') setShowRegionsSearch(false); }
    document.addEventListener('mousedown', handleClick);
    document.addEventListener('keydown', handleEsc);
    return () => {
      document.removeEventListener('mousedown', handleClick);
      document.removeEventListener('keydown', handleEsc);
    };
  }, [showRegionsSearch]);

  // Filtered data for branches
  const filteredBranches = currentBranches.filter(b => {
    const term = branchesSearchTerm.toLowerCase();
    return (
      (branchesShowColumns.branch && b.branch.toLowerCase().includes(term)) ||
      (branchesShowColumns.executives && b.executives.toString().includes(term)) ||
      (branchesShowColumns.inRegions && b.inRegions.toString().includes(term))
    );
  });
  // Filtered data for regions
  const filteredRegions = currentRegions.filter(r => {
    const term = regionsSearchTerm.toLowerCase();
    return (
      (regionsShowColumns.region && r.region.toLowerCase().includes(term)) ||
      (regionsShowColumns.branches && r.branches.toString().includes(term))
    );
  });
  // Download CSV for branches
  const handleDownloadBranches = () => {
    const header = [];
    if (branchesShowColumns.branch) header.push('Branch');
    if (branchesShowColumns.executives) header.push('Executives');
    if (branchesShowColumns.inRegions) header.push('In Regions');
    const rows = filteredBranches.map(b => [
      branchesShowColumns.branch ? b.branch : null,
      branchesShowColumns.executives ? b.executives : null,
      branchesShowColumns.inRegions ? b.inRegions : null
    ].filter(x => x !== null).join(','));
    const csv = header.join(',') + '\n' + rows.join('\n');
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'current_branches.csv';
    a.click();
    URL.revokeObjectURL(url);
  };
  // Download CSV for regions
  const handleDownloadRegions = () => {
    const header = [];
    if (regionsShowColumns.region) header.push('Region');
    if (regionsShowColumns.branches) header.push('Branches');
    const rows = filteredRegions.map(r => [
      regionsShowColumns.region ? r.region : null,
      regionsShowColumns.branches ? r.branches : null
    ].filter(x => x !== null).join(','));
    const csv = header.join(',') + '\n' + rows.join('\n');
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'current_regions.csv';
    a.click();
    URL.revokeObjectURL(url);
  };

  // 1. Add dedicated state for Current Mappings table dropdown
  const [showMappingsDropdown, setShowMappingsDropdown] = useState(false);
  const mappingsDropdownRef = useRef();
  const mappingsActionsRef = useRef();

  // 2. Add useEffect to close the dropdown on outside click or ESC
  useEffect(() => {
    if (!showMappingsDropdown) return;
    function handleClick(e) {
      if (
        mappingsDropdownRef.current &&
        !mappingsDropdownRef.current.contains(e.target) &&
        mappingsActionsRef.current &&
        !mappingsActionsRef.current.contains(e.target)
      ) {
        setShowMappingsDropdown(false);
      }
    }
    function handleEsc(e) { if (e.key === 'Escape') setShowMappingsDropdown(false); }
    document.addEventListener('mousedown', handleClick);
    document.addEventListener('keydown', handleEsc);
    return () => {
      document.removeEventListener('mousedown', handleClick);
      document.removeEventListener('keydown', handleEsc);
    };
  }, [showMappingsDropdown]);

  useEffect(() => {
    if (onStatsChange) onStatsChange({ branches: currentBranches.length, regions: currentRegions.length });
  }, [currentBranches, currentRegions, onStatsChange]);

  return (
    <div className="branch-region-mapping-root">
      <h2>Branch & Region Mapping</h2>
      <div className="brm-tabs">
        <button className={activeTab === 'manual' ? 'active' : ''} onClick={() => setActiveTab('manual')}>Manual Entry</button>
        <button className={activeTab === 'file' ? 'active' : ''} onClick={() => setActiveTab('file')}>File Upload</button>
      </div>
      {activeTab === 'manual' && (
        <div className="brm-manual-entry">
          <div className="brm-current-section">
            <div
              className="brm-table-container"
              onMouseEnter={() => setShowActions(true)}
              onMouseLeave={() => setShowActions(false)}
            >
              <h3>Current Branches</h3>
              <div className="em-mapping-header-row">
                <div className={`em-mapping-actions${showActions ? ' visible' : ''}`} ref={branchesActionsRef}>
                  <button className={`em-mapping-action em-eye-action${showBranchesDropdown ? ' active' : ''}`} title="Show/Hide Columns" onClick={() => setShowBranchesDropdown(v => !v)} tabIndex={0}>
                    <FaEye />
                  </button>
                  {showBranchesDropdown && (
                    <div className="em-mapping-dropdown" ref={branchesDropdownRef}>
                      <div className="em-mapping-dropdown-row disabled"><input type="checkbox" disabled /><span>(index)</span></div>
                      <div className="em-mapping-dropdown-row"><input type="checkbox" checked={branchesShowColumns.branch} onChange={() => setBranchesShowColumns(sc => ({ ...sc, branch: !sc.branch }))} /><span>Branch</span></div>
                      <div className="em-mapping-dropdown-row"><input type="checkbox" checked={branchesShowColumns.executives} onChange={() => setBranchesShowColumns(sc => ({ ...sc, executives: !sc.executives }))} /><span>Executives</span></div>
                      <div className="em-mapping-dropdown-row"><input type="checkbox" checked={branchesShowColumns.inRegions} onChange={() => setBranchesShowColumns(sc => ({ ...sc, inRegions: !sc.inRegions }))} /><span>In Regions</span></div>
                    </div>
                  )}
                  <button className="em-mapping-action" title="Zoom" onClick={() => setBranchesZoomed(z => !z)}><FiMaximize /></button>
                  <button className="em-mapping-action" title="Search" onClick={() => setShowBranchesSearch(true)}><FaSearch /></button>
                  <button className="em-mapping-action" title="Download CSV" onClick={handleDownloadBranches}><FaDownload /></button>
                </div>
              </div>
              {branchesZoomed && (
                <div className="em-zoom-modal" onClick={() => setBranchesZoomed(false)}>
                  <div className="em-zoom-modal-content" onClick={e => e.stopPropagation()}>
                    <button className="em-zoom-close" onClick={() => setBranchesZoomed(false)}>&times;</button>
                    <div className="em-zoomed">
                      <table className="brm-table-wide">
                        <thead>
                          <tr>
                            {branchesShowColumns.branch && <th>Branch</th>}
                            {branchesShowColumns.executives && <th>Executives</th>}
                            {branchesShowColumns.inRegions && <th>In Regions</th>}
                          </tr>
                        </thead>
                        <tbody>
                          {filteredBranches.map((b, i) => (
                            <tr key={i}>
                              {branchesShowColumns.branch && <td>{b.branch}</td>}
                              {branchesShowColumns.executives && <td>{b.executives}</td>}
                              {branchesShowColumns.inRegions && <td>{b.inRegions}</td>}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              )}
              {!branchesZoomed && (
                <table className="brm-table-wide">
                  <thead>
                    <tr>
                      {branchesShowColumns.branch && <th>Branch</th>}
                      {branchesShowColumns.executives && <th>Executives</th>}
                      {branchesShowColumns.inRegions && <th>In Regions</th>}
                    </tr>
                  </thead>
                  <tbody>
                    {filteredBranches.map((b, i) => (
                      <tr key={i}>
                        {branchesShowColumns.branch && <td>{b.branch}</td>}
                        {branchesShowColumns.executives && <td>{b.executives}</td>}
                        {branchesShowColumns.inRegions && <td>{b.inRegions}</td>}
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
              {showBranchesSearch && (
                <div className="em-search-modal">
                  <div className="em-search-modal-content" ref={branchesSearchModalRef}>
                    <input className="em-mapping-search em-mapping-search-modal" type="text" placeholder="Search..." value={branchesSearchTerm} autoFocus onChange={e => setBranchesSearchTerm(e.target.value)} />
                    <button className="em-search-close" onClick={() => setShowBranchesSearch(false)} title="Close">&times;</button>
                  </div>
                </div>
              )}
              <div className="brm-remove">
                <label>Remove Branch:</label>
                <select value={removeBranch} onChange={e => setRemoveBranch(e.target.value)}>
                  <option value="">Select Branch</option>
                  {branchesList.map((b, i) => <option key={i} value={b}>{b}</option>)}
                </select>
                <button className="brm-remove-btn" onClick={handleRemoveBranch}>Remove</button>
              </div>
            </div>
            <div
              className="brm-table-container"
              onMouseEnter={() => setShowActions(true)}
              onMouseLeave={() => setShowActions(false)}
            >
              <h3>Current Regions</h3>
              <div className="em-mapping-header-row">
                <div className={`em-mapping-actions${showActions ? ' visible' : ''}`} ref={regionsActionsRef}>
                  <button className={`em-mapping-action em-eye-action${showRegionsDropdown ? ' active' : ''}`} title="Show/Hide Columns" onClick={() => setShowRegionsDropdown(v => !v)} tabIndex={0}>
                    <FaEye />
                  </button>
                  {showRegionsDropdown && (
                    <div className="em-mapping-dropdown" ref={regionsDropdownRef}>
                      <div className="em-mapping-dropdown-row disabled"><input type="checkbox" disabled /><span>(index)</span></div>
                      <div className="em-mapping-dropdown-row"><input type="checkbox" checked={regionsShowColumns.region} onChange={() => setRegionsShowColumns(sc => ({ ...sc, region: !sc.region }))} /><span>Region</span></div>
                      <div className="em-mapping-dropdown-row"><input type="checkbox" checked={regionsShowColumns.branches} onChange={() => setRegionsShowColumns(sc => ({ ...sc, branches: !sc.branches }))} /><span>Branches</span></div>
                    </div>
                  )}
                  <button className="em-mapping-action" title="Zoom" onClick={() => setRegionsZoomed(z => !z)}><FiMaximize /></button>
                  <button className="em-mapping-action" title="Search" onClick={() => setShowRegionsSearch(true)}><FaSearch /></button>
                  <button className="em-mapping-action" title="Download CSV" onClick={handleDownloadRegions}><FaDownload /></button>
                </div>
              </div>
              {regionsZoomed && (
                <div className="em-zoom-modal" onClick={() => setRegionsZoomed(false)}>
                  <div className="em-zoom-modal-content" onClick={e => e.stopPropagation()}>
                    <button className="em-zoom-close" onClick={() => setRegionsZoomed(false)}>&times;</button>
                    <div className="em-zoomed">
                      <table className="brm-table-wide">
                        <thead>
                          <tr>
                            {regionsShowColumns.region && <th>Region</th>}
                            {regionsShowColumns.branches && <th>Branches</th>}
                          </tr>
                        </thead>
                        <tbody>
                          {filteredRegions.map((r, i) => (
                            <tr key={i}>
                              {regionsShowColumns.region && <td>{r.region}</td>}
                              {regionsShowColumns.branches && <td>{r.branches}</td>}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              )}
              {!regionsZoomed && (
                <table className="brm-table-wide">
                  <thead>
                    <tr>
                      {regionsShowColumns.region && <th>Region</th>}
                      {regionsShowColumns.branches && <th>Branches</th>}
                    </tr>
                  </thead>
                  <tbody>
                    {filteredRegions.map((r, i) => (
                      <tr key={i}>
                        {regionsShowColumns.region && <td>{r.region}</td>}
                        {regionsShowColumns.branches && <td>{r.branches}</td>}
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
              {showRegionsSearch && (
                <div className="em-search-modal">
                  <div className="em-search-modal-content" ref={regionsSearchModalRef}>
                    <input className="em-mapping-search em-mapping-search-modal" type="text" placeholder="Search..." value={regionsSearchTerm} autoFocus onChange={e => setRegionsSearchTerm(e.target.value)} />
                    <button className="em-search-close" onClick={() => setShowRegionsSearch(false)} title="Close">&times;</button>
                  </div>
                </div>
              )}
              <div className="brm-remove">
                <label>Remove Region:</label>
                <select value={removeRegion} onChange={e => setRemoveRegion(e.target.value)}>
                  <option value="">Select Region</option>
                  {regionsList.map((r, i) => <option key={i} value={r}>{r}</option>)}
                </select>
                <button className="brm-remove-btn" onClick={handleRemoveRegion}>Remove</button>
              </div>
            </div>
          </div>
          <div className="brm-map-exec-branch">
            <h3>Map Executives to Branches</h3>
            <label>Select Branch:</label>
            <select value={selectedBranch} onChange={e => setSelectedBranch(e.target.value)}>
              {branchesList.map((b, i) => <option key={i} value={b}>{b}</option>)}
            </select>
            <label>Branch Executives:</label>
            <div className="brm-multiselect">
              {selectedExecutives.map((exec, i) => (
                <span key={i} className="selected">
                  {exec} <span className="close" onClick={() => handleRemoveExecutive(exec)}>×</span>
                </span>
              ))}
            </div>
            <label>Add Executive:</label>
            <select
              value=""
              onChange={e => {
                const exec = e.target.value;
                if (exec && !selectedExecutives.includes(exec)) setSelectedExecutives([...selectedExecutives, exec]);
              }}
            >
              <option value="">Select Executive</option>
              {executivesList.filter(exec => !selectedExecutives.includes(exec)).map((exec, i) => (
                <option key={i} value={exec}>{exec}</option>
              ))}
            </select>
            <button onClick={handleUpdateBranchMapping}>Update Branch Mapping</button>
          </div>
          <div
            className={`brm-current-mappings${zoomed ? ' brm-zoomed' : ''}`}
            onMouseEnter={() => setShowActions(true)}
            onMouseLeave={() => { setShowActions(false); }}
          >
            <div className="em-mapping-header-row">
              <h3>Current Mappings</h3>
              <div className={`em-mapping-actions${showActions ? ' visible' : ''}`} ref={mappingsActionsRef}>
                <button
                  className={`em-mapping-action em-eye-action${showMappingsDropdown ? ' active' : ''}`}
                  title="Show/Hide Columns"
                  onClick={() => setShowMappingsDropdown(v => !v)}
                  tabIndex={0}
                >
                  <FaEye />
                </button>
                {showMappingsDropdown && (
                  <div className="em-mapping-dropdown" ref={mappingsDropdownRef}>
                    <div className="em-mapping-dropdown-row disabled"><input type="checkbox" disabled /><span>(index)</span></div>
                    {COLUMN_OPTIONS.map(col => (
                      <div key={col.key} className="em-mapping-dropdown-row">
                        <input
                          type="checkbox"
                          checked={showColumns[col.key]}
                          onChange={() => setShowColumns(sc => ({ ...sc, [col.key]: !sc[col.key] }))}
                        />
                        <span>{col.label}</span>
                      </div>
                    ))}
                  </div>
                )}
                <button className="em-mapping-action" title="Zoom" onClick={() => setZoomed(z => !z)}>
                  <FiMaximize />
                </button>
                <button className="em-mapping-action" title="Search" onClick={() => setShowSearch(true)}>
                  <FaSearch />
                </button>
                <button className="em-mapping-action" title="Download CSV" onClick={handleDownload}>
                  <FaDownload />
                </button>
              </div>
            </div>
            {zoomed && (
              <div className="em-zoom-modal" onClick={() => setZoomed(false)}>
                <div className="em-zoom-modal-content" onClick={e => e.stopPropagation()}>
                  <button className="em-zoom-close" onClick={() => setZoomed(false)}>&times;</button>
                  <div className="em-zoomed">
                    <table>
                      <thead>
                        <tr>
                          {COLUMN_OPTIONS.map(col => showColumns[col.key] && <th key={col.key}>{col.label}</th>)}
                        </tr>
                      </thead>
                      <tbody>
                        {filteredMappings.map((row, idx) => (
                          <tr key={idx}>
                            {COLUMN_OPTIONS.map(col => showColumns[col.key] && <td key={col.key}>{row[col.key]}</td>)}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}
            {!zoomed && (
              <table>
                <thead>
                  <tr>
                    {COLUMN_OPTIONS.map(col => showColumns[col.key] && <th key={col.key}>{col.label}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {filteredMappings.map((row, idx) => (
                    <tr key={idx}>
                      {COLUMN_OPTIONS.map(col => showColumns[col.key] && <td key={col.key}>{row[col.key]}</td>)}
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
            {showSearch && (
              <div className="em-search-modal">
                <div className="em-search-modal-content" ref={searchModalRef}>
                  <input
                    className="em-mapping-search em-mapping-search-modal"
                    type="text"
                    placeholder="Search..."
                    value={searchTerm}
                    autoFocus
                    onChange={e => setSearchTerm(e.target.value)}
                  />
                  <button className="em-search-close" onClick={() => setShowSearch(false)} title="Close">&times;</button>
                </div>
              </div>
            )}
          </div>
          <div className="brm-mapping-section">
            <div className="brm-map-branch-region">
              <h3>Map Branches to Regions</h3>
              <label>Select Region:</label>
              <select value={selectedRegion} onChange={e => setSelectedRegion(e.target.value)}>
                {regionsList.map((r, i) => <option key={i} value={r}>{r}</option>)}
              </select>
              <label>Region Branches:</label>
              <div className="brm-multiselect">
                {selectedBranches.map((branch, i) => (
                  <span key={i} className="selected">
                    {branch} <span className="close" onClick={() => handleRemoveBranchFromRegion(branch)}>×</span>
                  </span>
                ))}
              </div>
              <label>Add Branch:</label>
              <select
                value=""
                onChange={e => {
                  const branch = e.target.value;
                  if (branch && !selectedBranches.includes(branch)) setSelectedBranches([...selectedBranches, branch]);
                }}
              >
                <option value="">Select Branch</option>
                {branchesList.filter(branch => !selectedBranches.includes(branch)).map((branch, i) => (
                  <option key={i} value={branch}>{branch}</option>
                ))}
              </select>
              <button onClick={handleUpdateRegionMapping}>Update Region Mapping</button>
            </div>
          </div>
        </div>
      )}
      {activeTab === 'file' && (
        <div className="brm-file-upload">
          <h3>Upload Branch & Region Mapping File <span className="brm-link-icon">↔</span></h3>
          <div className="brm-file-upload-sub">Upload Branch-Region File</div>
          <div
            className={`brm-upload-dropzone${isDragging ? ' dragging' : ''}`}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current && fileInputRef.current.click()}
          >
            <div className="brm-upload-dropzone-content">
              <span className="brm-upload-icon">&#128230;</span>
              <span className="brm-upload-texts">
                <span className="brm-upload-main">Drag and drop file here</span>
                <span className="brm-upload-info">Limit 200MB per file • XLSX, XLS</span>
              </span>
            </div>
            <input
              type="file"
              ref={fileInputRef}
              style={{ display: 'none' }}
              accept=".xls,.xlsx"
              onChange={handleFileInput}
            />
            <button className="brm-upload-browse" type="button" onClick={e => { e.stopPropagation(); fileInputRef.current && fileInputRef.current.click(); }}>Browse files</button>
          </div>
        </div>
      )}
    </div>
  );
};

export default BranchRegionMapping; 