import React, { useState, useRef, useEffect } from 'react';
import '../dashboard.css';
import './CompanyProductMapping.css';
import { FaEye, FaSearch, FaTimes, FaTrash } from 'react-icons/fa';
import { FiMaximize } from 'react-icons/fi';

const defaultProducts = [
  { name: 'NITRIC ACID', companies: 1 },
  { name: 'PEROXIDE', companies: 1 },
  { name: 'PROBIOTIC', companies: 1 },
  { name: 'SAFOLITE', companies: 1 },
  { name: 'SALT', companies: 1 },
  { name: 'SBC', companies: 1 },
  { name: 'SBP', companies: 1 }
];
const defaultCompanies = [
  { name: 'CP', products: 5 },
  { name: 'GENEREL', products: 15 },
  { name: 'SOLVENTS', products: 5 },
  { name: 'TATA', products: 6 }
];
const defaultMappings = [
  { company: 'TATA', products: ['PROBIOTIC', 'SBC', 'SILICA', 'SODA ASH', 'SODA ASH DENSE', 'STPP'] },
  { company: 'SOLVENTS', products: ['ACETIC ACID GLACIAL', 'ETHYL ACETATE', 'ISOPROPYL ALCOHOL', 'MDC', 'SOLVENTS'] },
  { company: 'CP', products: ['CSF', 'CSL', 'HCL', 'HYPO', 'PEROXIDE'] },
  { company: 'GENEREL', products: ['AUXILIARIES', 'DAIRY', 'DISCHARGE', 'DYES', 'EMPTY CARBOY', 'FORMALDEHYDE', 'FORMIC ACID', 'GC', 'NITRIC ACID', 'SAFOLITE', 'SALT', 'SBP', 'SILICA', 'SODA ASH', 'SODA ASH DENSE', 'STPP'] }
];

const CompanyProductMapping = ({ onStatsChange = () => {} }) => {
  // Tabs: Manual Entry / File Upload
  const [activeTab, setActiveTab] = useState('manual');
  // Product/Company creation
  const [productName, setProductName] = useState('');
  const [companyName, setCompanyName] = useState('');
  const [products, setProducts] = useState([...defaultProducts]);
  const [companies, setCompanies] = useState([...defaultCompanies]);
  // Remove
  const [removeProduct, setRemoveProduct] = useState('');
  const [removeCompany, setRemoveCompany] = useState('');
  // Mapping
  const [mappings, setMappings] = useState([...defaultMappings]);
  const [selectedCompany, setSelectedCompany] = useState('TATA');
  const [selectedProducts, setSelectedProducts] = useState(['SODA ASH', 'SBC', 'SILICA', 'PROBIOTIC', 'SODA ASH DENSE', 'STPP']);
  // Table controls
  const [showColumns, setShowColumns] = useState({ name: true, companies: true });
  const [showCompanyColumns, setShowCompanyColumns] = useState({ name: true, products: true, count: true });
  const [showActions, setShowActions] = useState(false);
  const [showDropdown, setShowDropdown] = useState(false);
  const [showSearch, setShowSearch] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [zoomed, setZoomed] = useState(false);
  // Products table controls
  const [showProductActions, setShowProductActions] = useState(false);
  const [showProductDropdown, setShowProductDropdown] = useState(false);
  const [showProductSearch, setShowProductSearch] = useState(false);
  const [productSearchQuery, setProductSearchQuery] = useState('');
  const [productZoomed, setProductZoomed] = useState(false);
  const [showProductColumns, setShowProductColumns] = useState({
    name: true,
    companies: true
  });
  // Refs
  const actionsRef = useRef();
  const dropdownRef = useRef();
  const searchModalRef = useRef();
  const productActionsRef = useRef();
  const productDropdownRef = useRef();
  const productSearchModalRef = useRef();
  // Add state for file upload
  const [selectedFile, setSelectedFile] = useState(null);
  const fileInputRef = useRef();
  const [uploadMessage, setUploadMessage] = useState('');
  // Add these states and refs near the other useState/useRef hooks
  const [showCompanyGroupActions, setShowCompanyGroupActions] = useState(false);
  const [showCompanyGroupDropdown, setShowCompanyGroupDropdown] = useState(false);
  const [showCompanyGroupSearch, setShowCompanyGroupSearch] = useState(false);
  const [companyGroupSearchQuery, setCompanyGroupSearchQuery] = useState('');
  const [companyGroupZoomed, setCompanyGroupZoomed] = useState(false);
  const [showCompanyGroupColumns, setShowCompanyGroupColumns] = useState({
    name: true,
    products: true
  });
  const companyGroupActionsRef = useRef();
  const companyGroupDropdownRef = useRef();
  const companyGroupSearchModalRef = useRef();

  // Dropdown and search modal close for main table
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

  // Dropdown and search modal close for products table
  useEffect(() => {
    if (!showProductDropdown) return;
    function handleClick(e) {
      if (
        productDropdownRef.current &&
        !productDropdownRef.current.contains(e.target) &&
        productActionsRef.current &&
        !productActionsRef.current.contains(e.target)
      ) {
        setShowProductDropdown(false);
      }
    }
    function handleEsc(e) { if (e.key === 'Escape') setShowProductDropdown(false); }
    document.addEventListener('mousedown', handleClick);
    document.addEventListener('keydown', handleEsc);
    return () => {
      document.removeEventListener('mousedown', handleClick);
      document.removeEventListener('keydown', handleEsc);
    };
  }, [showProductDropdown]);

  useEffect(() => {
    if (!showProductSearch) return;
    function handleClick(e) {
      if (productSearchModalRef.current && !productSearchModalRef.current.contains(e.target)) setShowProductSearch(false);
    }
    function handleEsc(e) { if (e.key === 'Escape') setShowProductSearch(false); }
    document.addEventListener('mousedown', handleClick);
    document.addEventListener('keydown', handleEsc);
    return () => {
      document.removeEventListener('mousedown', handleClick);
      document.removeEventListener('keydown', handleEsc);
    };
  }, [showProductSearch]);

  // Product/Company creation
  const handleCreateProduct = (e) => {
    e.preventDefault();
    if (productName && !products.find(p => p.name === productName)) {
      setProducts([...products, { name: productName, companies: 0 }]);
      setProductName('');
    }
  };

  const handleCreateCompany = (e) => {
    e.preventDefault();
    if (companyName && !companies.find(c => c.name === companyName)) {
      setCompanies([...companies, { name: companyName, products: 0 }]);
      setCompanyName('');
    }
  };

  // Remove
  const handleRemoveProduct = () => {
    setProducts(products.filter(p => p.name !== removeProduct));
    setRemoveProduct('');
  };

  const handleRemoveCompany = () => {
    setCompanies(companies.filter(c => c.name !== removeCompany));
    setRemoveCompany('');
  };

  // Mapping
  const handleUpdateCompanyMapping = () => {
    setMappings(mappings => {
      const filtered = mappings.filter(m => m.company !== selectedCompany);
      return [...filtered, { company: selectedCompany, products: [...selectedProducts] }];
    });
  };

  // Download CSV
  const handleDownloadProducts = () => {
    const header = ['Product', 'Companies'].join(',') + '\n';
    const rows = products.map(p => [p.name, p.companies].join(','));
    const csv = header + rows.join('\n');
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'products.csv';
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleDownload = () => {
    const header = ['Company', 'Products', 'Count'].join(',') + '\n';
    const rows = mappings.map(m => [m.company, '"' + m.products.join(', ') + '"', m.products.length].join(','));
    const csv = header + rows.join('\n');
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'company_product_mappings.csv';
    a.click();
    URL.revokeObjectURL(url);
  };

  // Add this download function for company groups
  const handleDownloadCompanyGroups = () => {
    const header = ['Company', 'Products'].join(',') + '\n';
    const rows = companies.map(c => [c.name, c.products].join(','));
    const csv = header + rows.join('\n');
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'company_groups.csv';
    a.click();
    URL.revokeObjectURL(url);
  };

  // Table filters
  const filteredProducts = products.filter(p => p.name.toLowerCase().includes(productSearchQuery.toLowerCase()));
  const filteredCompanies = companies.filter(c => c.name.toLowerCase().includes(searchQuery.toLowerCase()));
  const filteredMappings = mappings.filter(m =>
    m.company.toLowerCase().includes(searchQuery.toLowerCase()) ||
    m.products.join(', ').toLowerCase().includes(searchQuery.toLowerCase())
  );

  // Products table UI
  const productsTableContent = (
    <>
      <div className="em-mapping-header-row">
        <h3>Current Products</h3>
        {!productZoomed && (
          <div className={`em-mapping-actions${showProductActions ? ' visible' : ''}`} ref={productActionsRef}>
            <button
              className={`em-mapping-action em-eye-action${showProductDropdown ? ' active' : ''}`}
              title="Show/Hide Columns"
              onClick={() => setShowProductDropdown(v => !v)}
              tabIndex={0}
            >
              <FaEye />
            </button>
            {showProductDropdown && (
              <div className="em-mapping-dropdown" ref={productDropdownRef}>
                <div className="em-mapping-dropdown-row">
                  <input
                    type="checkbox"
                    checked={showProductColumns.name}
                    onChange={() => setShowProductColumns(sc => ({ ...sc, name: !sc.name }))}
                  />
                  <span>Product Name</span>
                </div>
                <div className="em-mapping-dropdown-row">
                  <input
                    type="checkbox"
                    checked={showProductColumns.companies}
                    onChange={() => setShowProductColumns(sc => ({ ...sc, companies: !sc.companies }))}
                  />
                  <span>Companies</span>
                </div>
              </div>
            )}
            <button className="em-mapping-action" title="Zoom" onClick={() => setProductZoomed(z => !z)}>
              <FiMaximize />
            </button>
            <button className="em-mapping-action" title="Search" onClick={() => setShowProductSearch(true)}>
              <FaSearch />
            </button>
            <button className="em-mapping-action" title="Download CSV" onClick={handleDownloadProducts}>
              <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#aaa" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
            </button>
          </div>
        )}
      </div>
      <table>
        <thead>
          <tr>
            {showProductColumns.name && <th>Product Name</th>}
            {showProductColumns.companies && <th>Companies</th>}
          </tr>
        </thead>
        <tbody>
          {filteredProducts.map((p, idx) => (
            <tr key={idx}>
              {showProductColumns.name && <td>{p.name}</td>}
              {showProductColumns.companies && <td>{p.companies}</td>}
            </tr>
          ))}
        </tbody>
      </table>
      <div className="remove-executive-section">
        <label htmlFor="remove-product-select" style={{ fontWeight: 500, marginBottom: 6 }}>Remove Product:</label>
        <select
          id="remove-product-select"
          className="remove-exec-dropdown"
          value={removeProduct}
          onChange={e => setRemoveProduct(e.target.value)}
          style={{ display: 'block', width: '100%', marginBottom: 12, marginTop: 2 }}
        >
          <option value="">Select Product</option>
          {products.map((p, idx) => (
            <option key={idx} value={p.name}>{p.name}</option>
          ))}
        </select>
        <button
          className="remove-executive-btn remove-executive-btn-block"
          style={{ marginTop: 8, width: 'fit-content', display: 'flex', alignItems: 'center', gap: 8 }}
          disabled={!removeProduct}
          onClick={handleRemoveProduct}
        >
          <FaTrash style={{ marginRight: 4 }} /> Remove Product
        </button>
      </div>
    </>
  );

  // Main table UI
  const tableContent = (
    <>
      <div className="em-mapping-header-row">
        <h3>Current Mappings</h3>
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
                    checked={showCompanyColumns.name}
                    onChange={() => setShowCompanyColumns(sc => ({ ...sc, name: !sc.name }))}
                  />
                  <span>Company</span>
                </div>
                <div className="em-mapping-dropdown-row">
                  <input
                    type="checkbox"
                    checked={showCompanyColumns.products}
                    onChange={() => setShowCompanyColumns(sc => ({ ...sc, products: !sc.products }))}
                  />
                  <span>Products</span>
                </div>
                <div className="em-mapping-dropdown-row">
                  <input
                    type="checkbox"
                    checked={showCompanyColumns.count}
                    onChange={() => setShowCompanyColumns(sc => ({ ...sc, count: !sc.count }))}
                  />
                  <span>Count</span>
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
            {showCompanyColumns.name && <th>Company</th>}
            {showCompanyColumns.products && <th>Products</th>}
            {showCompanyColumns.count && <th>Count</th>}
          </tr>
        </thead>
        <tbody>
          {filteredMappings.map((m, idx) => (
            <tr key={idx}>
              {showCompanyColumns.name && <td>{m.company}</td>}
              {showCompanyColumns.products && <td>{m.products.join(', ')}</td>}
              {showCompanyColumns.count && <td>{m.products.length}</td>}
            </tr>
          ))}
        </tbody>
      </table>
    </>
  );

  // Add this filtered list
  const filteredCompanyGroups = companies.filter(c => c.name.toLowerCase().includes(companyGroupSearchQuery.toLowerCase()));

  // Add this table content block
  const companyGroupTableContent = (
    <>
      <div className="em-mapping-header-row">
        <h3>Current Company Groups</h3>
        {!companyGroupZoomed && (
          <div className={`em-mapping-actions${showCompanyGroupActions ? ' visible' : ''}`} ref={companyGroupActionsRef}>
            <button
              className={`em-mapping-action em-eye-action${showCompanyGroupDropdown ? ' active' : ''}`}
              title="Show/Hide Columns"
              onClick={() => setShowCompanyGroupDropdown(v => !v)}
              tabIndex={0}
            >
              <FaEye />
            </button>
            {showCompanyGroupDropdown && (
              <div className="em-mapping-dropdown" ref={companyGroupDropdownRef}>
                <div className="em-mapping-dropdown-row">
                  <input
                    type="checkbox"
                    checked={showCompanyGroupColumns.name}
                    onChange={() => setShowCompanyGroupColumns(sc => ({ ...sc, name: !sc.name }))}
                  />
                  <span>Company</span>
                </div>
                <div className="em-mapping-dropdown-row">
                  <input
                    type="checkbox"
                    checked={showCompanyGroupColumns.products}
                    onChange={() => setShowCompanyGroupColumns(sc => ({ ...sc, products: !sc.products }))}
                  />
                  <span>Products</span>
                </div>
              </div>
            )}
            <button className="em-mapping-action" title="Zoom" onClick={() => setCompanyGroupZoomed(z => !z)}>
              <FiMaximize />
            </button>
            <button className="em-mapping-action" title="Search" onClick={() => setShowCompanyGroupSearch(true)}>
              <FaSearch />
            </button>
            <button className="em-mapping-action" title="Download CSV" onClick={handleDownloadCompanyGroups}>
              <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#aaa" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
            </button>
          </div>
        )}
      </div>
      <table>
        <thead>
          <tr>
            {showCompanyGroupColumns.name && <th>Company</th>}
            {showCompanyGroupColumns.products && <th>Products</th>}
          </tr>
        </thead>
        <tbody>
          {filteredCompanyGroups.map((c, idx) => (
            <tr key={idx}>
              {showCompanyGroupColumns.name && <td>{c.name}</td>}
              {showCompanyGroupColumns.products && <td>{c.products}</td>}
            </tr>
          ))}
        </tbody>
      </table>
      <div className="remove-executive-section">
        <label htmlFor="remove-company-select" style={{ fontWeight: 500, marginBottom: 6 }}>Remove Company:</label>
        <select
          id="remove-company-select"
          className="remove-exec-dropdown"
          value={removeCompany}
          onChange={e => setRemoveCompany(e.target.value)}
          style={{ display: 'block', width: '100%', marginBottom: 12, marginTop: 2 }}
        >
          <option value="">Select Company</option>
          {companies.map((c, idx) => (
            <option key={idx} value={c.name}>{c.name}</option>
          ))}
        </select>
        <button
          className="remove-executive-btn remove-executive-btn-block"
          style={{ marginTop: 8, width: 'fit-content', display: 'flex', alignItems: 'center', gap: 8 }}
          disabled={!removeCompany}
          onClick={handleRemoveCompany}
        >
          <FaTrash style={{ marginRight: 4 }} /> Remove Company
        </button>
      </div>
      {showCompanyGroupSearch && (
        <div className="em-search-modal">
          <div className="em-search-modal-content" ref={companyGroupSearchModalRef}>
            <input
              className="em-mapping-search em-mapping-search-modal"
              type="text"
              placeholder="Search..."
              value={companyGroupSearchQuery}
              autoFocus
              onChange={e => setCompanyGroupSearchQuery(e.target.value)}
            />
            <button className="em-search-close" onClick={() => setShowCompanyGroupSearch(false)} title="Close">&times;</button>
          </div>
        </div>
      )}
    </>
  );

  useEffect(() => {
    if (onStatsChange) onStatsChange({ companies: companies.length, products: products.length });
  }, [companies, products, onStatsChange]);

  return (
    <div className="management-section">
      <h2>Company & Product Mapping</h2>
      <div className="management-tabs">
        <button
          className={`tab-btn ${activeTab === 'manual' ? 'active' : ''}`}
          onClick={() => setActiveTab('manual')}
        >
          Manual Entry
        </button>
        <button
          className={`tab-btn ${activeTab === 'file' ? 'active' : ''}`}
          onClick={() => setActiveTab('file')}
        >
          File Upload
        </button>
      </div>
      {activeTab === 'manual' ? (
        <>
          <div className="add-executive">
            <h3>Create Product Group</h3>
            <form onSubmit={handleCreateProduct}>
              <div className="form-group">
                <label>Product Group Name:</label>
                <input
                  type="text"
                  value={productName}
                  onChange={e => setProductName(e.target.value)}
                />
              </div>
              <button type="submit" className="add-btn">Create Product</button>
            </form>
            
            {/* Products Table */}
            <div 
              className="em-current-mappings" 
              onMouseEnter={() => setShowProductActions(true)} 
              onMouseLeave={() => { 
                setShowProductActions(false); 
                setShowProductDropdown(false); 
              }}
              style={{ marginTop: '20px' }}
            >
              {productZoomed && (
                <div className="em-zoom-modal">
                  <div className="em-zoom-modal-content">
                    <button className="em-zoom-close" onClick={() => setProductZoomed(false)} title="Close">&times;</button>
                    {productsTableContent}
                  </div>
                </div>
              )}
              {!productZoomed && productsTableContent}
              {showProductSearch && (
                <div className="em-search-modal">
                  <div className="em-search-modal-content" ref={productSearchModalRef}>
                    <input
                      className="em-mapping-search em-mapping-search-modal"
                      type="text"
                      placeholder="Search..."
                      value={productSearchQuery}
                      autoFocus
                      onChange={e => setProductSearchQuery(e.target.value)}
                    />
                    <button className="em-search-close" onClick={() => setShowProductSearch(false)} title="Close">&times;</button>
                  </div>
                </div>
              )}
            </div>
          </div>

          <div className="add-executive">
            <h3>Create Company Group</h3>
            <form onSubmit={handleCreateCompany}>
              <div className="form-group">
                <label>Company Group Name:</label>
                <input
                  type="text"
                  value={companyName}
                  onChange={e => setCompanyName(e.target.value)}
                />
              </div>
              <button type="submit" className="add-btn">Create Company</button>
            </form>
          </div>

          <div 
            className="em-current-mappings"
            onMouseEnter={() => setShowCompanyGroupActions(true)}
            onMouseLeave={() => { setShowCompanyGroupActions(false); setShowCompanyGroupDropdown(false); }}
            style={{ marginTop: '20px' }}
          >
            {companyGroupZoomed && (
              <div className="em-zoom-modal">
                <div className="em-zoom-modal-content">
                  <button className="em-zoom-close" onClick={() => setCompanyGroupZoomed(false)} title="Close">&times;</button>
                  {companyGroupTableContent}
                </div>
              </div>
            )}
            {!companyGroupZoomed && companyGroupTableContent}
          </div>

          <div className="em-current-mappings" onMouseEnter={() => setShowActions(true)} onMouseLeave={() => { setShowActions(false); setShowDropdown(false); }}>
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

          <div className="mapping-section">
            <h3>Map Products to Companies</h3>
            <div className="form-group">
              <label>Select Company:</label>
              <select
                value={selectedCompany}
                onChange={e => {
                  setSelectedCompany(e.target.value);
                  const found = mappings.find(m => m.company === e.target.value);
                  setSelectedProducts(found ? found.products : []);
                }}
              >
                {companies.map((c, idx) => (
                  <option key={idx} value={c.name}>{c.name}</option>
                ))}
              </select>
            </div>
            <div className="form-group">
              <label>Select Products:</label>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
                {products.map((p, idx) => (
                  <div
                    key={idx}
                    style={{
                      background: selectedProducts.includes(p.name) ? '#e74c3c' : '#222',
                      color: '#fff',
                      padding: '4px 12px',
                      borderRadius: 16,
                      cursor: 'pointer',
                      marginBottom: 4
                    }}
                    onClick={() => {
                      setSelectedProducts(sel =>
                        sel.includes(p.name)
                          ? sel.filter(n => n !== p.name)
                          : [...sel, p.name]
                      );
                    }}
                  >
                    {p.name}
                    {selectedProducts.includes(p.name) && (
                      <FaTimes style={{ marginLeft: 6, fontSize: 12 }} />
                    )}
                  </div>
                ))}
              </div>
            </div>
            <button className="add-btn" style={{ marginTop: 12 }} onClick={handleUpdateCompanyMapping}>
              Update Company Mapping
            </button>
          </div>
        </>
      ) : (
        <div className="bulk-assignment">
          <h3>Bulk Company-Product Assignment</h3>
          <p>Upload Company-Product File</p>
          <div className="drop-zone">
            <input
              type="file"
              id="file-upload"
              style={{ display: 'none' }}
              ref={fileInputRef}
              onChange={e => {
                setSelectedFile(e.target.files[0]);
                setUploadMessage('');
              }}
            />
            <label htmlFor="file-upload" className="drop-label" style={{ cursor: 'pointer' }}>
              <div className="drop-icon">[ ]</div>
              <p>Drag and drop file here</p>
              <p className="file-limit">Limit 200MB per file - XLS, XLSX</p>
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
                    onClick={e => {
                      e.preventDefault();
                      setUploadMessage('File uploaded successfully!');
                      setTimeout(() => setUploadMessage(''), 2000);
                    }}
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
    </div>
  );
};

export default CompanyProductMapping;