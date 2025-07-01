import React from 'react';
import './Sidebar.css';

const Sidebar = ({
  visible,
  onToggle,
  stats = {},
  onSaveAllMappings,
  onLogout
}) => {
  return (
    <div className={`sidebar-root${visible ? '' : ' collapsed'}`}>
      <button className="sidebar-toggle" onClick={onToggle} aria-label="Toggle sidebar">
        <span className={`arrow${visible ? '' : ' collapsed'}`}>{visible ? '\u25C0' : '\u25B6'}</span>
      </button>
      {visible && (
        <>
          <div className="sidebar-section global-ops">
            <h4>Global Operations</h4>
            <button className="save-btn" onClick={onSaveAllMappings}>Save All Mappings</button>
          </div>
          <div className="sidebar-section stats">
            <h4>System Statistics</h4>
            <div className="stat-row"><span>Executives</span><span>{stats.executives ?? '--'}</span></div>
            <div className="stat-row"><span>Branches</span><span>{stats.branches ?? '--'}</span></div>
            <div className="stat-row"><span>Regions</span><span>{stats.regions ?? '--'}</span></div>
            <div className="stat-row"><span>Companies</span><span>{stats.companies ?? '--'}</span></div>
            <div className="stat-row"><span>Products</span><span>{stats.products ?? '--'}</span></div>
          </div>
          <div className="sidebar-section logout-section">
            <button className="logout-btn" onClick={onLogout}>
              Logout
            </button>
          </div>
        </>
      )}
    </div>
  );
};

export default Sidebar; 