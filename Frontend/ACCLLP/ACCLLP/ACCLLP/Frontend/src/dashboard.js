import React, { useState } from 'react';
import './dashboard.css';
import ExecutiveManagement from './components/ExecutiveManagement';
import BranchRegionMapping from './components/BranchRegionMapping';
import CompanyProductMapping from './components/CompanyProductMapping';
import Sidebar from './components/Sidebar';
import BackupRestore from './components/BackupRestore';
import ConsolidatedDataView from './components/ConsolidatedDataView';
import FileProcessing from './components/FileProcessing';

const Dashboard = ({ onLogout }) => {
  const [activeNav, setActiveNav] = useState('executiveManagement');
  const [sidebarVisible, setSidebarVisible] = useState(true);

  // LIFTED STATE: Real-time stats
  const [stats, setStats] = useState({
    executives: 44,
    branches: 11,
    regions: 3,
    companies: 4,
    products: 31
  });

  // Handler to update stats from any module
  const handleStatsChange = (updates) => {
    setStats(prev => ({ ...prev, ...updates }));
  };

  const handleSaveAllMappings = () => {
    // Implement save all mappings logic
    alert('Save All Mappings clicked!');
  };

  const renderCompanyProductMapping = () => (
    <div className="management-section">
      <h2>Company & Product Mapping</h2>
      <p>Content for Company & Product Mapping will go here</p>
    </div>
  );

  const renderBackupRestore = () => (
    <div className="management-section">
      <h2>Backup & Restore</h2>
      <p>Content for Backup & Restore will go here</p>
    </div>
  );

  const renderConsolidatedDataView = () => (
    <div className="management-section">
      <h2>Consolidated Data View</h2>
      <p>Content for Consolidated Data View will go here</p>
    </div>
  );

  const renderFileProcessing = () => (
    <div className="management-section">
      <h2>File Processing</h2>
      <p>Content for File Processing will go here</p>
    </div>
  );

  return (
    <div className="dashboard-layout">
      {sidebarVisible && (
        <Sidebar
          visible={sidebarVisible}
          onToggle={() => setSidebarVisible(v => !v)}
          stats={stats}
          onSaveAllMappings={handleSaveAllMappings}
          onLogout={onLogout}
        />
      )}
      <div className={`dashboard-main${sidebarVisible ? '' : ' expanded'}`}>
        {!sidebarVisible && (
          <button
            className="sidebar-toggle-inside"
            onClick={() => setSidebarVisible(true)}
            aria-label="Show sidebar"
          >
            <span className="arrow">&#x25B6;</span>
          </button>
        )}
        <div className="executive-portal">
          <h1>Executive Mapping Administration Portal</h1>
          <div className="portal-nav">
            <button 
              className={`nav-btn ${activeNav === 'executiveManagement' ? 'active' : ''}`}
              onClick={() => setActiveNav('executiveManagement')}
            >
              Executive Management
            </button>
            <button 
              className={`nav-btn ${activeNav === 'branchRegionMapping' ? 'active' : ''}`}
              onClick={() => setActiveNav('branchRegionMapping')}
            >
              Branch & Region Mapping
            </button>
            <button 
              className={`nav-btn ${activeNav === 'companyProductMapping' ? 'active' : ''}`}
              onClick={() => setActiveNav('companyProductMapping')}
            >
              Company & Product Mapping
            </button>
            <button 
              className={`nav-btn ${activeNav === 'backupRestore' ? 'active' : ''}`}
              onClick={() => setActiveNav('backupRestore')}
            >
              Backup & Restore
            </button>
            <button 
              className={`nav-btn ${activeNav === 'consolidatedDataView' ? 'active' : ''}`}
              onClick={() => setActiveNav('consolidatedDataView')}
            >
              Consolidated Data View
            </button>
            <button 
              className={`nav-btn ${activeNav === 'fileProcessing' ? 'active' : ''}`}
              onClick={() => setActiveNav('fileProcessing')}
            >
              File Processing
            </button>
          </div>

          {activeNav === 'executiveManagement' && (
            <ExecutiveManagement onStatsChange={handleStatsChange} />
          )}
          {activeNav === 'branchRegionMapping' && (
            <BranchRegionMapping onStatsChange={handleStatsChange} />
          )}
          {activeNav === 'companyProductMapping' && (
            <CompanyProductMapping onStatsChange={handleStatsChange} />
          )}
          {activeNav === 'backupRestore' && <BackupRestore />}
          {activeNav === 'consolidatedDataView' && <ConsolidatedDataView />}
          {activeNav === 'fileProcessing' && <FileProcessing />}
        </div>
      </div>
    </div>
  );
};

export default Dashboard;