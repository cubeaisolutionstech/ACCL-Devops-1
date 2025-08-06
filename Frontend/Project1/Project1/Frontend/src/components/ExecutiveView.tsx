import React, { useState } from 'react';
import { Button } from '@/components/ui/button';
import { useFileUpload } from '@/context/FileUploadContext';
import { X, Check } from 'lucide-react';

const ExecutiveView: React.FC = () => {
  const { uploadedFiles } = useFileUpload();
  const [activeTab, setActiveTab] = useState('number-of-billed-customer');
  const [selectedAnalysis, setSelectedAnalysis] = useState('');
  const [activeSubTab, setActiveSubTab] = useState('billed-target');

  const files = [
    { name: 'Current Year Sales File', key: 'currentYearSales' },
    { name: 'Last Year Sales File', key: 'lastYearSales' },
    { name: 'Budget File', key: 'budget' },
    { name: 'OS-Previous Month File', key: 'osPrevious' },
    { name: 'OS-Current Month File', key: 'osCurrent' },
  ];

  const reportTabs = [
    { id: 'budget-vs-billed', label: 'Budget vs Billed' },
    { id: 'od-target-vs-collection', label: 'Od Target vs Collection' },
    { id: 'product-growth', label: 'Product Growth' },
    { id: 'number-of-billed-customer', label: 'Number of Billed customer & OD Target' }
  ];

  const subTabs = [
    { id: 'billed-target', label: 'Billed Target' },
    { id: 'od-target', label: 'OD Target' }
  ];

  return (
    <div className="container mx-auto px-4 py-8">
      <div className="text-center mb-8">
        <h1 className="text-2xl font-bold text-blue-600 mb-8 text-center">Integrated Reports Dashboard</h1>
        <div className="w-full max-w-4xl mx-auto">
          {/* Required Files Section */}
          <div className="bg-white rounded-lg border border-gray-200 p-6 shadow-sm mb-6">
            <h2 className="text-lg font-semibold text-blue-600 mb-4">Required Files:</h2>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              {files.map((file, index) => {
                const isUploaded = uploadedFiles[file.key as keyof typeof uploadedFiles];
                return (
                  <div key={index} className="flex items-center space-x-3">
                    <div className={`flex-shrink-0 ${
                      !isUploaded ? 'text-red-500' : 'text-green-500'
                    }`}>
                      {!isUploaded ? (
                        <X className="h-5 w-5" />
                      ) : (
                        <Check className="h-5 w-5" />
                      )}
                    </div>
                    <span className="text-sm text-gray-700">{file.name}:</span>
                    <span className={`text-sm font-medium ${
                      !isUploaded ? 'text-red-600' : 'text-green-600'
                    }`}>
                      {!isUploaded ? 'Missing' : 'Uploaded'}
                    </span>
                  </div>
                );
              })}
            </div>
          </div>

          {/* Report Tabs */}
          <div className="bg-white rounded-lg border border-gray-200 p-6 shadow-sm mb-6">
            <div className="flex flex-wrap gap-2 mb-6">
              {reportTabs.map((tab) => (
                <Button
                  key={tab.id}
                  variant={activeTab === tab.id ? "default" : "outline"}
                  onClick={() => setActiveTab(tab.id)}
                  className={activeTab === tab.id ? "bg-blue-600 hover:bg-blue-700" : ""}
                >
                  {tab.label}
                </Button>
              ))}
            </div>

            {/* Active Tab Content */}
            {activeTab === 'number-of-billed-customer' ? (
              <div className="space-y-6">
                {/* Sub Tabs for Billed Target and OD Target */}
                <div className="flex gap-2 mb-4">
                  {subTabs.map((subTab) => (
                    <Button
                      key={subTab.id}
                      variant={activeSubTab === subTab.id ? "default" : "outline"}
                      onClick={() => setActiveSubTab(subTab.id)}
                      className={activeSubTab === subTab.id ? "bg-blue-600 hover:bg-blue-700" : ""}
                      size="sm"
                    >
                      {subTab.label}
                    </Button>
                  ))}
                </div>

                {/* Billed Target Content */}
                {activeSubTab === 'billed-target' && (
                  <div className="bg-white rounded-lg border border-gray-200 p-6 shadow-sm">
                    <h2 className="text-xl font-semibold text-blue-600 mb-4">Billed Target Analysis</h2>
                    <div className="bg-blue-600 text-white p-4 rounded-lg">
                      <p className="text-center">Billed Target report will be generated here</p>
                    </div>
                  </div>
                )}

                {/* OD Target Content */}
                {activeSubTab === 'od-target' && (
                  <div className="bg-white rounded-lg border border-gray-200 p-6 shadow-sm">
                    <h2 className="text-xl font-semibold text-blue-600 mb-4">OD Target Analysis</h2>
                    <div className="space-y-4">
                      <p className="text-sm text-gray-600 mb-4">Choose OD file for OD Target calculation:</p>
                      <div className="space-y-2">
                        <div className="flex items-center space-x-2">
                          <input 
                            type="radio" 
                            id="previous-month" 
                            name="od-analysis" 
                            value="previous-month"
                            onChange={(e) => setSelectedAnalysis(e.target.value)}
                            className="text-blue-600"
                          />
                          <label htmlFor="previous-month" className="text-sm text-gray-700">OS-Previous Month</label>
                        </div>
                        <div className="flex items-center space-x-2">
                          <input 
                            type="radio" 
                            id="current-month" 
                            name="od-analysis" 
                            value="current-month"
                            onChange={(e) => setSelectedAnalysis(e.target.value)}
                            className="text-blue-600"
                          />
                          <label htmlFor="current-month" className="text-sm text-gray-700">OS-Current Month</label>
                        </div>
                      </div>
                      {!selectedAnalysis && (
                        <div className="bg-blue-600 text-white p-4 rounded-lg flex items-center space-x-2">
                          <span className="text-white">⚠</span>
                          <span>No OD file selected</span>
                        </div>
                      )}
                    </div>
                  </div>
                )}

                {/* Summary Report Generator */}
                <div className="bg-white rounded-lg border border-gray-200 p-6 shadow-sm">
                  <h2 className="text-xl font-semibold text-blue-600 mb-4">Summary Report Generator</h2>
                  <div className="bg-blue-600 text-white p-4 rounded-lg">
                    <p className="text-center">Generate at least one report from any tab to enable the consolidated report option</p>
                  </div>
                </div>
              </div>
            ) : (
              <div className="space-y-6">
                <div className="bg-white rounded-lg border border-gray-200 p-6 shadow-sm">
                  <h2 className="text-xl font-semibold text-blue-600 mb-4">
                    {reportTabs.find(tab => tab.id === activeTab)?.label} Report
                  </h2>
                  <div className="bg-blue-600 text-white p-4 rounded">
                    Please upload the required files to generate this report
                  </div>
                </div>
                {/* Summary Report Generator for other tabs */}
                <div className="bg-white rounded-lg border border-gray-200 p-6 shadow-sm">
                  <h2 className="text-xl font-semibold text-blue-600 mb-4">Summary Report Generator</h2>
                  <div className="bg-blue-600 text-white p-4 rounded-lg">
                    <p className="text-center">Generate at least one report from any tab to enable the consolidated report option</p>
                  </div>
                </div>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

export default ExecutiveView;