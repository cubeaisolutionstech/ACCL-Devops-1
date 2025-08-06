
import React, { useState } from 'react';
import { Button } from '@/components/ui/button';
import { useFileUpload } from '@/context/FileUploadContext';
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from '@/components/ui/table';

const ReportTabs: React.FC = () => {
  const [activeTab, setActiveTab] = useState('budget-vs-billed');
  const { uploadedFiles } = useFileUpload();

  const tabs = [
    { id: 'budget-vs-billed', label: 'Budget vs Billed' },
    { id: 'od-target-vs-collection', label: 'Od Target vs Collection' },
    { id: 'product-growth', label: 'Product Growth' },
    { id: 'number-of-billed-customer', label: 'Number of Billed customer' },
  ];

  const renderBudgetVsBilled = () => {
    const hasRequiredFiles = uploadedFiles.currentYearSales && uploadedFiles.budget;
    
    if (!hasRequiredFiles) {
      return (
        <div className="bg-blue-600 text-white p-4 rounded-lg">
          <p className="text-center">Please upload both Sales and Budget files in the sidebar</p>
        </div>
      );
    }

    return (
      <div className="space-y-4">
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead>Department</TableHead>
              <TableHead>Budget</TableHead>
              <TableHead>Billed</TableHead>
              <TableHead>Variance</TableHead>
              <TableHead>Percentage</TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
            <TableRow>
              <TableCell>Sales Team A</TableCell>
              <TableCell>₹50,000</TableCell>
              <TableCell>₹45,000</TableCell>
              <TableCell>-₹5,000</TableCell>
              <TableCell className="text-red-600">90%</TableCell>
            </TableRow>
            <TableRow>
              <TableCell>Sales Team B</TableCell>
              <TableCell>₹75,000</TableCell>
              <TableCell>₹80,000</TableCell>
              <TableCell className="text-green-600">+₹5,000</TableCell>
              <TableCell className="text-green-600">107%</TableCell>
            </TableRow>
            <TableRow>
              <TableCell>Marketing</TableCell>
              <TableCell>₹30,000</TableCell>
              <TableCell>₹28,500</TableCell>
              <TableCell>-₹1,500</TableCell>
              <TableCell className="text-red-600">95%</TableCell>
            </TableRow>
          </TableBody>
        </Table>
      </div>
    );
  };

  const renderOdTargetVsCollection = () => {
    const hasRequiredFiles = uploadedFiles.osPrevious && uploadedFiles.osCurrent;
    
    if (!hasRequiredFiles) {
      return (
        <div className="bg-gray-100 p-4 rounded-lg">
          <p className="text-center text-gray-600">Upload required OS files to generate this report</p>
        </div>
      );
    }

    return (
      <div className="space-y-4">
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead>Customer</TableHead>
              <TableHead>Target</TableHead>
              <TableHead>Collection</TableHead>
              <TableHead>Outstanding</TableHead>
              <TableHead>Status</TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
            <TableRow>
              <TableCell>ABC Corp</TableCell>
              <TableCell>₹1,00,000</TableCell>
              <TableCell>₹85,000</TableCell>
              <TableCell>₹15,000</TableCell>
              <TableCell className="text-yellow-600">Pending</TableCell>
            </TableRow>
            <TableRow>
              <TableCell>XYZ Ltd</TableCell>
              <TableCell>₹2,50,000</TableCell>
              <TableCell>₹2,50,000</TableCell>
              <TableCell>₹0</TableCell>
              <TableCell className="text-green-600">Collected</TableCell>
            </TableRow>
          </TableBody>
        </Table>
      </div>
    );
  };

  const renderProductGrowth = () => {
    const hasRequiredFiles = uploadedFiles.currentYearSales && uploadedFiles.lastYearSales;
    
    if (!hasRequiredFiles) {
      return (
        <div className="bg-gray-100 p-4 rounded-lg">
          <p className="text-center text-gray-600">Upload current and last year sales files to generate this report</p>
        </div>
      );
    }

    return (
      <div className="space-y-4">
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead>Product</TableHead>
              <TableHead>Last Year</TableHead>
              <TableHead>Current Year</TableHead>
              <TableHead>Growth</TableHead>
              <TableHead>Growth %</TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
            <TableRow>
              <TableCell>Product A</TableCell>
              <TableCell>₹1,20,000</TableCell>
              <TableCell>₹1,50,000</TableCell>
              <TableCell className="text-green-600">+₹30,000</TableCell>
              <TableCell className="text-green-600">25%</TableCell>
            </TableRow>
            <TableRow>
              <TableCell>Product B</TableCell>
              <TableCell>₹80,000</TableCell>
              <TableCell>₹75,000</TableCell>
              <TableCell className="text-red-600">-₹5,000</TableCell>
              <TableCell className="text-red-600">-6.25%</TableCell>
            </TableRow>
          </TableBody>
        </Table>
      </div>
    );
  };

  const renderBilledCustomers = () => {
    const hasRequiredFiles = uploadedFiles.currentYearSales;
    
    if (!hasRequiredFiles) {
      return (
        <div className="bg-gray-100 p-4 rounded-lg">
          <p className="text-center text-gray-600">Upload current year sales file to generate this report</p>
        </div>
      );
    }

    return (
      <div className="space-y-4">
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead>Month</TableHead>
              <TableHead>New Customers</TableHead>
              <TableHead>Repeat Customers</TableHead>
              <TableHead>Total Billed</TableHead>
              <TableHead>Revenue</TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
            <TableRow>
              <TableCell>January</TableCell>
              <TableCell>15</TableCell>
              <TableCell>45</TableCell>
              <TableCell>60</TableCell>
              <TableCell>₹3,50,000</TableCell>
            </TableRow>
            <TableRow>
              <TableCell>February</TableCell>
              <TableCell>12</TableCell>
              <TableCell>48</TableCell>
              <TableCell>60</TableCell>
              <TableCell>₹3,75,000</TableCell>
            </TableRow>
          </TableBody>
        </Table>
      </div>
    );
  };

  return (
    <div>
      <div className="flex flex-wrap gap-2 mb-6">
        {tabs.map((tab) => (
          <Button
            key={tab.id}
            variant={activeTab === tab.id ? 'default' : 'outline'}
            onClick={() => setActiveTab(tab.id)}
            className={`${
              activeTab === tab.id
                ? 'bg-blue-600 hover:bg-blue-700 text-white'
                : 'hover:bg-blue-50 text-blue-600 border-blue-200'
            }`}
          >
            {tab.label}
          </Button>
        ))}
      </div>

      <div className="space-y-6">
        {activeTab === 'budget-vs-billed' && (
          <div className="bg-white rounded-lg border border-gray-200 p-6 shadow-sm">
            <h3 className="text-lg font-semibold text-blue-600 mb-4">Budget vs Billed Report</h3>
            {renderBudgetVsBilled()}
          </div>
        )}

        {activeTab === 'od-target-vs-collection' && (
          <div className="bg-white rounded-lg border border-gray-200 p-6 shadow-sm">
            <h3 className="text-lg font-semibold text-blue-600 mb-4">Od Target vs Collection Report</h3>
            {renderOdTargetVsCollection()}
          </div>
        )}

        {activeTab === 'product-growth' && (
          <div className="bg-white rounded-lg border border-gray-200 p-6 shadow-sm">
            <h3 className="text-lg font-semibold text-blue-600 mb-4">Product Growth Report</h3>
            {renderProductGrowth()}
          </div>
        )}

        {activeTab === 'number-of-billed-customer' && (
          <div className="bg-white rounded-lg border border-gray-200 p-6 shadow-sm">
            <h3 className="text-lg font-semibold text-blue-600 mb-4">Number of Billed Customer Report</h3>
            {renderBilledCustomers()}
          </div>
        )}

        <div className="bg-white rounded-lg border border-gray-200 p-6 shadow-sm">
          <h3 className="text-lg font-semibold text-blue-600 mb-4">Consolidated Report</h3>
          <div className="bg-blue-600 text-white p-4 rounded-lg">
            <p className="text-center">Generate at least one report from any tab to enable the consolidated report option</p>
          </div>
        </div>
      </div>
    </div>
  );
};

export default ReportTabs;
