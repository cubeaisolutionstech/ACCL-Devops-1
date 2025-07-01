import React, { useState } from 'react';
import { Button } from '@/components/ui/button';
import { useNavigate, useLocation } from 'react-router-dom';

const UserTabs = () => {
  const navigate = useNavigate();
  const location = useLocation();

  const getActiveTab = () => {
    if (location.pathname === '/branch') return 'branch';
    if (location.pathname === '/executive') return 'executive';
    return '';
  };

  const activeTab = getActiveTab();

  const tabs = [
    { id: 'branch', label: 'Branch', path: '/branch' },
    { id: 'executive', label: 'Executive', path: '/executive' },
  ];

  const handleTabClick = (path) => {
    navigate(path);
  };

  return (
    <div className="w-full">
      <div className="flex justify-center mb-6">
        <div className="flex bg-gray-100 rounded-lg p-1">
          {tabs.map((tab) => (
            <Button
              key={tab.id}
              variant="ghost"
              onClick={() => handleTabClick(tab.path)}
              className={`px-6 py-2 rounded-md transition-colors ${
                activeTab === tab.id
                  ? 'bg-blue-600 text-white hover:bg-blue-700'
                  : 'text-gray-600 hover:text-gray-900 hover:bg-gray-200'
              }`}
            >
              {tab.label}
            </Button>
          ))}
        </div>
      </div>
    </div>
  );
};

export default UserTabs;