import React from 'react';
import { SidebarProvider, SidebarTrigger } from '@/components/ui/sidebar';
import { AppSidebar } from '@/components/AppSidebar';
import RequiredFiles from '@/components/RequiredFiles';
import ReportTabs from '@/components/ReportTabs';
import UserTabs from '@/components/UserTabs';
import { FileUploadProvider } from '@/context/FileUploadContext';

const BranchPage = () => {
  return (
    <FileUploadProvider>
      <SidebarProvider>
        <div className="min-h-screen flex w-full bg-gray-50">
          <AppSidebar />
          <main className="flex-1">
            <div className="container mx-auto px-4 py-8">
              {/* Header */}
              <div className="text-center mb-8">
                <div className="flex items-center justify-between mb-4">
                  <SidebarTrigger />
                  <div className="flex-1">
                    <UserTabs />
                  </div>
                </div>
                <h1 className="text-2xl font-bold text-blue-600 mb-8 text-center">
                  Accllp Reports Dashboard
                </h1>
                {/* Main Content */}
                <div className="space-y-6 max-w-4xl mx-auto">
                  {/* Required Files Section */}
                  <div className="bg-white rounded-lg border border-gray-200 p-6 shadow-sm">
                    
                    <RequiredFiles />
                  </div>
                  {/* Report Tabs Section */}
                  <div className="bg-white rounded-lg border border-gray-200 p-6 shadow-sm">
                    <ReportTabs />
                  </div>
                </div>
              </div>
            </div>
          </main>
        </div>
      </SidebarProvider>
    </FileUploadProvider>
  );
};

export default BranchPage;