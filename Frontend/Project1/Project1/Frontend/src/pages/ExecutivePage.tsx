
import React from 'react';
import { SidebarProvider, SidebarTrigger } from '@/components/ui/sidebar';
import { AppSidebar } from '@/components/AppSidebar';
import ExecutiveView from '@/components/ExecutiveView';
import UserTabs from '@/components/UserTabs';
import { FileUploadProvider } from '@/context/FileUploadContext';

const ExecutivePage = () => {
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
                <ExecutiveView />
              </div>
            </div>
          </main>
        </div>
      </SidebarProvider>
    </FileUploadProvider>
  );
};

export default ExecutivePage;
