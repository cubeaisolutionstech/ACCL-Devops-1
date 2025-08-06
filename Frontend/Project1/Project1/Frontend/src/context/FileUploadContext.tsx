
import React, { createContext, useContext, useState, ReactNode } from 'react';

interface FileUploadContextType {
  uploadedFiles: {
    currentYearSales: boolean;
    lastYearSales: boolean;
    budget: boolean;
    osPrevious: boolean;
    osCurrent: boolean;
    logo: boolean;
  };
  setFileUploaded: (fileType: string, uploaded: boolean) => void;
}

const FileUploadContext = createContext<FileUploadContextType | undefined>(undefined);

export const FileUploadProvider: React.FC<{ children: ReactNode }> = ({ children }) => {
  const [uploadedFiles, setUploadedFiles] = useState({
    currentYearSales: false,
    lastYearSales: false,
    budget: false,
    osPrevious: false,
    osCurrent: false,
    logo: false,
  });

  const setFileUploaded = (fileType: string, uploaded: boolean) => {
    setUploadedFiles(prev => ({
      ...prev,
      [fileType]: uploaded
    }));
  };

  return (
    <FileUploadContext.Provider value={{ uploadedFiles, setFileUploaded }}>
      {children}
    </FileUploadContext.Provider>
  );
};

export const useFileUpload = () => {
  const context = useContext(FileUploadContext);
  if (context === undefined) {
    throw new Error('useFileUpload must be used within a FileUploadProvider');
  }
  return context;
};
