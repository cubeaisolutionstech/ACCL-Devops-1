
import React from 'react';
import { X, Check } from 'lucide-react';
import { useFileUpload } from '@/context/FileUploadContext';

const RequiredFiles: React.FC = () => {
  const { uploadedFiles } = useFileUpload();

  const files = [
    { name: 'Current Year Sales File', key: 'currentYearSales' },
    { name: 'Last Year Sales File', key: 'lastYearSales' },
    { name: 'Budget File', key: 'budget' },
    { name: 'OS-Previous Month File', key: 'osPrevious' },
    { name: 'OS-Current Month File', key: 'osCurrent' },
  ];

  return (
    <div className="">
      <h2 className="text-lg font-semibold text-blue-600 mb-4">Required Files:</h2>
      
      <ul className="space-y-3">
        {files.map((file, index) => {
          const isUploaded = uploadedFiles[file.key as keyof typeof uploadedFiles];
          return (
            <li key={index} className="flex items-center space-x-3">
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
            </li>
          )
        })}
      </ul>
    </div>
  );
};

export default RequiredFiles;
