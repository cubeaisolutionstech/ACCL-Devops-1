
import React, { useState } from 'react';
import { Upload, File, X, ChevronDown } from 'lucide-react';
import { Button } from '@/components/ui/button';
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from '@/components/ui/dropdown-menu';
import { useFileUpload } from '@/context/FileUploadContext';

interface FileUploadProps {
  title: string;
  fileType: string;
  maxSize?: string;
  uploadKey: string;
}

const FileUpload: React.FC<FileUploadProps> = ({ 
  title, 
  fileType, 
  maxSize = "10MB",
  uploadKey
}) => {
  const [isDragOver, setIsDragOver] = useState(false);
  const [uploadedFile, setUploadedFile] = useState<File | null>(null);
  const { setFileUploaded } = useFileUpload();

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(true);
  };

  const handleDragLeave = () => {
    setIsDragOver(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
    const files = Array.from(e.dataTransfer.files);
    if (files.length > 0) {
      setUploadedFile(files[0]);
      setFileUploaded(uploadKey, true);
    }
  };

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      setUploadedFile(files[0]);
      setFileUploaded(uploadKey, true);
    }
  };

  const removeFile = () => {
    setUploadedFile(null);
    setFileUploaded(uploadKey, false);
  };

  const exampleFiles = [
    { name: "Sales_Report_Q1.xlsx", url: "#" },
    { name: "Sales_Report_Q2.xlsx", url: "#" },
    { name: "Sales_Report_Q3.xlsx", url: "#" }
  ];

  return (
    <div className="bg-white rounded-lg border border-gray-200 p-3 shadow-sm">
      <h3 className="font-medium text-gray-900 mb-2 text-sm">{title}</h3>
      
      <div
        className={`border-2 border-dashed rounded-lg p-4 text-center transition-colors ${
          isDragOver
            ? 'border-blue-400 bg-blue-50'
            : uploadedFile
            ? 'border-green-400 bg-green-50'
            : 'border-gray-300 hover:border-gray-400'
        }`}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
      >
        {uploadedFile ? (
          <div className="flex items-center justify-center space-x-2">
            <File className="h-6 w-6 text-green-500" />
            <div className="flex-1 min-w-0">
              <p className="text-xs font-medium text-green-700 truncate">{uploadedFile.name}</p>
              <p className="text-xs text-green-600">
                {(uploadedFile.size / 1024 / 1024).toFixed(2)} MB
              </p>
            </div>
            <Button
              variant="ghost"
              size="sm"
              onClick={removeFile}
              className="text-red-500 hover:text-red-700 h-6 w-6 p-0"
            >
              <X className="h-3 w-3" />
            </Button>
          </div>
        ) : (
          <>
            <Upload className="mx-auto h-8 w-8 text-gray-400 mb-2" />
            <p className="text-xs text-gray-600 mb-1">Drag and drop file here</p>
            <p className="text-xs text-gray-500 mb-3">
              {fileType} files only • Max {maxSize}
            </p>
          </>
        )}
      </div>

      <div className={`flex space-x-2 mt-3 ${fileType === 'Image' ? 'justify-center' : ''}`}>
        <div className={fileType === 'Image' ? 'w-full' : 'flex-1'}>
          <input
            type="file"
            className="hidden"
            id={`file-input-${title}`}
            onChange={handleFileSelect}
            accept={fileType === 'Excel' ? '.xlsx,.xls' : fileType === 'Image' ? 'image/*' : '*'}
          />
          <Button 
            variant="outline" 
            size="sm" 
            className="w-full text-xs"
            onClick={() => document.getElementById(`file-input-${title}`)?.click()}
          >
            Browse
          </Button>
        </div>
        
        {fileType !== 'Image' && (
          <DropdownMenu>
            <DropdownMenuTrigger asChild>
              <Button variant="outline" size="sm" className="flex-1 text-xs">
                View
                <ChevronDown className="h-3 w-3 ml-1" />
              </Button>
            </DropdownMenuTrigger>
            <DropdownMenuContent className="w-56 bg-white">
              {exampleFiles.map((file, index) => (
                <DropdownMenuItem key={index} asChild>
                  <a href={file.url} className="cursor-pointer text-xs">
                    {file.name}
                  </a>
                </DropdownMenuItem>
              ))}
            </DropdownMenuContent>
          </DropdownMenu>
        )}
      </div>
    </div>
  );
};

export default FileUpload;
