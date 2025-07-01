
import React, { useRef } from "react";
import { Upload } from "lucide-react";
import { Button } from "@/components/ui/button";

interface UploadCardProps {
  onFileSelect?: (file: File) => void;
  accept?: string;
  maxSizeMB?: number;
  description?: string;
  supportedFormats?: string;
  children?: React.ReactNode;
}

const UploadCard: React.FC<UploadCardProps> = ({
  onFileSelect,
  accept = ".xlsx,.xls,.csv",
  maxSizeMB = 50,
  description = "Drag and drop your company & product mapping file here",
  supportedFormats = "Supported formats: CSV, Excel (.csv, .xls, .xlsx) - Max 50MB",
  children,
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleBrowseClick = () => {
    fileInputRef.current?.click();
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file && onFileSelect) {
      onFileSelect(file);
    }
  };

  const handleDrop = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    const file = event.dataTransfer.files?.[0];
    if (file && onFileSelect) {
      onFileSelect(file);
    }
  };

  const handleDragOver = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
  };

  return (
    <div
      className="w-full max-w-2xl bg-blue-50 border-2 border-dashed border-blue-400 rounded-2xl flex flex-col items-center justify-center px-6 py-16 text-center shadow transition hover:shadow-md"
      onDrop={handleDrop}
      onDragOver={handleDragOver}
      tabIndex={0}
      role="button"
      aria-label="File Upload"
    >
      <Upload className="w-12 h-12 text-blue-500 mx-auto mb-4" />
      <div className="font-semibold text-blue-700 text-lg mb-2">
        {description}
      </div>
      {children}
      <Button
        type="button"
        className="bg-white border border-blue-400 text-blue-600 hover:bg-blue-50 font-semibold px-8 py-3 rounded text-base my-3 transition shadow-none"
        onClick={handleBrowseClick}
        variant={undefined}
      >
        Browse Files
      </Button>
      <input
        ref={fileInputRef}
        type="file"
        className="hidden"
        accept={accept}
        onChange={handleFileChange}
      />
      <div className="text-sm text-blue-600 mt-2">{supportedFormats}</div>
    </div>
  );
};

export default UploadCard;

