
import { useRef } from "react";
import UploadCard from "@/components/UploadCard";
import { Button } from "@/components/ui/button";

// Sample/placeholder counts. Replace with dynamic values if available.
const branchesCount = 1;
const regionsCount = 1;
const companiesCount = 1;

const BackupRestore = () => {
  const handleFileSelect = (file: File) => {
    // Handle file logic here
    console.log("Selected file: ", file);
  };

  return (
    <div className="mt-8 flex flex-col lg:flex-row gap-12 min-h-[500px]">
      {/* Backup Mappings Section */}
      <div className="flex-1 pr-0 lg:pr-8 border-r border-gray-200">
        <h2 className="text-3xl font-bold text-blue-800 mb-6">Backup Mappings</h2>
        <div className="text-lg text-blue-800 font-semibold mb-4">
          Export branch, region, and company mappings:
        </div>
        <ul className="ml-6 text-blue-900 text-base font-medium space-y-2 mb-8">
          <li>Branches: {branchesCount}</li>
          <li>Regions: {regionsCount}</li>
          <li>Companies: {companiesCount}</li>
        </ul>
        <Button className="bg-blue-700 hover:bg-blue-800 text-white px-6 py-2 rounded font-semibold">
          Create Backup
        </Button>
      </div>
      {/* Restore Mappings Section */}
      <div className="flex-1 flex flex-col items-center justify-start pl-0 lg:pl-8 mt-12 lg:mt-0">
        <h2 className="text-3xl font-bold text-blue-800 mb-8 text-left w-full">Upload Company & Product Mapping File</h2>
        <UploadCard
          onFileSelect={handleFileSelect}
          description="Drag and drop your company & product mapping file here"
          supportedFormats="Supported formats: CSV, Excel (.csv, .xls, .xlsx) - Max 50MB"
        />
      </div>
    </div>
  );
};

export default BackupRestore;

