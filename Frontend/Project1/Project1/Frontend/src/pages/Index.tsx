import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Badge } from "@/components/ui/badge";
import { Textarea } from "@/components/ui/textarea";
import { ChevronRight, Search, Filter, ChevronUp, ChevronDown, Upload, CloudUpload, Download } from "lucide-react";
import { useToast } from "@/hooks/use-toast";
import BackupRestore from "@/components/BackupRestore";
import { toast } from "sonner";
import { FileText } from "lucide-react";
import { cn } from "@/lib/utils";
import { useNavigate } from "react-router-dom";

type SortKey = 'name' | 'code' | 'customers' | 'branches';
type SortOrder = 'asc' | 'desc';

// Sample data
const sampleData = [
  {
    customerCode: "CUST001",
    customerName: "Acme Corporation",
    executive: "John Smith",
    executiveCode: "EXE001",
    branch: "New York",
    region: "Northeast"
  },
  {
    customerCode: "CUST002", 
    customerName: "Global Tech Solutions",
    executive: "Sarah Johnson",
    executiveCode: "EXE002",
    branch: "California",
    region: "West Coast"
  },
  {
    customerCode: "CUST003",
    customerName: "Metro Industries",
    executive: "Michael Davis",
    executiveCode: "EXE003", 
    branch: "Texas",
    region: "South"
  }
];

const executives = ["All Executives", ...Array.from(new Set(sampleData.map(d => d.executive)))];
const branches = ["All Branches", ...Array.from(new Set(sampleData.map(d => d.branch)))];
const regions = ["All Regions", ...Array.from(new Set(sampleData.map(d => d.region)))];

const Index = () => {
  const [activeSection, setActiveSection] = useState("executive-management");
  const [activeTab, setActiveTab] = useState("executive-creation");
  const [sortKey, setSortKey] = useState<SortKey>('name');
  const [sortOrder, setSortOrder] = useState<SortOrder>('asc');
  const [selectedExecutive, setSelectedExecutive] = useState("All Executives");
  const [removeCustomerCodes, setRemoveCustomerCodes] = useState("");
  const [assignCustomerCodes, setAssignCustomerCodes] = useState("");
  const [newCustomerCodes, setNewCustomerCodes] = useState("");
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const { toast } = useToast();

  // Branch & Region Mapping states
  const [branchRegionTab, setBranchRegionTab] = useState("manual-entry");
  const [newBranch, setNewBranch] = useState({ name: "" });
  const [newRegion, setNewRegion] = useState({ name: "" });
  const [selectedBranch, setSelectedBranch] = useState("All Branches");
  const [selectedRegion, setSelectedRegion] = useState("");
  const [selectedExecutives, setSelectedExecutives] = useState("");
  const [selectedBranches, setSelectedBranches] = useState("");
  const [removeBranches, setRemoveBranches] = useState("");
  const [removeRegions, setRemoveRegions] = useState("");

  // Company & Product Mapping states
  const [companyProductTab, setCompanyProductTab] = useState("manual-entry");
  const [newProductGroup, setNewProductGroup] = useState({ name: "" });
  const [newCompanyGroup, setNewCompanyGroup] = useState({ name: "" });
  const [selectedCompany, setSelectedCompany] = useState("");
  const [selectedExecutivesForCompany, setSelectedExecutivesForCompany] = useState("");
  const [removeProducts, setRemoveProducts] = useState("");
  const [removeCompanies, setRemoveCompanies] = useState("");
  
  // Sample data for executives
  const [executivesData, setExecutives] = useState([
    { id: 1, name: "John Smith", code: "EXE001", customers: 145, branches: 8 },
    { id: 2, name: "Sarah Johnson", code: "EXE002", customers: 230, branches: 12 },
    { id: 3, name: "Michael Brown", code: "EXE003", customers: 187, branches: 6 },
  ]);

  // Sample data for branches and regions
  const [branchesData, setBranches] = useState([
    { id: 1, name: "Mumbai Central", executives: 3, regions: 2 },
    { id: 2, name: "Delhi North", executives: 5, regions: 1 },
    { id: 3, name: "Chennai South", executives: 2, regions: 3 },
  ]);

  const [regionsData, setRegions] = useState([
    { id: 1, name: "Western Region", branches: 8 },
    { id: 2, name: "Northern Region", branches: 12 },
    { id: 3, name: "Southern Region", branches: 6 },
  ]);

  // Sample data for products and companies
  const [productsData, setProducts] = useState([
    { id: 1, name: "Software Solutions", companies: 15 },
    { id: 2, name: "Hardware Equipment", companies: 8 },
    { id: 3, name: "Consulting Services", companies: 12 },
  ]);

  const [companiesData, setCompanies] = useState([
    { id: 1, name: "Tech Corp", products: 3 },
    { id: 2, name: "Innovation Ltd", products: 2 },
    { id: 3, name: "Digital Solutions", products: 4 },
  ]);

  // Sample mapping data
  const [mappingData, setMappingData] = useState([
    { branch: "Mumbai Central", region: "Western Region", executives: ["John Smith", "Sarah Johnson"], count: 2 },
    { branch: "Delhi North", region: "Northern Region", executives: ["Michael Brown"], count: 1 },
    { branch: "Chennai South", region: "Southern Region", executives: ["John Smith"], count: 1 },
  ]);

  // Sample company mapping data
  const [companyMappingData] = useState([
    { company: "Tech Corp", products: ["Software Solutions", "Hardware Equipment"], count: 2 },
    { company: "Innovation Ltd", products: ["Consulting Services"], count: 1 },
    { company: "Digital Solutions", products: ["Software Solutions", "Consulting Services"], count: 2 },
  ]);

  const [newExecutive, setNewExecutive] = useState({
    name: "",
    code: ""
  });

  // Sample assigned customers data
  const [assignedCustomers] = useState([
    { code: "CUST001", name: "ABC Corporation" },
    { code: "CUST002", name: "XYZ Industries" },
  ]);

  const [tableData, setTableData] = useState([
    { name: "John", age: 25 },
    { name: "Jane", age: 30 }
  ]);

  const [editRow, setEditRow] = useState({ name: "", age: "" });
  const [editIndex, setEditIndex] = useState<number | null>(null);

  const handleEdit = (row: typeof tableData[number], idx: number) => {
    setEditRow({ name: row.name, age: String(row.age) });
    setEditIndex(idx);
  };

  const handleUpdate = () => {
    const updated = [...tableData];
    updated[editIndex!] = { ...editRow, age: Number(editRow.age) };
    setTableData(updated);
    setEditRow({ name: "", age: "" });
    setEditIndex(null);
  };

  const handleSort = (key: SortKey) => {
    if (sortKey === key) {
      setSortOrder(sortOrder === 'asc' ? 'desc' : 'asc');
    } else {
      setSortKey(key);
      setSortOrder('asc');
    }
  };

  const sortedExecutives = [...executivesData].sort((a, b) => {
    let aValue = a[sortKey];
    let bValue = b[sortKey];

    if (typeof aValue === 'string') {
      aValue = aValue.toLowerCase();
      bValue = (bValue as string).toLowerCase();
    }

    if (sortOrder === 'asc') {
      return aValue < bValue ? -1 : aValue > bValue ? 1 : 0;
    } else {
      return aValue > bValue ? -1 : aValue < bValue ? 1 : 0;
    }
  });

  const SortIcon = ({ column }: { column: SortKey }) => {
    if (sortKey !== column) {
      return <div className="w-4 h-4" />; // Empty space to maintain alignment
    }
    return sortOrder === 'asc' ? 
      <ChevronUp className="w-4 h-4" /> : 
      <ChevronDown className="w-4 h-4" />;
  };

  const handleAddExecutive = () => {
    if (newExecutive.name && newExecutive.code) {
      setExecutives([...executivesData, {
        id: executivesData.length + 1,
        name: newExecutive.name,
        code: newExecutive.code,
        customers: 0,
        branches: 0
      }]);
      setNewExecutive({ name: "", code: "" });
    }
  };

  const handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      // Validate file type (accept CSV, Excel files)
      const allowedTypes = [
        'text/csv',
        'application/vnd.ms-excel',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      ];
      
      if (!allowedTypes.includes(file.type)) {
        toast({
          title: "Invalid file type",
          description: "Please select a CSV or Excel file (.csv, .xls, .xlsx)",
          variant: "destructive"
        });
        return;
      }

      // Check file size (limit to 5MB)
      if (file.size > 5 * 1024 * 1024) {
        toast({
          title: "File too large",
          description: "Please select a file smaller than 5MB",
          variant: "destructive"
        });
        return;
      }

      setSelectedFile(file);
      toast({
        title: "File selected",
        description: `Selected: ${file.name}`,
      });
    }
  };

  const handleFileDrop = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    const file = event.dataTransfer.files[0];
    if (file) {
      // Create a synthetic event to reuse the file validation logic
      const fileList = [file];
      const syntheticEvent = {
        target: { files: fileList }
      } as unknown as React.ChangeEvent<HTMLInputElement>;
      handleFileSelect(syntheticEvent);
    }
  };

  const handleDragOver = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
  };

  const handleFileUpload = () => {
    if (!selectedFile) {
      toast({
        title: "No file selected",
        description: "Please select a file first",
        variant: "destructive"
      });
      return;
    }

    // Here you would typically upload the file to your backend
    // For now, we'll just show a success message
    toast({
      title: "File uploaded successfully",
      description: `${selectedFile.name} has been processed`,
    });
    
    // Reset the selected file
    setSelectedFile(null);
  };

  // Branch & Region functions
  const handleAddBranch = () => {
    if (newBranch.name) {
      setBranches([...branchesData, {
        id: branchesData.length + 1,
        name: newBranch.name,
        executives: 0,
        regions: 0
      }]);
      setNewBranch({ name: "" });
      toast({
        title: "Branch added successfully",
        description: `${newBranch.name} has been created`,
      });
    }
  };

  const handleAddRegion = () => {
    if (newRegion.name) {
      setRegions([...regionsData, {
        id: regionsData.length + 1,
        name: newRegion.name,
        branches: 0
      }]);
      setNewRegion({ name: "" });
      toast({
        title: "Region added successfully",
        description: `${newRegion.name} has been created`,
      });
    }
  };

  // Company & Product functions
  const handleAddProductGroup = () => {
    if (newProductGroup.name) {
      setProducts([...productsData, {
        id: productsData.length + 1,
        name: newProductGroup.name,
        companies: 0
      }]);
      setNewProductGroup({ name: "" });
      toast({
        title: "Product group added successfully",
        description: `${newProductGroup.name} has been created`,
      });
    }
  };

  const handleAddCompanyGroup = () => {
    if (newCompanyGroup.name) {
      setCompanies([...companiesData, {
        id: companiesData.length + 1,
        name: newCompanyGroup.name,
        products: 0
      }]);
      setNewCompanyGroup({ name: "" });
      toast({
        title: "Company group added successfully",
        description: `${newCompanyGroup.name} has been created`,
      });
    }
  };

  const handleRemoveCustomers = () => {
    if (removeCustomerCodes) {
      toast({
        title: "Customers removed",
        description: `Customer ${removeCustomerCodes} has been removed from the executive`,
      });
      setRemoveCustomerCodes("");
    }
  };

  const handleRemoveBranches = () => {
    if (removeBranches) {
      setBranches(branchesData.filter(branch => branch.name !== removeBranches));
      toast({
        title: "Branch removed",
        description: `${removeBranches} has been removed`,
      });
      setRemoveBranches("");
    }
  };

  const handleRemoveRegions = () => {
    if (removeRegions) {
      setRegions(regionsData.filter(region => region.name !== removeRegions));
      toast({
        title: "Region removed", 
        description: `${removeRegions} has been removed`,
      });
      setRemoveRegions("");
    }
  };

  const handleRemoveProducts = () => {
    if (removeProducts) {
      setProducts(productsData.filter(product => product.name !== removeProducts));
      toast({
        title: "Product removed",
        description: `${removeProducts} has been removed`,
      });
      setRemoveProducts("");
    }
  };

  const handleRemoveCompanies = () => {
    if (removeCompanies) {
      setCompanies(companiesData.filter(company => company.name !== removeCompanies));
      toast({
        title: "Company removed",
        description: `${removeCompanies} has been removed`,
      });
      setRemoveCompanies("");
    }
  };

  const roleButtons = [
    { id: "admin", label: "Admin" },
    { id: "branch", label: "Branch" },
    { id: "executive", label: "Executive" }
  ];

  const navigationSections = [
    { id: "executive-management", label: "Executive Management" },
    { id: "branch-region", label: "Branch & Region Mapping" },
    { id: "company-product", label: "Company & Product Mapping" },
    { id: "consolidated-data", label: "Consolidated Data View" },
    { id: "file-processing", label: "File Processing" }
  ];

  const [selectedExecutiveFilter, setSelectedExecutiveFilter] = useState("All Executives");
  const [selectedBranchFilter, setSelectedBranchFilter] = useState("All Branches");
  const [selectedRegionFilter, setSelectedRegionFilter] = useState("All Regions");
  
  const filteredData = sampleData.filter(item => {
    const executiveMatch = selectedExecutiveFilter === "All Executives" || item.executive === selectedExecutiveFilter;
    const branchMatch = selectedBranchFilter === "All Branches" || item.branch === selectedBranchFilter;
    const regionMatch = selectedRegionFilter === "All Regions" || item.region === selectedRegionFilter;
    return executiveMatch && branchMatch && regionMatch;
  });

  const handleDownloadCSV = () => {
    const headers = ["Customer Code", "Customer Name", "Executive", "Executive Code", "Branch", "Region"];
    const csvContent = [
      headers.join(","),
      ...filteredData.map(row => 
        [row.customerCode, row.customerName, row.executive, row.executiveCode, row.branch, row.region]
          .map(field => `"${field}"`)
          .join(",")
      )
    ].join("\n");

    const blob = new Blob([csvContent], { type: "text/csv" });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "consolidated_data.csv";
    a.click();
    window.URL.revokeObjectURL(url);
    
    toast({
      title: "Success",
      description: "CSV file downloaded successfully!",
    });
  };

  // For "Map Executives to Branches"
  const handleUpdateBranchMapping = () => {
    if (!selectedBranch || !selectedExecutives) {
      toast({
        title: "Select both branch and executive",
        variant: "destructive"
      });
      return;
    }
    // Find region for the selected branch
    const branchObj = branchesData.find(b => b.name === selectedBranch);
    const region = branchObj
      ? (mappingData.find(m => m.branch === selectedBranch)?.region || "")
      : "";

    // Find executive name from code
    const executiveObj = executivesData.find(e => e.code === selectedExecutives);
    const executiveName = executiveObj ? executiveObj.name : selectedExecutives;

    // Check if mapping already exists
    const mappingIndex = mappingData.findIndex(m => m.branch === selectedBranch);

    let updatedMappingData = [...mappingData];
    if (mappingIndex !== -1) {
      // Update existing mapping
      const execs = updatedMappingData[mappingIndex].executives;
      if (!execs.includes(executiveName)) {
        execs.push(executiveName);
        updatedMappingData[mappingIndex].count = execs.length;
      }
    } else {
      // Add new mapping
      updatedMappingData.push({
        branch: selectedBranch,
        region: region || "",
        executives: [executiveName],
        count: 1
      });
    }
    setMappingData(updatedMappingData);
    toast({
      title: "Branch mapping updated",
      description: `Executive mapped to branch successfully!`
    });
  };

  // For "Map Executives to Regions"
  const handleUpdateRegionMapping = () => {
    if (!selectedRegion || !selectedBranches) {
      toast({
        title: "Select both region and branch",
        variant: "destructive"
      });
      return;
    }
    // Find mapping for the branch
    const mappingIndex = mappingData.findIndex(m => m.branch === selectedBranches);

    let updatedMappingData = [...mappingData];
    if (mappingIndex !== -1) {
      updatedMappingData[mappingIndex].region = selectedRegion;
    } else {
      updatedMappingData.push({
        branch: selectedBranches,
        region: selectedRegion,
        executives: [],
        count: 0
      });
    }
    setMappingData(updatedMappingData);
    toast({
      title: "Region mapping updated",
      description: `Branch mapped to region successfully!`
    });
  };

  const handleLogout = () => {
    localStorage.removeItem('auth_token');
    localStorage.removeItem('user_data');
     toast({
      title: "Logged out",
      description: "You have been logged out successfully."
    });
    window.location.reload();
  };

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Header with breadcrumb */}
      <div className="bg-white border-b">
        <div className="container mx-auto px-6 py-4">
          <div className="flex items-center text-gray-600 mb-4">
            <ChevronRight className="w-4 h-4" />
          </div>
          
          {/* Role Selection */}
          <div className="flex justify-center gap-2 mb-6">
            <Button
              key="admin"
              variant={"default"}
              className={"px-6 py-2 rounded-full bg-blue-600 text-white"}
            >
              Admin
            </Button>
            <button
              onClick={handleLogout}
              className="ml-4 px-3 py-1 bg-red-500 text-white rounded hover:bg-red-600 text-sm font-medium"
            >
              Logout
            </button>
          </div>

          {/* Main Title */}
          <h1 className="text-3xl font-bold text-blue-600 text-center">
            Executive Mapping Administration Portal
          </h1>
        </div>
      </div>

      {/* Navigation Sections */}
      <div className="container mx-auto px-6 py-6">
        <div className="flex flex-wrap justify-center gap-2 mb-8">
          {navigationSections.map((section) => (
            <Button
              key={section.id}
              onClick={() => setActiveSection(section.id)}
              variant={activeSection === section.id ? "default" : "outline"}
              className={`px-4 py-2 rounded-full text-sm ${
                activeSection === section.id 
                  ? "bg-blue-600 text-white" 
                  : "bg-white text-blue-600 border-blue-200 hover:bg-blue-50"
              }`}
            >
              {section.label}
            </Button>
          ))}
        </div>

        {/* Main Content */}
        {activeSection === "executive-management" && (
          <div className="space-y-6">
            {/* Sub Navigation */}
            <Tabs value={activeTab} onValueChange={setActiveTab} className="w-full">
              <div className="flex justify-center">
                <TabsList className="grid w-fit grid-cols-2 bg-white border">
                  <TabsTrigger 
                    value="executive-creation"
                    className="px-6 py-2 data-[state=active]:bg-blue-600 data-[state=active]:text-white"
                  >
                    Executive Creation
                  </TabsTrigger>
                  <TabsTrigger 
                    value="customer-code"
                    className="px-6 py-2 data-[state=active]:bg-blue-600 data-[state=active]:text-white"
                  >
                    Customer Code Management
                  </TabsTrigger>
                </TabsList>
              </div>

              <TabsContent value="executive-creation" className="mt-8">
                <div className="grid lg:grid-cols-2 gap-8">
                  {/* Add New Executive Form */}
                  <Card className="h-fit">
                    <CardHeader>
                      <CardTitle className="text-xl font-semibold text-blue-600 text-center">
                        Add New Executive
                      </CardTitle>
                    </CardHeader>
                    <CardContent className="space-y-6">
                      <div className="space-y-2">
                        <Label htmlFor="executive-name" className="text-blue-600 font-medium">
                          Executive Name :
                        </Label>
                        <Input
                          id="executive-name"
                          value={newExecutive.name}
                          onChange={(e) => setNewExecutive({...newExecutive, name: e.target.value})}
                          className="bg-gray-100 border-gray-300"
                          placeholder="Enter executive name"
                        />
                      </div>
                      
                      <div className="space-y-2">
                        <Label htmlFor="executive-code" className="text-blue-600 font-medium">
                          Executive Code:
                        </Label>
                        <Input
                          id="executive-code"
                          value={newExecutive.code}
                          onChange={(e) => setNewExecutive({...newExecutive, code: e.target.value})}
                          className="bg-gray-100 border-gray-300"
                          placeholder="Enter executive code"
                        />
                      </div>
                      
                      <div className="flex justify-center">
                        <Button 
                          onClick={handleAddExecutive}
                          className="bg-blue-600 hover:bg-blue-700 text-white px-8 py-2"
                        >
                          Add Executive
                        </Button>
                      </div>
                    </CardContent>
                  </Card>

                  {/* Current Executives Table */}
                  <Card>
                    <CardHeader>
                      <CardTitle className="text-xl font-semibold text-blue-600 text-center">
                        Current Executives
                      </CardTitle>
                    </CardHeader>
                    <CardContent>
                      <div className="overflow-x-auto">
                        <Table>
                          <TableHeader>
                            <TableRow className="bg-gray-50">
                              <TableHead 
                                className="font-semibold text-gray-700 text-center cursor-pointer hover:bg-gray-100 transition-colors"
                                onClick={() => handleSort('name')}
                              >
                                <div className="flex items-center justify-center gap-1">
                                  Executive
                                  <SortIcon column="name" />
                                </div>
                              </TableHead>
                              <TableHead 
                                className="font-semibold text-gray-700 text-center cursor-pointer hover:bg-gray-100 transition-colors"
                                onClick={() => handleSort('code')}
                              >
                                <div className="flex items-center justify-center gap-1">
                                  Code
                                  <SortIcon column="code" />
                                </div>
                              </TableHead>
                              <TableHead 
                                className="font-semibold text-gray-700 text-center cursor-pointer hover:bg-gray-100 transition-colors"
                                onClick={() => handleSort('customers')}
                              >
                                <div className="flex items-center justify-center gap-1">
                                  Customers
                                  <SortIcon column="customers" />
                                </div>
                              </TableHead>
                              <TableHead 
                                className="font-semibold text-gray-700 text-center cursor-pointer hover:bg-gray-100 transition-colors"
                                onClick={() => handleSort('branches')}
                              >
                                <div className="flex items-center justify-center gap-1">
                                  Branches
                                  <SortIcon column="branches" />
                                </div>
                              </TableHead>
                            </TableRow>
                          </TableHeader>
                          <TableBody>
                            {sortedExecutives.map((executive) => (
                              <TableRow key={executive.id} className="hover:bg-gray-50">
                                <TableCell className="font-medium text-center">{executive.name}</TableCell>
                                <TableCell className="text-center">
                                  <div className="flex justify-center">
                                    <Badge variant="secondary" className="bg-blue-100 text-blue-800">
                                      {executive.code}
                                    </Badge>
                                  </div>
                                </TableCell>
                                <TableCell className="text-center">{executive.customers}</TableCell>
                                <TableCell className="text-center">{executive.branches}</TableCell>
                              </TableRow>
                            ))}
                          </TableBody>
                        </Table>
                      </div>
                      
                      {/* Executive Code Filter */}
                      <div className="mt-6 p-4 bg-gray-50 rounded-lg">
                        <Label className="text-blue-600 font-medium mb-2 block text-center">
                          Executive Code:
                        </Label>
                        <Select>
                          <SelectTrigger className="bg-white">
                            <SelectValue placeholder="Select executive code" />
                          </SelectTrigger>
                          <SelectContent>
                            {executivesData.map((executive) => (
                              <SelectItem key={executive.id} value={executive.code}>
                                {executive.code} - {executive.name}
                              </SelectItem>
                            ))}
                          </SelectContent>
                        </Select>
                      </div>
                    </CardContent>
                  </Card>
                </div>
              </TabsContent>

              <TabsContent value="customer-code" className="mt-8">
                <div className="space-y-8">
                  {/* Bulk Customer Assignment */}
                  <Card>
                    <CardHeader>
                      <CardTitle className="text-xl font-semibold text-blue-600">
                        Bulk Customer Assignment
                      </CardTitle>
                    </CardHeader>
                    <CardContent>
                      <div 
                        className="border-2 border-dashed border-blue-300 rounded-lg p-8 bg-blue-50 text-center relative"
                        onDrop={handleFileDrop}
                        onDragOver={handleDragOver}
                      >
                        <CloudUpload className="w-8 h-8 text-blue-600 mx-auto mb-2" />
                        <p className="text-blue-700 mb-4">
                          {selectedFile ? `Selected: ${selectedFile.name}` : "Drag and drop your file here"}
                        </p>
                        <div className="space-y-2">
                          <input
                            type="file"
                            id="file-upload"
                            className="hidden"
                            accept=".csv,.xls,.xlsx"
                            onChange={handleFileSelect}
                          />
                          <label htmlFor="file-upload">
                            <Button 
                              variant="outline" 
                              className="bg-white text-blue-600 border-blue-300 hover:bg-blue-50"
                              asChild
                            >
                              <span className="cursor-pointer">Browse Files</span>
                            </Button>
                          </label>
                          {selectedFile && (
                            <div className="mt-4">
                              <Button 
                                onClick={handleFileUpload}
                                className="bg-blue-600 hover:bg-blue-700 text-white"
                              >
                                Upload File
                              </Button>
                            </div>
                          )}
                        </div>
                        <p className="text-xs text-gray-500 mt-2">
                          Supported formats: CSV, Excel (.csv, .xls, .xlsx) - Max 5MB
                        </p>
                      </div>
                    </CardContent>
                  </Card>

                  {/* Manual Customer Management */}
                  <Card>
                    <CardHeader>
                      <CardTitle className="text-xl font-semibold text-blue-600">
                        Manual Customer Management
                      </CardTitle>
                    </CardHeader>
                    <CardContent>
                      <div className="grid lg:grid-cols-2 gap-8">
                        {/* Left Column - Select Executive & Assigned Customers */}
                        <div className="space-y-6">
                          <div className="space-y-2">
                            <Label className="text-blue-600 font-medium">
                              Select Executive:
                            </Label>
                            <Select value={selectedExecutive} onValueChange={setSelectedExecutive}>
                              <SelectTrigger className="bg-gray-100">
                                <SelectValue placeholder="Select an executive" />
                              </SelectTrigger>
                              <SelectContent>
                                {executivesData.map((executive) => (
                                  <SelectItem key={executive.id} value={executive.code}>
                                    {executive.code} - {executive.name}
                                  </SelectItem>
                                ))}
                              </SelectContent>
                            </Select>
                          </div>

                          <div>
                            <h3 className="text-blue-600 font-semibold mb-4">
                              Assigned Customers ({assignedCustomers.length})
                            </h3>
                            <div className="border rounded-lg">
                              <Table>
                                <TableHeader>
                                  <TableRow className="bg-gray-100">
                                    <TableHead className="text-center font-semibold">Code</TableHead>
                                    <TableHead className="text-center font-semibold">Name</TableHead>
                                  </TableRow>
                                </TableHeader>
                                <TableBody>
                                  {assignedCustomers.map((customer, index) => (
                                    <TableRow key={index}>
                                      <TableCell className="text-center">{customer.code}</TableCell>
                                      <TableCell className="text-center">{customer.name}</TableCell>
                                    </TableRow>
                                  ))}
                                </TableBody>
                              </Table>
                            </div>
                          </div>
                        </div>

                        {/* Right Column - Actions */}
                        <div className="space-y-6">
                          <div>
                            <h3 className="text-blue-600 font-semibold mb-4">Action</h3>
                            
                            {/* Remove Customers */}
                            <div className="space-y-4 mb-6">
                              <Label className="text-blue-600 font-medium">
                                Remove Customers:
                              </Label>
                              <Select value={removeCustomerCodes} onValueChange={setRemoveCustomerCodes}>
                                <SelectTrigger className="bg-gray-100">
                                  <SelectValue placeholder="Select customers to remove" />
                                </SelectTrigger>
                                <SelectContent>
                                  {assignedCustomers.map((customer, index) => (
                                    <SelectItem key={index} value={customer.code}>
                                      {customer.code} - {customer.name}
                                    </SelectItem>
                                  ))}
                                </SelectContent>
                              </Select>
                              <Button 
                                onClick={handleRemoveCustomers}
                                className="bg-red-600 hover:bg-red-700 text-white w-full"
                              >
                                Remove Customers
                              </Button>
                            </div>

                            {/* Assign Customers */}
                            <div className="space-y-4 mb-6">
                              <Label className="text-blue-600 font-medium">
                                Assign Customers:
                              </Label>
                              <Select value={assignCustomerCodes} onValueChange={setAssignCustomerCodes}>
                                <SelectTrigger className="bg-gray-100">
                                  <SelectValue placeholder="Select customers to assign" />
                                </SelectTrigger>
                                <SelectContent>
                                  <SelectItem value="CUST003">CUST003 - New Company Ltd</SelectItem>
                                  <SelectItem value="CUST004">CUST004 - Tech Solutions Inc</SelectItem>
                                </SelectContent>
                              </Select>
                              <Button className="bg-blue-600 hover:bg-blue-700 text-white w-full">
                                Assign Selected
                              </Button>
                            </div>

                            {/* Add New Customer Codes */}
                            <div className="space-y-4">
                              <Label className="text-blue-600 font-medium">
                                Add New Customer Codes (one per line):
                              </Label>
                              <Textarea
                                value={newCustomerCodes}
                                onChange={(e) => setNewCustomerCodes(e.target.value)}
                                className="bg-gray-100 border-gray-300 min-h-[100px]"
                                placeholder="Enter customer codes, one per line"
                              />
                              <Button className="bg-blue-600 hover:bg-blue-700 text-white w-full">
                                Add Codes
                              </Button>
                            </div>
                          </div>
                        </div>
                      </div>
                    </CardContent>
                  </Card>
                </div>
              </TabsContent>
            </Tabs>
          </div>
        )}

        {/* Branch & Region Mapping Section */}
        {activeSection === "branch-region" && (
          <div className="space-y-6">
            {/* Sub Navigation */}
            <Tabs value={branchRegionTab} onValueChange={setBranchRegionTab} className="w-full">
              <div className="flex justify-center">
                <TabsList className="grid w-fit grid-cols-2 bg-white border">
                  <TabsTrigger 
                    value="manual-entry"
                    className="px-6 py-2 data-[state=active]:bg-blue-600 data-[state=active]:text-white"
                  >
                    Manual Entry
                  </TabsTrigger>
                  <TabsTrigger 
                    value="file-upload"
                    className="px-6 py-2 data-[state=active]:bg-blue-600 data-[state=active]:text-white"
                  >
                    File Upload
                  </TabsTrigger>
                </TabsList>
              </div>

              <TabsContent value="manual-entry" className="mt-8">
                <div className="grid lg:grid-cols-2 gap-8">
                  {/* Left Column - Create Branch and Region */}
                  <div className="space-y-8">
                    {/* Create Branch */}
                    <Card>
                      <CardHeader>
                        <CardTitle className="text-xl font-semibold text-blue-600">
                          Create Branch
                        </CardTitle>
                      </CardHeader>
                      <CardContent className="space-y-4">
                        <div className="space-y-2">
                          <Label className="text-blue-600 font-medium">
                            Branch Name:
                          </Label>
                          <Input
                            value={newBranch.name}
                            onChange={(e) => setNewBranch({name: e.target.value})}
                            className="bg-gray-100 border-gray-300"
                            placeholder="Enter branch name"
                          />
                        </div>
                        <Button 
                          onClick={handleAddBranch}
                          className="bg-blue-600 hover:bg-blue-700 text-white w-full"
                        >
                          Create Branch
                        </Button>
                      </CardContent>
                    </Card>

                    {/* Create Region */}
                    <Card>
                      <CardHeader>
                        <CardTitle className="text-xl font-semibold text-blue-600">
                          Create Region
                        </CardTitle>
                      </CardHeader>
                      <CardContent className="space-y-4">
                        <div className="space-y-2">
                          <Label className="text-blue-600 font-medium">
                            Region Name:
                          </Label>
                          <Input
                            value={newRegion.name}
                            onChange={(e) => setNewRegion({name: e.target.value})}
                            className="bg-gray-100 border-gray-300"
                            placeholder="Enter region name"
                          />
                        </div>
                        <Button 
                          onClick={handleAddRegion}
                          className="bg-blue-600 hover:bg-blue-700 text-white w-full"
                        >
                          Create Region
                        </Button>
                      </CardContent>
                    </Card>
                  </div>

                  {/* Right Column - Current Branches and Regions */}
                  <div className="space-y-8">
                    {/* Current Branches */}
                    <Card>
                      <CardHeader>
                        <CardTitle className="text-xl font-semibold text-blue-600">
                          Current Branches
                        </CardTitle>
                      </CardHeader>
                      <CardContent>
                        <div className="overflow-x-auto mb-4">
                          <Table>
                            <TableHeader>
                              <TableRow className="bg-gray-100">
                                <TableHead className="text-center font-semibold">Branch</TableHead>
                                <TableHead className="text-center font-semibold">Executives</TableHead>
                                <TableHead className="text-center font-semibold">In Regions</TableHead>
                              </TableRow>
                            </TableHeader>
                            <TableBody>
                              {branchesData.map((branch) => (
                                <TableRow key={branch.id}>
                                  <TableCell className="text-center">{branch.name}</TableCell>
                                  <TableCell className="text-center">{branch.executives}</TableCell>
                                  <TableCell className="text-center">{branch.regions}</TableCell>
                                </TableRow>
                              ))}
                            </TableBody>
                          </Table>
                        </div>
                        <div className="space-y-2">
                          <Label className="text-blue-600 font-medium">Remove Branches:</Label>
                          <Select value={removeBranches} onValueChange={setRemoveBranches}>
                            <SelectTrigger className="bg-gray-100">
                              <SelectValue placeholder="Select branches to remove" />
                            </SelectTrigger>
                            <SelectContent>
                              {branchesData.map((branch) => (
                                <SelectItem key={branch.id} value={branch.name}>
                                  {branch.name}
                                </SelectItem>
                              ))}
                            </SelectContent>
                          </Select>
                          <Button 
                            onClick={handleRemoveBranches}
                            className="bg-red-600 hover:bg-red-700 text-white w-full"
                          >
                            Remove Branch
                          </Button>
                        </div>
                      </CardContent>
                    </Card>

                    {/* Current Regions */}
                    <Card>
                      <CardHeader>
                        <CardTitle className="text-xl font-semibold text-blue-600">
                          Current Regions
                        </CardTitle>
                      </CardHeader>
                      <CardContent>
                        <div className="overflow-x-auto mb-4">
                          <Table>
                            <TableHeader>
                              <TableRow className="bg-gray-100">
                                <TableHead className="text-center font-semibold">Region</TableHead>
                                <TableHead className="text-center font-semibold">Branches</TableHead>
                              </TableRow>
                            </TableHeader>
                            <TableBody>
                              {regionsData.map((region) => (
                                <TableRow key={region.id}>
                                  <TableCell className="text-center">{region.name}</TableCell>
                                  <TableCell className="text-center">{region.branches}</TableCell>
                                </TableRow>
                              ))}
                            </TableBody>
                          </Table>
                        </div>
                        <div className="space-y-2">
                          <Label className="text-blue-600 font-medium">Remove Regions:</Label>
                          <Select value={removeRegions} onValueChange={setRemoveRegions}>
                            <SelectTrigger className="bg-gray-100">
                              <SelectValue placeholder="Select regions to remove" />
                            </SelectTrigger>
                            <SelectContent>
                              {regionsData.map((region) => (
                                <SelectItem key={region.id} value={region.name}>
                                  {region.name}
                                </SelectItem>
                              ))}
                            </SelectContent>
                          </Select>
                          <Button 
                            onClick={handleRemoveRegions}
                            className="bg-red-600 hover:bg-red-700 text-white w-full"
                          >
                            Remove Region
                          </Button>
                        </div>
                      </CardContent>
                    </Card>
                  </div>
                </div>

                {/* Mapping Sections */}
                <div className="grid lg:grid-cols-2 gap-8 mt-8">
                  {/* Map Executives to Branches */}
                  <Card>
                    <CardHeader>
                      <CardTitle className="text-xl font-semibold text-blue-600">
                        Map Executives to Branches
                      </CardTitle>
                    </CardHeader>
                    <CardContent className="space-y-4">
                      <div className="space-y-2">
                        <Label className="text-blue-600 font-medium">Select Branch:</Label>
                        <Select value={selectedBranch} onValueChange={setSelectedBranch}>
                          <SelectTrigger className="bg-gray-100">
                            <SelectValue placeholder="Select a branch" />
                          </SelectTrigger>
                          <SelectContent>
                            {branchesData.map((branch) => (
                              <SelectItem key={branch.id} value={branch.name}>
                                {branch.name}
                              </SelectItem>
                            ))}
                          </SelectContent>
                        </Select>
                      </div>
                      <div className="space-y-2">
                        <Label className="text-blue-600 font-medium">Select Executives:</Label>
                        <Select value={selectedExecutives} onValueChange={setSelectedExecutives}>
                          <SelectTrigger className="bg-gray-100">
                            <SelectValue placeholder="Select executives" />
                          </SelectTrigger>
                          <SelectContent>
                            {executivesData.map((executive) => (
                              <SelectItem key={executive.id} value={executive.code}>
                                {executive.code} - {executive.name}
                              </SelectItem>
                            ))}
                          </SelectContent>
                        </Select>
                      </div>
                      <Button className="bg-blue-600 hover:bg-blue-700 text-white w-full" onClick={handleUpdateBranchMapping}>
                        Update Branch Mapping
                      </Button>
                    </CardContent>
                  </Card>

                  {/* Map Executives to Regions */}
                  <Card>
                    <CardHeader>
                      <CardTitle className="text-xl font-semibold text-blue-600">
                        Map Executives to Regions
                      </CardTitle>
                    </CardHeader>
                    <CardContent className="space-y-4">
                      <div className="space-y-2">
                        <Label className="text-blue-600 font-medium">Select Region:</Label>
                        <Select value={selectedRegion} onValueChange={setSelectedRegion}>
                          <SelectTrigger className="bg-gray-100">
                            <SelectValue placeholder="Select a region" />
                          </SelectTrigger>
                          <SelectContent>
                            {regionsData.map((region) => (
                              <SelectItem key={region.id} value={region.name}>
                                {region.name}
                              </SelectItem>
                            ))}
                          </SelectContent>
                        </Select>
                      </div>
                      <div className="space-y-2">
                        <Label className="text-blue-600 font-medium">Select Branches:</Label>
                        <Select value={selectedBranches} onValueChange={setSelectedBranches}>
                          <SelectTrigger className="bg-gray-100">
                            <SelectValue placeholder="Select branches" />
                          </SelectTrigger>
                          <SelectContent>
                            {branchesData.map((branch) => (
                              <SelectItem key={branch.id} value={branch.name}>
                                {branch.name}
                              </SelectItem>
                            ))}
                          </SelectContent>
                        </Select>
                      </div>
                      <Button className="bg-blue-600 hover:bg-blue-700 text-white w-full" onClick={handleUpdateRegionMapping}>
                        Update Branch Mapping
                      </Button>
                    </CardContent>
                  </Card>
                </div>

                {/* Current Mapping Table */}
                <Card className="mt-8">
                  <CardHeader>
                    <CardTitle className="text-xl font-semibold text-blue-600">
                      Current Mapping
                    </CardTitle>
                  </CardHeader>
                  <CardContent>
                    <div className="overflow-x-auto">
                      <Table>
                        <TableHeader>
                          <TableRow className="bg-gray-100">
                            <TableHead className="text-center font-semibold">Branch</TableHead>
                            <TableHead className="text-center font-semibold">Region</TableHead>
                            <TableHead className="text-center font-semibold">Executives</TableHead>
                            <TableHead className="text-center font-semibold">Count</TableHead>
                          </TableRow>
                        </TableHeader>
                        <TableBody>
                          {mappingData.map((mapping, index) => (
                            <TableRow key={index}>
                              <TableCell className="text-center">{mapping.branch}</TableCell>
                              <TableCell className="text-center">{mapping.region}</TableCell>
                              <TableCell className="text-center">{mapping.executives.join(", ")}</TableCell>
                              <TableCell className="text-center">{mapping.count}</TableCell>
                            </TableRow>
                          ))}
                        </TableBody>
                      </Table>
                    </div>
                  </CardContent>
                </Card>
              </TabsContent>

              <TabsContent value="file-upload" className="mt-8">
                <Card>
                  <CardHeader>
                    <CardTitle className="text-xl font-semibold text-blue-600">
                      Upload Branch & Region Mapping File
                    </CardTitle>
                  </CardHeader>
                  <CardContent>
                    <div 
                      className="border-2 border-dashed border-blue-300 rounded-lg p-8 bg-blue-50 text-center relative"
                      onDrop={handleFileDrop}
                      onDragOver={handleDragOver}
                    >
                      <CloudUpload className="w-8 h-8 text-blue-600 mx-auto mb-2" />
                      <p className="text-blue-700 mb-4">
                        {selectedFile ? `Selected: ${selectedFile.name}` : "Drag and drop your mapping file here"}
                      </p>
                      <div className="space-y-2">
                        <input
                          type="file"
                          id="mapping-file-upload"
                          className="hidden"
                          accept=".csv,.xls,.xlsx"
                          onChange={handleFileSelect}
                        />
                        <label htmlFor="mapping-file-upload">
                          <Button 
                            variant="outline" 
                            className="bg-white text-blue-600 border-blue-300 hover:bg-blue-50"
                            asChild
                          >
                            <span className="cursor-pointer">Browse Files</span>
                          </Button>
                        </label>
                        {selectedFile && (
                          <div className="mt-4">
                            <Button 
                              onClick={handleFileUpload}
                              className="bg-blue-600 hover:bg-blue-700 text-white"
                            >
                              Upload Mapping File
                            </Button>
                          </div>
                        )}
                      </div>
                      <p className="text-xs text-gray-500 mt-2">
                        Supported formats: CSV, Excel (.csv, .xls, .xlsx) - Max 5MB
                      </p>
                    </div>
                  </CardContent>
                </Card>
              </TabsContent>
            </Tabs>
          </div>
        )}

        {/* Company & Product Mapping Section */}
        {activeSection === "company-product" && (
          <div className="space-y-6">
            {/* Sub Navigation */}
            <Tabs value={companyProductTab} onValueChange={setCompanyProductTab} className="w-full">
              <div className="flex justify-center">
                <TabsList className="grid w-fit grid-cols-2 bg-white border">
                  <TabsTrigger 
                    value="manual-entry"
                    className="px-6 py-2 data-[state=active]:bg-blue-600 data-[state=active]:text-white"
                  >
                    Manual Entry
                  </TabsTrigger>
                  <TabsTrigger 
                    value="file-upload"
                    className="px-6 py-2 data-[state=active]:bg-blue-600 data-[state=active]:text-white"
                  >
                    File Upload
                  </TabsTrigger>
                </TabsList>
              </div>

              <TabsContent value="manual-entry" className="mt-8">
                <div className="grid lg:grid-cols-2 gap-8">
                  {/* Left Column - Create Product and Company Groups */}
                  <div className="space-y-8">
                    {/* Create Product Group */}
                    <Card>
                      <CardHeader>
                        <CardTitle className="text-xl font-semibold text-blue-600">
                          Create Product Group
                        </CardTitle>
                      </CardHeader>
                      <CardContent className="space-y-4">
                        <div className="space-y-2">
                          <Label className="text-blue-600 font-medium">
                            Product Group Name:
                          </Label>
                          <Input
                            value={newProductGroup.name}
                            onChange={(e) => setNewProductGroup({name: e.target.value})}
                            className="bg-gray-100 border-gray-300"
                            placeholder="Enter product group name"
                          />
                        </div>
                        <Button 
                          onClick={handleAddProductGroup}
                          className="bg-blue-600 hover:bg-blue-700 text-white w-full"
                        >
                          Create Product
                        </Button>
                      </CardContent>
                    </Card>

                    {/* Create Company Group */}
                    <Card>
                      <CardHeader>
                        <CardTitle className="text-xl font-semibold text-blue-600">
                          Create Company Group
                        </CardTitle>
                      </CardHeader>
                      <CardContent className="space-y-4">
                        <div className="space-y-2">
                          <Label className="text-blue-600 font-medium">
                            Company Group Name:
                          </Label>
                          <Input
                            value={newCompanyGroup.name}
                            onChange={(e) => setNewCompanyGroup({name: e.target.value})}
                            className="bg-gray-100 border-gray-300"
                            placeholder="Enter company group name"
                          />
                        </div>
                        <Button 
                          onClick={handleAddCompanyGroup}
                          className="bg-blue-600 hover:bg-blue-700 text-white w-full"
                        >
                          Create Company
                        </Button>
                      </CardContent>
                    </Card>
                  </div>

                  {/* Right Column - Current Products and Companies */}
                  <div className="space-y-8">
                    {/* Current Products */}
                    <Card>
                      <CardHeader>
                        <CardTitle className="text-xl font-semibold text-blue-600">
                          Current Products
                        </CardTitle>
                      </CardHeader>
                      <CardContent>
                        <div className="overflow-x-auto mb-4">
                          <Table>
                            <TableHeader>
                              <TableRow className="bg-gray-100">
                                <TableHead className="text-center font-semibold">Product</TableHead>
                                <TableHead className="text-center font-semibold">In Companies</TableHead>
                              </TableRow>
                            </TableHeader>
                            <TableBody>
                              {productsData.map((product) => (
                                <TableRow key={product.id}>
                                  <TableCell className="text-center">{product.name}</TableCell>
                                  <TableCell className="text-center">{product.companies}</TableCell>
                                </TableRow>
                              ))}
                            </TableBody>
                          </Table>
                        </div>
                        <div className="space-y-2">
                          <Label className="text-blue-600 font-medium">Remove Products:</Label>
                          <Select value={removeProducts} onValueChange={setRemoveProducts}>
                            <SelectTrigger className="bg-gray-100">
                              <SelectValue placeholder="Select products to remove" />
                            </SelectTrigger>
                            <SelectContent>
                              {productsData.map((product) => (
                                <SelectItem key={product.id} value={product.name}>
                                  {product.name}
                                </SelectItem>
                              ))}
                            </SelectContent>
                          </Select>
                          <Button 
                            onClick={handleRemoveProducts}
                            className="bg-red-600 hover:bg-red-700 text-white w-full"
                          >
                            Remove Product
                          </Button>
                        </div>
                      </CardContent>
                    </Card>

                    {/* Current Companies */}
                    <Card>
                      <CardHeader>
                        <CardTitle className="text-xl font-semibold text-blue-600">
                          Current Companies
                        </CardTitle>
                      </CardHeader>
                      <CardContent>
                        <div className="overflow-x-auto mb-4">
                          <Table>
                            <TableHeader>
                              <TableRow className="bg-gray-100">
                                <TableHead className="text-center font-semibold">Company</TableHead>
                                <TableHead className="text-center font-semibold">Products</TableHead>
                              </TableRow>
                            </TableHeader>
                            <TableBody>
                              {companiesData.map((company) => (
                                <TableRow key={company.id}>
                                  <TableCell className="text-center">{company.name}</TableCell>
                                  <TableCell className="text-center">{company.products}</TableCell>
                                </TableRow>
                              ))}
                            </TableBody>
                          </Table>
                        </div>
                        <div className="space-y-2">
                          <Label className="text-blue-600 font-medium">Remove Company:</Label>
                          <Select value={removeCompanies} onValueChange={setRemoveCompanies}>
                            <SelectTrigger className="bg-gray-100">
                              <SelectValue placeholder="Select companies to remove" />
                            </SelectTrigger>
                            <SelectContent>
                              {companiesData.map((company) => (
                                <SelectItem key={company.id} value={company.name}>
                                  {company.name}
                                </SelectItem>
                              ))}
                            </SelectContent>
                          </Select>
                          <Button 
                            onClick={handleRemoveCompanies}
                            className="bg-red-600 hover:bg-red-700 text-white w-full"
                          >
                            Remove Company
                          </Button>
                        </div>
                      </CardContent>
                    </Card>
                  </div>
                </div>

                {/* Map Executives to Companies */}
                <Card className="mt-8">
                  <CardHeader>
                    <CardTitle className="text-xl font-semibold text-blue-600">
                      Map Executives to Companies
                    </CardTitle>
                  </CardHeader>
                  <CardContent className="space-y-4">
                    <div className="grid lg:grid-cols-2 gap-6">
                      <div className="space-y-2">
                        <Label className="text-blue-600 font-medium">Select Branch:</Label>
                        <Select value={selectedBranch} onValueChange={setSelectedBranch}>
                          <SelectTrigger className="bg-gray-100">
                            <SelectValue placeholder="Select a branch" />
                          </SelectTrigger>
                          <SelectContent>
                            {branchesData.map((branch) => (
                              <SelectItem key={branch.id} value={branch.name}>
                                {branch.name}
                              </SelectItem>
                            ))}
                          </SelectContent>
                        </Select>
                      </div>
                      <div className="space-y-2">
                        <Label className="text-blue-600 font-medium">Select Executives:</Label>
                        <Select value={selectedExecutivesForCompany} onValueChange={setSelectedExecutivesForCompany}>
                          <SelectTrigger className="bg-gray-100">
                            <SelectValue placeholder="Select executives" />
                          </SelectTrigger>
                          <SelectContent>
                            {executivesData.map((executive) => (
                              <SelectItem key={executive.id} value={executive.code}>
                                {executive.code} - {executive.name}
                              </SelectItem>
                            ))}
                          </SelectContent>
                        </Select>
                      </div>
                    </div>
                    <Button className="bg-blue-600 hover:bg-blue-700 text-white w-full">
                      Update Company Mapping
                    </Button>
                  </CardContent>
                </Card>

                {/* Current Mappings */}
                <Card className="mt-8">
                  <CardHeader>
                    <CardTitle className="text-xl font-semibold text-blue-600">
                      Current Mappings
                    </CardTitle>
                  </CardHeader>
                  <CardContent>
                    <div className="overflow-x-auto">
                      <Table>
                        <TableHeader>
                          <TableRow className="bg-gray-100">
                            <TableHead className="text-center font-semibold">Company</TableHead>
                            <TableHead className="text-center font-semibold">Products</TableHead>
                            <TableHead className="text-center font-semibold">Count</TableHead>
                          </TableRow>
                        </TableHeader>
                        <TableBody>
                          {companyMappingData.map((mapping, index) => (
                            <TableRow key={index}>
                              <TableCell className="text-center">{mapping.company}</TableCell>
                              <TableCell className="text-center">{mapping.products.join(", ")}</TableCell>
                              <TableCell className="text-center">{mapping.count}</TableCell>
                            </TableRow>
                          ))}
                        </TableBody>
                      </Table>
                    </div>
                  </CardContent>
                </Card>
              </TabsContent>

              <TabsContent value="file-upload" className="mt-8">
                <Card>
                  <CardHeader>
                    <CardTitle className="text-xl font-semibold text-blue-600">
                      Upload Company & Product Mapping File
                    </CardTitle>
                  </CardHeader>
                  <CardContent>
                    <div 
                      className="border-2 border-dashed border-blue-300 rounded-lg p-8 bg-blue-50 text-center relative"
                      onDrop={handleFileDrop}
                      onDragOver={handleDragOver}
                    >
                      <CloudUpload className="w-8 h-8 text-blue-600 mx-auto mb-2" />
                      <p className="text-blue-700 mb-4">
                        {selectedFile ? `Selected: ${selectedFile.name}` : "Drag and drop your company & product mapping file here"}
                      </p>
                      <div className="space-y-2">
                        <input
                          type="file"
                          id="company-product-file-upload"
                          className="hidden"
                          accept=".csv,.xls,.xlsx"
                          onChange={handleFileSelect}
                        />
                        <label htmlFor="company-product-file-upload">
                          <Button 
                            variant="outline" 
                            className="bg-white text-blue-600 border-blue-300 hover:bg-blue-50"
                            asChild
                          >
                            <span className="cursor-pointer">Browse Files</span>
                          </Button>
                        </label>
                        {selectedFile && (
                          <div className="mt-4">
                            <Button 
                              onClick={handleFileUpload}
                              className="bg-blue-600 hover:bg-blue-700 text-white"
                            >
                              Upload Mapping File
                            </Button>
                          </div>
                        )}
                      </div>
                      <p className="text-xs text-gray-500 mt-2">
                        Supported formats: CSV, Excel (.csv, .xls, .xlsx) - Max 5MB
                      </p>
                    </div>
                  </CardContent>
                </Card>
              </TabsContent>
            </Tabs>
          </div>
        )}

        {/* Backup & Restore Section */}
        {activeSection === "backup-restore" && (
          <BackupRestore />
        )}

        {/* File Processing Section */}
        {activeSection === "file-processing" && (
          <div>
            <FileProcessing />
          </div>
        )}
   
      </div>

      {/* Consolidated Data View Section */}
      {activeSection === "consolidated-data" && (
        <div className="space-y-6">
          
          {/* Filters Container */}
          <div className="bg-white rounded-lg shadow p-6 mb-6">
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Filter by Executive:
                </label>
                <Select value={selectedExecutive} onValueChange={setSelectedExecutive}>
                  <SelectTrigger className="w-full">
                    <SelectValue placeholder="Select Executive" />
                  </SelectTrigger>
                  <SelectContent>
                    {executives.map(exec => (
                      <SelectItem key={exec} value={exec}>{exec}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Filter by Branch:
                </label>
                <Select value={selectedBranch} onValueChange={setSelectedBranch}>
                  <SelectTrigger className="w-full">
                    <SelectValue placeholder="Select Branch" />
                  </SelectTrigger>
                  <SelectContent>
                    {branches.map(branch => (
                      <SelectItem key={branch} value={branch}>{branch}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Filter by Region:
                </label>
                <Select value={selectedRegion} onValueChange={setSelectedRegion}>
                  <SelectTrigger className="w-full">
                    <SelectValue placeholder="Select Region" />
                  </SelectTrigger>
                  <SelectContent>
                    {regions.map(region => (
                      <SelectItem key={region} value={region}>{region}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
            </div>
          </div>

          {/* Table Container */}
          <div className="bg-white rounded-lg shadow p-6">
            <div className="flex justify-between items-center mb-4">
              <p className="text-sm text-gray-600">
                Results: <span className="font-semibold">{filteredData.length} records</span>
              </p>
              <Button onClick={handleDownloadCSV} className="bg-blue-600 hover:bg-blue-700">
                <Download className="w-4 h-4 mr-2" />
                Download CSV
              </Button>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full border-collapse border border-gray-300">
                <thead>
                  <tr className="bg-gray-100">
                    <th className="border border-gray-300 px-4 py-3 text-left text-sm font-medium text-gray-700">
                      Customer Code
                    </th>
                    <th className="border border-gray-300 px-4 py-3 text-left text-sm font-medium text-gray-700">
                      Customer Name
                    </th>
                    <th className="border border-gray-300 px-4 py-3 text-left text-sm font-medium text-gray-700">
                      Executive
                    </th>
                    <th className="border border-gray-300 px-4 py-3 text-left text-sm font-medium text-gray-700">
                      Executive Code
                    </th>
                    <th className="border border-gray-300 px-4 py-3 text-left text-sm font-medium text-gray-700">
                      Branch
                    </th>
                    <th className="border border-gray-300 px-4 py-3 text-left text-sm font-medium text-gray-700">
                      Region
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {filteredData.map((row, index) => (
                    <tr key={index} className="hover:bg-gray-50">
                      <td className="border border-gray-300 px-4 py-3 text-sm text-gray-900">
                        {row.customerCode}
                      </td>
                      <td className="border border-gray-300 px-4 py-3 text-sm text-gray-900">
                        {row.customerName}
                      </td>
                      <td className="border border-gray-300 px-4 py-3 text-sm text-gray-900">
                        {row.executive}
                      </td>
                      <td className="border border-gray-300 px-4 py-3 text-sm text-gray-900">
                        {row.executiveCode}
                      </td>
                      <td className="border border-gray-300 px-4 py-3 text-sm text-gray-900">
                        {row.branch}
                      </td>
                      <td className="border border-gray-300 px-4 py-3 text-sm text-gray-900">
                        {row.region}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* Footer */}
      <div className="bg-white border-t py-4">
        <div className="container mx-auto px-6">
          <p className="text-center text-gray-500 text-sm">
            &copy; 2023 Executive Mapping Portal. All rights reserved.
          </p>
        </div>
      </div>
    </div>
  );
};

export default Index;

export const FileProcessing = () => {
  const [activeSubTab, setActiveSubTab] = useState("Budget Processing");

  const subTabs = [
    "Budget Processing",
    "Sales Processing", 
    "OS Processing"
  ];

  const handleFileUpload = () => {
    toast.success("File upload initiated!");
  };

  const handleBrowseFiles = () => {
    // Create a file input element and trigger click
    const input = document.createElement('input');
    input.type = 'file';
    input.multiple = true;
    input.accept = '.csv,.xlsx,.xls';
    input.onchange = (e) => {
      const files = (e.target as HTMLInputElement).files;
      if (files && files.length > 0) {
        toast.success(`Selected ${files.length} file(s) for upload`);
      }
    };
    input.click();
  };

  const handleDragOver = (e) => {
    e.preventDefault();
    e.stopPropagation();
  };

  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      toast.success(`Dropped ${files.length} file(s) for upload`);
    }
  };

  const getUploadText = (tabType) => {
    switch (tabType) {
      case "Budget Processing":
        return "Drag and drop your budget file here";
      case "Sales Processing":
        return "Drag and drop your sales file here";
      case "OS Processing":
        return "Drag and drop your OS file here";
      default:
        return "Drag and drop your file here";
    }
  };

  const renderTabContent = () => {
    switch (activeSubTab) {
      case "Budget Processing":
      case "Sales Processing":
      case "OS Processing":
        return (
          <div className="space-y-6">
            <h3 className="text-xl font-semibold text-blue-600">
              Process {activeSubTab.replace(" Processing", "")} File
            </h3>
            {/* Drag and Drop Area */}
            <div 
              className="bg-blue-50 border-2 border-dashed border-blue-300 rounded-lg p-12 text-center transition-colors hover:bg-blue-100"
              onDragOver={handleDragOver}
              onDrop={handleDrop}
            >
              <div className="flex flex-col items-center space-y-4">
                <CloudUpload className="w-12 h-12 text-blue-500" />
                <div>
                  <p className="text-lg font-medium text-blue-600">
                    {getUploadText(activeSubTab)}
                  </p>
                  <p className="text-gray-500 text-sm mt-2">
                    Supported formats: CSV, Excel (.csv, .xls, .xlsx) - Max 5MB
                  </p>
                </div>
                <Button 
                  onClick={handleBrowseFiles}
                  variant="outline"
                  className="bg-white text-blue-600 hover:bg-blue-50 border-blue-300 mt-4"
                >
                  Browse Files
                </Button>
              </div>
            </div>
          </div>
        );
      default:
        return null;
    }
  };

  return (
    <div className="space-y-6">
      <h2 className="text-2xl font-bold text-blue-600 flex items-center">
      </h2>
      
      {/* Sub-tab Navigation */}
      <div className="flex space-x-1 bg-gray-100 rounded-lg p-1">
        {subTabs.map((tab) => (
          <button
            key={tab}
            onClick={() => setActiveSubTab(tab)}
            className={cn(
              "px-4 py-2 rounded-md text-sm font-medium transition-all duration-200 flex-1",
              activeSubTab === tab
                ? "bg-blue-600 text-white shadow-sm"
                : "text-gray-600 hover:text-gray-900 hover:bg-white"
            )}
          >
            {tab}
          </button>
        ))}
      </div>

      {/* Content Area */}
      <div className="bg-white rounded-lg border border-gray-200 p-6">
        {renderTabContent()}
      </div>
    </div>
  );
};

export function MapExecutivesToCompanies() {
  // State for mappings
  const [companyMappings, setCompanyMappings] = useState([
    // Example: { executive: "John Smith", company: "Acme Corporation" }
  ]);
  // State for form
  const [selectedExecutive, setSelectedExecutive] = useState("");
  const [selectedCompany, setSelectedCompany] = useState("");

  // Update handler
  const handleUpdateCompanyMapping = () => {
    if (!selectedExecutive || !selectedCompany) return;
    // Check if mapping exists
    const idx = companyMappings.findIndex(
      m => m.executive === selectedExecutive && m.company === selectedCompany
    );
    if (idx === -1) {
      setCompanyMappings([
        ...companyMappings,
        { executive: selectedExecutive, company: selectedCompany }
      ]);
    }
    // Optionally, show a toast or clear form here
  };

  return (
    <div>
      <h2>Map Executives to Companies</h2>
      <div>
        <select value={selectedExecutive} onChange={e => setSelectedExecutive(e.target.value)}>
          <option value="">Select Executive</option>
          <option value="John Smith">John Smith</option>
          <option value="Sarah Johnson">Sarah Johnson</option>
        </select>
        <select value={selectedCompany} onChange={e => setSelectedCompany(e.target.value)}>
          <option value="">Select Company</option>
          <option value="Acme Corporation">Acme Corporation</option>
          <option value="Global Tech Solutions">Global Tech Solutions</option>
        </select>
        <button onClick={handleUpdateCompanyMapping}>Update Company Mapping</button>
      </div>

      <h3>Current Mappings</h3>
      <table border={1} cellPadding={8}>
        <thead>
          <tr>
            <th>Executive</th>
            <th>Company</th>
          </tr>
        </thead>
        <tbody>
          {companyMappings.map((row, idx) => (
            <tr key={idx}>
              <td>{row.executive}</td>
              <td>{row.company}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
