// @ts-nocheck
// @ts-ignore
import Link from "next/link";
import { Label } from "@/components/ui/label";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import { Progress } from "@/components/ui/progress";
import { useState } from "react";

export function Change_Finder() {
  const [mainFile, setMainFile] = useState(null);
  const [variantFile, setVariantFile] = useState(null);
  const [progress, setProgress] = useState(0);

  const handleFileUpload = async () => {
    if (!mainFile || !variantFile) {
      alert("Please select both main and variant Excel files.");
      return;
    }

    const formData = new FormData();
    formData.append("mainFile", mainFile);
    formData.append("variantFile", variantFile);

    try {
      function generateBoundary() {
        return (
          "---------------------------" +
          Math.floor(Math.random() * Math.pow(10, 15)).toString(36)
        );
      }
      
      const boundary = generateBoundary();
      const headers = {
        'Content-Type': `multipart/form-data; boundary=${boundary}`
      };
    
      const response = await fetch("/api/excel", {
        method: "POST", 
        body: formData,
        headers: headers,
      });
    
      const data = await response.json();
      console.log(data);
      // Handle the response from the API route
    } catch (error) {
      console.error("Error:", error);
    }
  };

  return (
    <div className="flex min-h-screen flex-col">
      <header className="bg-gray-900 py-4 px-6 text-white">
        <div className="container mx-auto flex items-center justify-between">
          <div className="flex items-center space-x-4">
            <FileSpreadsheetIcon className="h-6 w-6" />
            <h1 className="text-2xl font-bold">Excel Comparator</h1>
          </div>
        </div>
      </header>
      <main className="flex-1 bg-gray-100 py-12 px-6 dark:bg-gray-900 ">
        <div className="container mx-auto max-w-3xl space-y-8">
          <div className="space-y-4">
            <h2 className="text-3xl font-bold text-gray-900 dark:text-white">
              Compare Excel Files
            </h2>
            <p className="text-gray-600 dark:text-gray-400">
              Upload your main and variant Excel files to compare the changes.
            </p>
          </div>
          <div className="space-y-6">
            <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
              <div>
                <Label htmlFor="main-file">Main File</Label>
                <Input
                  accept=".xlsx"
                  className="mt-1 block w-full"
                  id="main-file"
                  name="mainFile"
                  type="file"
                  onChange={(e) => setMainFile(e.target.files[0])}
                />
              </div>
              <div>
                <Label htmlFor="variant-file">Variant File</Label>
                <Input
                  accept=".xlsx"
                  className="mt-1 block w-full"
                  id="variant-file"
                  name="variantFile"
                  type="file"
                  onChange={(e) => setVariantFile(e.target.files[0])}
                />
              </div>
            </div>
            <Button className="w-full" onClick={handleFileUpload}>
              Compare Files
            </Button>
            <div className="text-gray-900 dark:text-white">
              Comparison Progress:
            </div>
            <div className="space-y-2 text-center">
              <Progress
                className="h-2 bg-gray-300 dark:bg-gray-800"
                value={progress}
              />
              <div className="text-gray-600 dark:text-gray-400">
                {progress}% Complete
              </div>
            </div>
          </div>
        </div>
      </main>
      <footer className="bg-gray-900 py-4 px-6 text-white">
        <div className="container mx-auto flex items-center ">
          <p>
            © 2024 Excel Comparator. All rights reserved by{" "}
            <Link
              className="hover:underline"
              href="http://booksmartconsult.com/"
            >
              &nbsp;BCL
            </Link>
          </p>
        </div>
      </footer>
    </div>
  );
}

function FileSpreadsheetIcon(props) {
  return (
    <svg
      {...props}
      xmlns="http://www.w3.org/2000/svg"
      width="24"
      height="24"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
      strokeLinecap="round"
      strokeLinejoin="round"
    >
      <path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z" />
      <path d="M14 2v4a2 2 0 0 0 2 2h4" />
      <path d="M8 13h2" />
      <path d="M14 13h2" />
      <path d="M8 17h2" />
      <path d="M14 17h2" />
    </svg>
  );
}