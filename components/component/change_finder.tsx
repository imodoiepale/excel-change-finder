// @ts-nocheck
// @ts-ignore
import { useEffect, useState } from 'react';
import Link from "next/link";
import { Label } from "@/components/ui/label";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import { Progress } from "@/components/ui/progress";
import { Loader2 } from "lucide-react";
import toast, { Toaster } from 'react-hot-toast';

export function Change_Finder() {
  const [mainFile, setMainFile] = useState(null);
  const [variantFile, setVariantFile] = useState(null);
  const [progress, setProgress] = useState(0);
  const [isLoading, setIsLoading] = useState(false);
  const [consoleLog, setConsoleLog] = useState('');

  useEffect(() => {
    console.log('EventSource connection established');
    const eventSource = new EventSource('/api/progress');
  
    eventSource.onopen = () => {
      console.log('EventSource connection opened');
    };
  
    eventSource.onmessage = (event) => {
      const data = JSON.parse(event.data);
      console.log('Received progress update:', data);
      setProgress(data.progress);
      setConsoleLog(data.consoleLog || '');
    };
  
    eventSource.onerror = (error) => {
      console.error('EventSource error:', error);
    };
  
    return () => {
      console.log('Cleanup function: EventSource connection closed');
      eventSource.onopen = null;
      eventSource.onmessage = null;
      eventSource.onerror = null;
      eventSource.close();
    };
  }, []);
  
  const handleFileUpload = async () => {
    if (!mainFile || !variantFile) {
      alert("Please select both main and variant Excel files.");
      return;
    }

    setIsLoading(true);

    const formData = new FormData();
    formData.append("mainFile", mainFile);
    formData.append("variantFile", variantFile);

    try {
      const response = await fetch('/api/excel', {
        method: 'POST',
        body: formData,
      });
  
      if (!response.ok) {
        throw new Error('Network response was not ok');
      }
  
      // const blob = await response.blob();
      // const url = window.URL.createObjectURL(blob);
      // const a = document.createElement('a');
      // a.href = url;
      // a.download = 'Comparison Results.xlsx';
      // document.body.appendChild(a);
      // a.click();
      // window.URL.revokeObjectURL(url);

      toast.success('Successfully Compared Files!');
    } catch (error) {
      console.error('Error:', error);
      alert('An error occurred during file upload.');
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className="flex min-h-screen flex-col">
      <Toaster position="bottom-right" reverseOrder={false} />
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
            <Button
              className="w-full"
              onClick={handleFileUpload}
              disabled={isLoading}
            >
              {isLoading ? (
                <>
                  <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                  Comparing...
                </>
              ) : (
                "Compare Files"
              )}
            </Button>
            <div className=" justify-center text-center flex space-x-6 mt-6">
              <div>
                <Link href="/Comparison Results.xlsx" passHref>
                  <Button className="w-full ">Download Compared File</Button>
                </Link>
              </div>
              <div>
                <Button className="w-full" onClick={() => window.location.reload()}>
                  Refresh Page
                </Button>
              </div>
            </div>
          </div>
        </div>
      </main>
      <footer className="bg-gray-900 py-4 px-6 text-white">
        <div className="container mx-auto flex items-center justify-center">
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
