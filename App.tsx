
import React, { useState, useCallback } from 'react';
import { Lead } from './types';
import { processExcelFiles, exportToExcel } from './services/excelService';
import { UploadIcon } from './components/icons/UploadIcon';
import { DownloadIcon } from './components/icons/DownloadIcon';
import { FileIcon } from './components/icons/FileIcon';
import { SpinnerIcon } from './components/icons/SpinnerIcon';


interface FileUploaderProps {
  onFileSelect: (files: FileList) => void;
  disabled: boolean;
}

const FileUploader: React.FC<FileUploaderProps> = ({ onFileSelect, disabled }) => {
    const [isDragging, setIsDragging] = useState(false);

    const handleDragEnter = (e: React.DragEvent<HTMLLabelElement>) => {
        e.preventDefault();
        e.stopPropagation();
        if (!disabled) setIsDragging(true);
    };

    const handleDragLeave = (e: React.DragEvent<HTMLLabelElement>) => {
        e.preventDefault();
        e.stopPropagation();
        setIsDragging(false);
    };

    const handleDragOver = (e: React.DragEvent<HTMLLabelElement>) => {
        e.preventDefault();
        e.stopPropagation();
    };

    const handleDrop = (e: React.DragEvent<HTMLLabelElement>) => {
        e.preventDefault();
        e.stopPropagation();
        setIsDragging(false);
        if (!disabled && e.dataTransfer.files && e.dataTransfer.files.length > 0) {
            onFileSelect(e.dataTransfer.files);
        }
    };

    const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        if (e.target.files && e.target.files.length > 0) {
            onFileSelect(e.target.files);
        }
    };
    
    return (
        <label
            onDragEnter={handleDragEnter}
            onDragLeave={handleDragLeave}
            onDragOver={handleDragOver}
            onDrop={handleDrop}
            className={`flex justify-center w-full h-48 px-4 transition bg-white border-2 border-gray-300 border-dashed rounded-md appearance-none cursor-pointer hover:border-indigo-400 focus:outline-none ${isDragging ? 'border-indigo-500 bg-indigo-50' : ''} ${disabled ? 'cursor-not-allowed opacity-50' : ''}`}
        >
            <span className="flex flex-col items-center justify-center space-y-2 text-center">
                <UploadIcon className="w-12 h-12 text-gray-500"/>
                <span className="font-medium text-gray-600">
                    Drop files to attach, or <span className="text-indigo-600 underline">browse</span>
                </span>
                <span className="text-xs text-gray-500">Supports .xlsx and .xls files</span>
            </span>
            <input
                type="file"
                name="file_upload"
                className="hidden"
                multiple
                accept=".xlsx, .xls"
                onChange={handleFileChange}
                disabled={disabled}
            />
        </label>
    );
};


const App: React.FC = () => {
    const [files, setFiles] = useState<FileList | null>(null);
    const [leads, setLeads] = useState<Lead[]>([]);
    const [isLoading, setIsLoading] = useState<boolean>(false);
    const [error, setError] = useState<string | null>(null);

    const handleFileSelect = (selectedFiles: FileList) => {
        setFiles(selectedFiles);
        setError(null);
        setLeads([]);
    };

    const handleProcessFiles = useCallback(async () => {
        if (!files || files.length === 0) {
            setError('Please select at least one file.');
            return;
        }

        setIsLoading(true);
        setError(null);
        setLeads([]);

        try {
            const extractedLeads = await processExcelFiles(files);
            if(extractedLeads.length === 0){
                setError("No leads found in the selected files. Please check column names for 'Name', 'Email', or 'Mobile'.");
            }
            setLeads(extractedLeads);
        } catch (err) {
            // FIX: Use String(err) for robust error message handling.
            // This avoids property access on an 'unknown' type and safely converts
            // any thrown value (Error objects, strings, etc.) to a string.
            setError(String(err));
        } finally {
            setIsLoading(false);
        }
    }, [files]);
    
    const handleDownload = () => {
        if(leads.length > 0){
            exportToExcel(leads, 'consolidated_leads');
        }
    };

    const handleReset = () => {
        setFiles(null);
        setLeads([]);
        setError(null);
        setIsLoading(false);
    };

    return (
        <div className="bg-slate-100 min-h-screen font-sans text-gray-800 flex items-center justify-center p-4">
            <div className="w-full max-w-3xl mx-auto">
                <header className="text-center mb-8">
                    <h1 className="text-4xl font-bold text-slate-800">Leads Consolidator</h1>
                    <p className="text-slate-600 mt-2">Merge Facebook lead sheets into one master file.</p>
                </header>

                <main className="bg-white rounded-xl shadow-lg p-8 space-y-6">
                    {leads.length === 0 && !isLoading && (
                        <div className="space-y-4">
                            <FileUploader onFileSelect={handleFileSelect} disabled={isLoading} />
                            {files && files.length > 0 && (
                                <div className="p-4 border rounded-md bg-gray-50">
                                    <h3 className="font-semibold text-gray-700 mb-2">Selected Files:</h3>
                                    <ul className="space-y-1 text-sm text-gray-600">
                                        {Array.from(files).map((file, index) => (
                                            <li key={index} className="flex items-center">
                                               <FileIcon className="w-4 h-4 mr-2 text-gray-400"/> {file.name}
                                            </li>
                                        ))}
                                    </ul>
                                </div>
                            )}
                        </div>
                    )}

                    {error && (
                        <div className="bg-red-100 border-l-4 border-red-500 text-red-700 p-4 rounded-md" role="alert">
                            <p className="font-bold">Error</p>
                            <p>{error}</p>
                        </div>
                    )}
                    
                    {isLoading && (
                        <div className="flex flex-col items-center justify-center text-center p-8 space-y-4">
                            <SpinnerIcon className="animate-spin h-12 w-12 text-indigo-600"/>
                            <p className="text-lg font-medium text-slate-700">Processing files...</p>
                            <p className="text-sm text-slate-500">Extracting and consolidating leads. Please wait.</p>
                        </div>
                    )}

                    {leads.length > 0 && !isLoading && (
                        <div className="space-y-4">
                            <div className="bg-green-100 border-l-4 border-green-500 text-green-800 p-4 rounded-md">
                                <p className="font-bold">Success!</p>
                                <p>Found {leads.length} leads across {files?.length} file(s). You can now download the consolidated file.</p>
                            </div>

                            <h3 className="text-lg font-semibold text-gray-800">Preview (First 10 Leads)</h3>
                            <div className="overflow-x-auto border rounded-lg">
                                <table className="min-w-full divide-y divide-gray-200">
                                    <thead className="bg-gray-50">
                                        <tr>
                                            <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Name</th>
                                            <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Email</th>
                                            <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Mobile</th>
                                        </tr>
                                    </thead>
                                    <tbody className="bg-white divide-y divide-gray-200">
                                        {leads.slice(0, 10).map((lead, index) => (
                                            <tr key={index} className="hover:bg-gray-50">
                                                <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-700">{lead.name || 'N/A'}</td>
                                                <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-700">{lead.email || 'N/A'}</td>
                                                <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-700">{lead.mobile || 'N/A'}</td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    )}

                    <div className="flex items-center justify-between pt-4">
                        <button
                            onClick={handleReset}
                            className="px-6 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-md shadow-sm hover:bg-gray-50 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500 disabled:opacity-50"
                            disabled={isLoading}
                        >
                            Reset
                        </button>

                       {leads.length === 0 ? (
                            <button
                                onClick={handleProcessFiles}
                                disabled={!files || files.length === 0 || isLoading}
                                className="inline-flex items-center justify-center px-6 py-2 border border-transparent text-sm font-medium rounded-md shadow-sm text-white bg-indigo-600 hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500 disabled:bg-indigo-300 disabled:cursor-not-allowed"
                            >
                               {isLoading ? <SpinnerIcon className="animate-spin -ml-1 mr-3 h-5 w-5 text-white"/> : <UploadIcon className="-ml-1 mr-2 h-5 w-5" />}
                                Process Files
                            </button>
                       ) : (
                            <button
                                onClick={handleDownload}
                                className="inline-flex items-center justify-center px-6 py-2 border border-transparent text-sm font-medium rounded-md shadow-sm text-white bg-green-600 hover:bg-green-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-green-500"
                            >
                                <DownloadIcon className="-ml-1 mr-2 h-5 w-5"/>
                                Download Consolidated File
                            </button>
                       )}
                    </div>
                </main>

                <footer className="text-center mt-8 text-sm text-slate-500">
                    <p>&copy; {new Date().getFullYear()} Lead Consolidator. Built for efficiency.</p>
                </footer>
            </div>
        </div>
    );
};

export default App;
