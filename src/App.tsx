import React, { useState, useCallback, useMemo } from "react";
import { 
  Upload, 
  FileText, 
  Trash2, 
  Download, 
  Settings2, 
  History, 
  CheckCircle2, 
  AlertCircle, 
  ChevronRight, 
  FileSpreadsheet,
  Database,
  RefreshCw,
  X
} from "lucide-react";
import { motion, AnimatePresence } from "motion/react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import yaml from "js-yaml";
import { cn } from "./lib/utils";
import { 
  CleaningConfig, 
  FileData, 
  RunHistory, 
  NullOption, 
  BlankOption, 
  ZeroOption, 
  DuplicateOption 
} from "./types";

export default function App() {
  const [files, setFiles] = useState<FileData[]>([]);
  const [selectedFileId, setSelectedFileId] = useState<string | null>(null);
  const [history, setHistory] = useState<RunHistory[]>([]);
  const [config, setConfig] = useState<CleaningConfig>({
    nulls: "leave",
    blanks: "leave",
    zeros: "leave",
    duplicates: "keep",
  });
  const [isCleaning, setIsCleaning] = useState(false);

  const selectedFile = useMemo(() => 
    files.find(f => f.id === selectedFileId), 
    [files, selectedFileId]
  );

  const calculateMetadata = (data: any[]) => {
    if (!data.length) return null;
    const cols = Object.keys(data[0]);
    const columnStats: Record<string, any> = {};
    let totalNulls = 0;
    let totalBlanks = 0;

    cols.forEach(col => {
      let nulls = 0;
      let blanks = 0;
      const uniques = new Set();
      let isNumeric = true;

      data.forEach(row => {
        const val = row[col];
        if (val === null || val === undefined || val === "") {
          nulls++;
          totalNulls++;
        }
        if (typeof val === "string" && val.trim() === "") {
          blanks++;
          totalBlanks++;
        }
        if (val !== null && val !== undefined) {
          uniques.add(val);
          if (isNaN(Number(val)) && typeof val !== "number") {
            isNumeric = false;
          }
        }
      });

      columnStats[col] = {
        nulls,
        blanks,
        uniques: uniques.size,
        type: isNumeric ? "numeric" : "string"
      };
    });

    return {
      rows: data.length,
      cols: cols.length,
      nullCount: totalNulls,
      blankCount: totalBlanks,
      uniqueCount: new Set(data.map(r => JSON.stringify(r))).size,
      columnStats
    };
  };

  const handleFileUpload = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const uploadedFiles = e.target.files;
    if (!uploadedFiles) return;

    const newFiles: FileData[] = [];

    for (let i = 0; i < uploadedFiles.length; i++) {
      const file = uploadedFiles[i];
      const reader = new FileReader();

      const parsePromise = new Promise<any[]>((resolve, reject) => {
        const ext = file.name.split(".").pop()?.toLowerCase();
        
        if (ext === "csv") {
          Papa.parse(file, {
            header: true,
            dynamicTyping: true,
            complete: (results) => resolve(results.data),
            error: (err) => reject(err)
          });
        } else if (ext === "xlsx" || ext === "xls") {
          reader.onload = (e) => {
            const data = new Uint8Array(e.target?.result as ArrayBuffer);
            const workbook = XLSX.read(data, { type: "array" });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            resolve(XLSX.utils.sheet_to_json(firstSheet));
          };
          reader.readAsArrayBuffer(file);
        } else if (ext === "json") {
          reader.onload = (e) => {
            try {
              resolve(JSON.parse(e.target?.result as string));
            } catch (err) { reject(err); }
          };
          reader.readAsText(file);
        } else if (ext === "yaml" || ext === "yml") {
          reader.onload = (e) => {
            try {
              resolve(yaml.load(e.target?.result as string) as any[]);
            } catch (err) { reject(err); }
          };
          reader.readAsText(file);
        } else {
          reject(new Error("Unsupported file type"));
        }
      });

      try {
        const data = await parsePromise;
        const metadata = calculateMetadata(data);
        if (metadata) {
          const fileId = Math.random().toString(36).substring(7);
          newFiles.push({
            id: fileId,
            name: file.name,
            type: file.type || "application/octet-stream",
            data,
            originalData: [...data],
            metadata
          });
        }
      } catch (err) {
        console.error(`Error parsing ${file.name}:`, err);
      }
    }

    setFiles(prev => [...prev, ...newFiles]);
    if (newFiles.length > 0 && !selectedFileId) {
      setSelectedFileId(newFiles[0].id);
    }
  }, [selectedFileId]);

  const cleanData = () => {
    if (!selectedFile) return;
    setIsCleaning(true);

    setTimeout(() => {
      let cleaned = [...selectedFile.data];
      const initialRows = cleaned.length;
      const changes: string[] = [];

      // 1. Handle Duplicates
      if (config.duplicates === "remove") {
        const seen = new Set();
        const beforeCount = cleaned.length;
        cleaned = cleaned.filter(row => {
          const str = JSON.stringify(row);
          if (seen.has(str)) return false;
          seen.add(str);
          return true;
        });
        if (beforeCount !== cleaned.length) {
          changes.push(`Removed ${beforeCount - cleaned.length} duplicate rows`);
        }
      }

      // 2. Handle Blanks (Strings)
      if (config.blanks !== "leave") {
        let blankCount = 0;
        cleaned = cleaned.map(row => {
          const newRow = { ...row };
          Object.keys(newRow).forEach(key => {
            if (typeof newRow[key] === "string" && newRow[key].trim() === "") {
              if (config.blanks === "toNull") {
                newRow[key] = null;
                blankCount++;
              }
            }
          });
          return newRow;
        });

        if (config.blanks === "drop") {
          const before = cleaned.length;
          cleaned = cleaned.filter(row => {
            return !Object.values(row).some(v => typeof v === "string" && v.trim() === "");
          });
          if (before !== cleaned.length) {
            changes.push(`Dropped ${before - cleaned.length} rows with blank strings`);
          }
        } else if (blankCount > 0) {
          changes.push(`Converted ${blankCount} blank strings to NULL`);
        }
      }

      // 3. Handle Zeros
      if (config.zeros !== "leave") {
        const cols = Object.keys(cleaned[0] || {});
        cols.forEach(col => {
          const stats = selectedFile.metadata.columnStats[col];
          if (stats.type === "numeric") {
            let zeroCount = 0;
            const values = cleaned.map(r => Number(r[col])).filter(v => !isNaN(v));
            const mean = values.reduce((a, b) => a + b, 0) / values.length;
            const sorted = [...values].sort((a, b) => a - b);
            const median = sorted[Math.floor(sorted.length / 2)];

            cleaned = cleaned.map(row => {
              if (row[col] === 0) {
                zeroCount++;
                if (config.zeros === "toNull") return { ...row, [col]: null };
                if (config.zeros === "replaceMean") return { ...row, [col]: mean };
                if (config.zeros === "replaceMedian") return { ...row, [col]: median };
              }
              return row;
            });
            if (zeroCount > 0) {
              changes.push(`Handled ${zeroCount} zero values in column '${col}' using ${config.zeros}`);
            }
          }
        });
      }

      // 4. Handle NULLs
      if (config.nulls !== "leave") {
        if (config.nulls === "drop") {
          const before = cleaned.length;
          cleaned = cleaned.filter(row => {
            return !Object.values(row).some(v => v === null || v === undefined || v === "");
          });
          if (before !== cleaned.length) {
            changes.push(`Dropped ${before - cleaned.length} rows with NULL values`);
          }
        } else {
          const cols = Object.keys(cleaned[0] || {});
          cols.forEach(col => {
            const stats = selectedFile.metadata.columnStats[col];
            let fillCount = 0;
            
            cleaned = cleaned.map(row => {
              if (row[col] === null || row[col] === undefined || row[col] === "") {
                fillCount++;
                if (config.nulls === "fill0") return { ...row, [col]: 0 };
                if (config.nulls === "fillMean" && stats.type === "numeric") {
                  const values = selectedFile.data.map(r => Number(r[col])).filter(v => !isNaN(v));
                  const mean = values.reduce((a, b) => a + b, 0) / values.length;
                  return { ...row, [col]: mean };
                }
                if (config.nulls === "fillMode") {
                  const counts: Record<any, number> = {};
                  selectedFile.data.forEach(r => {
                    const v = r[col];
                    if (v !== null && v !== undefined && v !== "") {
                      counts[v] = (counts[v] || 0) + 1;
                    }
                  });
                  const mode = Object.keys(counts).reduce((a, b) => counts[a] > counts[b] ? a : b, "");
                  return { ...row, [col]: mode };
                }
              }
              return row;
            });
            if (fillCount > 0) {
              changes.push(`Filled ${fillCount} NULLs in column '${col}' using ${config.nulls}`);
            }
          });
        }
      }

      const newMetadata = calculateMetadata(cleaned);
      if (newMetadata) {
        setFiles(prev => prev.map(f => f.id === selectedFile.id ? { ...f, data: cleaned, metadata: newMetadata } : f));
        
        const run: RunHistory = {
          id: Math.random().toString(36).substring(7),
          timestamp: Date.now(),
          fileName: selectedFile.name,
          config: { ...config },
          rowsRemoved: initialRows - cleaned.length,
          changes
        };
        setHistory(prev => [run, ...prev]);
      }
      setIsCleaning(false);
    }, 800);
  };

  const resetFile = () => {
    if (!selectedFile) return;
    const originalMetadata = calculateMetadata(selectedFile.originalData);
    if (originalMetadata) {
      setFiles(prev => prev.map(f => f.id === selectedFile.id ? { 
        ...f, 
        data: [...selectedFile.originalData], 
        metadata: originalMetadata 
      } : f));
      
      const run: RunHistory = {
        id: Math.random().toString(36).substring(7),
        timestamp: Date.now(),
        fileName: selectedFile.name,
        config: config,
        rowsRemoved: 0,
        changes: ["Reset to original data"]
      };
      setHistory(prev => [run, ...prev]);
    }
  };

  const downloadCSV = () => {
    if (!selectedFile) return;
    const csv = Papa.unparse(selectedFile.data);
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);
    link.setAttribute("href", url);
    link.setAttribute("download", `cleaned_${selectedFile.name.split(".")[0]}.csv`);
    link.style.visibility = "hidden";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  return (
    <div className="min-h-screen bg-[#fafafa] text-slate-900 font-sans selection:bg-indigo-100">
      {/* Header */}
      <header className="sticky top-0 z-50 bg-white/80 backdrop-blur-md border-b border-slate-200">
        <div className="max-w-7xl mx-auto px-4 h-16 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="w-9 h-9 bg-indigo-600 rounded-xl flex items-center justify-center shadow-lg shadow-indigo-200">
              <RefreshCw className="w-5 h-5 text-white" />
            </div>
            <h1 className="text-xl font-bold tracking-tight text-slate-800">CleanSheet</h1>
          </div>
          
          <div className="flex items-center gap-4">
            <label className="cursor-pointer bg-slate-900 hover:bg-slate-800 text-white px-4 py-2 rounded-lg text-sm font-medium transition-all flex items-center gap-2 shadow-sm">
              <Upload className="w-4 h-4" />
              Upload Files/Folder
              <input 
                type="file" 
                multiple 
                {...({ webkitdirectory: "", directory: "" } as any)}
                className="hidden" 
                onChange={handleFileUpload}
                accept=".csv,.xlsx,.xls,.json,.yaml,.yml"
              />
            </label>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-8">
        {!files.length ? (
          <div className="flex flex-col items-center justify-center py-24 border-2 border-dashed border-slate-200 rounded-3xl bg-white/50">
            <div className="w-20 h-20 bg-indigo-50 rounded-full flex items-center justify-center mb-6">
              <FileSpreadsheet className="w-10 h-10 text-indigo-500" />
            </div>
            <h2 className="text-2xl font-semibold text-slate-800 mb-2">No files uploaded yet</h2>
            <p className="text-slate-500 mb-8 max-w-md text-center">
              Upload CSV, Excel, JSON or YAML files to start cleaning your data. 
              You can select multiple files or drag them here.
            </p>
            <label className="cursor-pointer bg-white border border-slate-200 hover:border-indigo-300 hover:bg-indigo-50/50 px-8 py-3 rounded-2xl text-slate-700 font-medium transition-all flex items-center gap-3 shadow-sm group">
              <Upload className="w-5 h-5 text-slate-400 group-hover:text-indigo-500" />
              Choose Files or Folder
              <input 
                type="file" 
                multiple 
                {...({ webkitdirectory: "", directory: "" } as any)}
                className="hidden" 
                onChange={handleFileUpload}
                accept=".csv,.xlsx,.xls,.json,.yaml,.yml"
              />
            </label>
          </div>
        ) : (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
            {/* Sidebar: File List & Config */}
            <div className="lg:col-span-4 space-y-6">
              {/* File Selector */}
              <section className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
                <h3 className="text-sm font-semibold text-slate-400 uppercase tracking-wider mb-4 flex items-center gap-2">
                  <Database className="w-4 h-4" />
                  Your Datasets
                </h3>
                <div className="space-y-2">
                  {files.map(file => (
                    <button
                      key={file.id}
                      onClick={() => setSelectedFileId(file.id)}
                      className={cn(
                        "w-full flex items-center justify-between p-3 rounded-xl transition-all group",
                        selectedFileId === file.id 
                          ? "bg-indigo-50 text-indigo-700 ring-1 ring-indigo-200" 
                          : "hover:bg-slate-50 text-slate-600"
                      )}
                    >
                      <div className="flex items-center gap-3 overflow-hidden">
                        <FileText className={cn("w-5 h-5 shrink-0", selectedFileId === file.id ? "text-indigo-500" : "text-slate-400")} />
                        <span className="text-sm font-medium truncate">{file.name}</span>
                      </div>
                      <button 
                        onClick={(e) => {
                          e.stopPropagation();
                          setFiles(prev => prev.filter(f => f.id !== file.id));
                          if (selectedFileId === file.id) setSelectedFileId(null);
                        }}
                        className="opacity-0 group-hover:opacity-100 p-1 hover:bg-red-50 hover:text-red-500 rounded-md transition-all"
                      >
                        <X className="w-4 h-4" />
                      </button>
                    </button>
                  ))}
                </div>
              </section>

              {/* Cleaning Options */}
              <section className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
                <h3 className="text-sm font-semibold text-slate-400 uppercase tracking-wider mb-6 flex items-center gap-2">
                  <Settings2 className="w-4 h-4" />
                  Cleaning Rules
                </h3>
                
                <div className="space-y-5">
                  <div className="space-y-2">
                    <label className="text-xs font-bold text-slate-500 uppercase tracking-wide">NULL Values</label>
                    <select 
                      value={config.nulls}
                      onChange={(e) => setConfig(prev => ({ ...prev, nulls: e.target.value as NullOption }))}
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2.5 text-sm focus:ring-2 focus:ring-indigo-500 outline-none transition-all"
                    >
                      <option value="leave">Leave unchanged</option>
                      <option value="drop">Drop rows with NULLs</option>
                      <option value="fill0">Fill with 0</option>
                      <option value="fillMean">Fill with mean (Numeric)</option>
                      <option value="fillMode">Fill with mode</option>
                    </select>
                  </div>

                  <div className="space-y-2">
                    <label className="text-xs font-bold text-slate-500 uppercase tracking-wide">Blank Strings</label>
                    <select 
                      value={config.blanks}
                      onChange={(e) => setConfig(prev => ({ ...prev, blanks: e.target.value as BlankOption }))}
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2.5 text-sm focus:ring-2 focus:ring-indigo-500 outline-none transition-all"
                    >
                      <option value="leave">Leave unchanged</option>
                      <option value="toNull">Convert to NULL</option>
                      <option value="drop">Drop rows with blanks</option>
                    </select>
                  </div>

                  <div className="space-y-2">
                    <label className="text-xs font-bold text-slate-500 uppercase tracking-wide">Zero Values</label>
                    <select 
                      value={config.zeros}
                      onChange={(e) => setConfig(prev => ({ ...prev, zeros: e.target.value as ZeroOption }))}
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2.5 text-sm focus:ring-2 focus:ring-indigo-500 outline-none transition-all"
                    >
                      <option value="leave">Leave unchanged</option>
                      <option value="toNull">Convert to NULL</option>
                      <option value="replaceMean">Replace with mean</option>
                      <option value="replaceMedian">Replace with median</option>
                    </select>
                  </div>

                  <div className="space-y-2">
                    <label className="text-xs font-bold text-slate-500 uppercase tracking-wide">Duplicates</label>
                    <select 
                      value={config.duplicates}
                      onChange={(e) => setConfig(prev => ({ ...prev, duplicates: e.target.value as DuplicateOption }))}
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2.5 text-sm focus:ring-2 focus:ring-indigo-500 outline-none transition-all"
                    >
                      <option value="keep">Keep duplicates</option>
                      <option value="remove">Remove duplicates</option>
                    </select>
                  </div>

                  <button
                    onClick={cleanData}
                    disabled={isCleaning || !selectedFile}
                    className={cn(
                      "w-full py-3.5 rounded-xl font-bold text-sm transition-all flex items-center justify-center gap-2 shadow-lg",
                      isCleaning 
                        ? "bg-slate-100 text-slate-400 cursor-not-allowed" 
                        : "bg-indigo-600 hover:bg-indigo-700 text-white shadow-indigo-200 active:scale-[0.98]"
                    )}
                  >
                    {isCleaning ? (
                      <>
                        <RefreshCw className="w-4 h-4 animate-spin" />
                        Cleaning...
                      </>
                    ) : (
                      <>
                        <CheckCircle2 className="w-4 h-4" />
                        Clean Data
                      </>
                    )}
                  </button>
                </div>
              </section>

              {/* History */}
              <section className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
                <h3 className="text-sm font-semibold text-slate-400 uppercase tracking-wider mb-4 flex items-center gap-2">
                  <History className="w-4 h-4" />
                  Run History
                </h3>
                <div className="space-y-3 max-h-[300px] overflow-y-auto pr-2 custom-scrollbar">
                  {history.length === 0 ? (
                    <p className="text-xs text-slate-400 text-center py-4 italic">No cleaning runs yet</p>
                  ) : (
                    history.map(run => (
                      <div key={run.id} className="p-3 bg-slate-50 rounded-xl border border-slate-100 space-y-2">
                        <div className="flex items-center justify-between">
                          <span className="text-[10px] font-bold text-slate-400 uppercase">{new Date(run.timestamp).toLocaleTimeString()}</span>
                          <span className="text-[10px] font-bold bg-red-50 text-red-600 px-1.5 py-0.5 rounded">-{run.rowsRemoved} rows</span>
                        </div>
                        <p className="text-xs font-semibold text-slate-700 truncate">{run.fileName}</p>
                        <ul className="text-[10px] text-slate-500 space-y-1 pl-2 border-l border-slate-200">
                          {run.changes.slice(0, 2).map((c, idx) => (
                            <li key={idx} className="truncate">• {c}</li>
                          ))}
                          {run.changes.length > 2 && <li className="italic">+ {run.changes.length - 2} more</li>}
                        </ul>
                      </div>
                    ))
                  )}
                </div>
              </section>
            </div>

            {/* Main Content: Preview & Stats */}
            <div className="lg:col-span-8 space-y-6">
              {selectedFile ? (
                <>
                  {/* Stats Cards */}
                  <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                    <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm">
                      <p className="text-[10px] font-bold text-slate-400 uppercase mb-1">Total Rows</p>
                      <p className="text-2xl font-bold text-slate-800">{selectedFile.metadata.rows.toLocaleString()}</p>
                    </div>
                    <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm">
                      <p className="text-[10px] font-bold text-slate-400 uppercase mb-1">Columns</p>
                      <p className="text-2xl font-bold text-slate-800">{selectedFile.metadata.cols}</p>
                    </div>
                    <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm">
                      <p className="text-[10px] font-bold text-slate-400 uppercase mb-1">Null Values</p>
                      <p className="text-2xl font-bold text-indigo-600">
                        {((selectedFile.metadata.nullCount / (selectedFile.metadata.rows * selectedFile.metadata.cols)) * 100).toFixed(1)}%
                      </p>
                    </div>
                    <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm">
                      <p className="text-[10px] font-bold text-slate-400 uppercase mb-1">Unique Rows</p>
                      <p className="text-2xl font-bold text-slate-800">
                        {((selectedFile.metadata.uniqueCount / selectedFile.metadata.rows) * 100).toFixed(0)}%
                      </p>
                    </div>
                  </div>

                  {/* Data Preview */}
                  <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden flex flex-col">
                    <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
                      <div className="flex items-center gap-3">
                        <h2 className="font-bold text-slate-800">Data Preview</h2>
                        <span className="text-[10px] bg-slate-200 text-slate-600 px-2 py-0.5 rounded-full font-bold uppercase">Showing first 100 rows</span>
                      </div>
                      <div className="flex items-center gap-2">
                        <button 
                          onClick={resetFile}
                          className="text-slate-400 hover:text-red-500 text-sm font-bold flex items-center gap-1.5 px-3 py-1.5 rounded-lg hover:bg-red-50 transition-all"
                        >
                          <RefreshCw className="w-4 h-4" />
                          Reset
                        </button>
                        <button 
                          onClick={downloadCSV}
                          className="text-indigo-600 hover:text-indigo-700 text-sm font-bold flex items-center gap-1.5 px-3 py-1.5 rounded-lg hover:bg-indigo-50 transition-all"
                        >
                          <Download className="w-4 h-4" />
                          Download CSV
                        </button>
                      </div>
                    </div>
                    
                    <div className="overflow-x-auto overflow-y-auto max-h-[600px] custom-scrollbar">
                      <table className="w-full text-left border-collapse">
                        <thead className="sticky top-0 z-10 bg-white shadow-sm">
                          <tr>
                            {Object.keys(selectedFile.data[0] || {}).map(col => (
                              <th key={col} className="px-6 py-4 text-xs font-bold text-slate-400 uppercase tracking-wider border-b border-slate-100 min-w-[150px]">
                                <div className="flex flex-col gap-1">
                                  <span>{col}</span>
                                  <div className="flex items-center gap-2">
                                    <span className={cn(
                                      "text-[8px] px-1.5 py-0.5 rounded-sm",
                                      selectedFile.metadata.columnStats[col].type === "numeric" ? "bg-blue-50 text-blue-600" : "bg-amber-50 text-amber-600"
                                    )}>
                                      {selectedFile.metadata.columnStats[col].type}
                                    </span>
                                    <span className="text-[8px] text-slate-300 font-normal">
                                      {selectedFile.metadata.columnStats[col].nulls} nulls
                                    </span>
                                  </div>
                                </div>
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-50">
                          {selectedFile.data.slice(0, 100).map((row, idx) => (
                            <tr key={idx} className="hover:bg-slate-50/50 transition-colors">
                              {Object.values(row).map((val: any, vIdx) => (
                                <td key={vIdx} className="px-6 py-3.5 text-sm text-slate-600 whitespace-nowrap">
                                  {val === null || val === undefined || val === "" ? (
                                    <span className="text-slate-300 italic text-xs">null</span>
                                  ) : (
                                    String(val)
                                  )}
                                </td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </>
              ) : (
                <div className="h-full flex items-center justify-center bg-white rounded-2xl border border-slate-200 border-dashed p-12">
                  <div className="text-center">
                    <AlertCircle className="w-12 h-12 text-slate-200 mx-auto mb-4" />
                    <p className="text-slate-400 font-medium">Select a file to view details</p>
                  </div>
                </div>
              )}
            </div>
          </div>
        )}
      </main>

      <style dangerouslySetInnerHTML={{ __html: `
        .custom-scrollbar::-webkit-scrollbar {
          width: 6px;
          height: 6px;
        }
        .custom-scrollbar::-webkit-scrollbar-track {
          background: transparent;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb {
          background: #e2e8f0;
          border-radius: 10px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover {
          background: #cbd5e1;
        }
      `}} />
    </div>
  );
}
