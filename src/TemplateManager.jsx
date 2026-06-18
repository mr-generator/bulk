import React, { useState, useRef } from 'react';
import { 
  FileSpreadsheet, Upload, Clipboard, CheckCircle, 
  XCircle, AlertTriangle, Download, Layers, ShieldCheck, FileArchive, RefreshCw, Trash2
} from 'lucide-react';
import * as XLSX from 'xlsx';

// ==========================================
// 1. DYNAMIC TEMPLATE CONFIGURATION PROFILES
// ==========================================
const TEMPLATE_CONFIGS = {
  VENDOR: {
    templateId: "VENDOR",
    templateName: "Vendor Payment Template",
    description: "Official configuration mapped to Vendor global template.xlsx. Enforces text casting on identifiers, banking keys, and handles Amount cleanly as a numeric field.",
    columns: [
      "Receiving Party Code", "Receiving Party Name", "Transaction Code", 
      "Component Code/FMR Code", "Payment Mode", "Unique Transaction ID", 
      "Expense Type", "Amount", "Account Number", "Remarks", "Action Type"
    ],
    mandatory: ["Receiving Party Code", "Receiving Party Name", "Unique Transaction ID", "Amount"],
    dates: [], // Add date columns here if required by future iterations
    amounts: ["Amount"] // Strict mapping: ONLY "Amount" is parsed as a true numeric value
  },
  BENEFICIARY: {
    templateId: "BENEFICIARY",
    templateName: "Beneficiary Payment Template",
    description: "Configured for direct benefit transfers. Enforces exact structural indexing with float numeric parsing on Centre Share amounts.",
    columns: [
      "CPSMS Beneficiary Code", "Beneficiary Name", "Purpose", 
      "Centre Share Payment Amount", "Payment From Date", "Payment To Date", 
      "Transaction Id", "Identifier", "Payment Mode"
    ],
    mandatory: ["CPSMS Beneficiary Code", "Beneficiary Name", "Centre Share Payment Amount"],
    dates: ["Payment From Date", "Payment To Date"],
    amounts: ["Centre Share Payment Amount"]
  }
};

export default function TemplateManager() {
  const [selectedTemplateKey, setSelectedTemplateKey] = useState('');
  const [pasteData, setPasteData] = useState('');
  const [dragActive, setDragActive] = useState(false);
  const [fileName, setFileName] = useState('');
  const [activeStep, setActiveStep] = useState(1);
  const fileInputRef = useRef(null);
  
  const [validationSummary, setValidationSummary] = useState(null);
  const [processedData, setProcessedData] = useState([]);
  const [isProcessed, setIsProcessed] = useState(false);

  const currentTemplate = TEMPLATE_CONFIGS[selectedTemplateKey];

  // ==========================================
  // 2. DATA TRANSLATION UTILITIES
  // ==========================================
  const normalizeDate = (dateVal) => {
    if (!dateVal) return '';
    const dateStr = String(dateVal).trim();

    const strictDdmmyyyyRegex = /^(\d{2})[-/](\d{2})[-/](\d{4})$/;
    if (strictDdmmyyyyRegex.test(dateStr)) {
      return dateStr.replace(/\//g, '-');
    }

    if (typeof dateVal === 'number') {
      const parsedDate = XLSX.SSF.parse_date_code(dateVal);
      const day = String(parsedDate.d).padStart(2, '0');
      const month = String(parsedDate.m).padStart(2, '0');
      return `${day}-${month}-${parsedDate.y}`;
    }

    let d = new Date(dateVal);
    if (isNaN(d.getTime())) return dateStr;
    
    return `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}`;
  };

  const normalizeTextAndAmounts = (val) => {
    if (val === undefined || val === null) return '';
    let strValue = String(val).trim();
    if (strValue.toLowerCase().includes('e') && !isNaN(Number(strValue))) {
      strValue = Number(strValue).toLocaleString('fullwide', { useGrouping: false });
    }
    return strValue; 
  };

  // ==========================================
  // 3. VALIDATION PIPELINE EXECUTION ENGINE
  // ==========================================
  const executeDataPipeline = (headers, rows) => {
    if (!currentTemplate) return;

    const cleanHeaders = headers.map(h => String(h).trim());
    const missingColumns = currentTemplate.columns.filter(c => !cleanHeaders.includes(c));
    const extraColumns = cleanHeaders.filter(c => !currentTemplate.columns.includes(c));
    const duplicates = cleanHeaders.filter((item, index) => cleanHeaders.indexOf(item) !== index);
    
    let isOrderCorrect = true;
    currentTemplate.columns.forEach((col, idx) => {
      if (cleanHeaders[idx] !== col) isOrderCorrect = false;
    });

    if (missingColumns.length > 0 || extraColumns.length > 0 || duplicates.length > 0) {
      setValidationSummary({
        isValid: false,
        criticalHeaderError: true,
        missingColumns,
        extraColumns,
        duplicates: [...new Set(duplicates)],
        orderMismatch: !isOrderCorrect,
        rowErrors: []
      });
      setIsProcessed(false);
      setActiveStep(3);
      return;
    }

    let cleanedRows = [];
    let rowErrors = [];

    rows.forEach((row, rowIndex) => {
      let cleanedRow = {};
      let rowHasError = false;
      let targetRowMessages = [];

      currentTemplate.columns.forEach(col => {
        let rawValue = row[col];
        
        if (currentTemplate.mandatory.includes(col) && (rawValue === undefined || rawValue === null || String(rawValue).trim() === "")) {
          rowHasError = true;
          targetRowMessages.push(`Missing mandatory field: [${col}]`);
        }

        if (currentTemplate.dates.includes(col)) {
          cleanedRow[col] = normalizeDate(rawValue);
        } else if (currentTemplate.amounts.includes(col)) {
          // Dynamic conversion back to proper numeric layout properties
          const cleanAmount = String(rawValue || "").replace(/[^0-9.]/g, ''); 
          cleanedRow[col] = cleanAmount !== "" ? parseFloat(cleanAmount) : 0.00;
        } else {
          // Explicit text lock on IDs, Codes, and Account Numbers to retain formatting
          cleanedRow[col] = normalizeTextAndAmounts(rawValue);
        }
      });

      if (rowHasError) {
        rowErrors.push({ rowNumber: rowIndex + 1, messages: targetRowMessages });
      }
      cleanedRows.push(cleanedRow);
    });

    setValidationSummary({
      isValid: rowErrors.length === 0,
      criticalHeaderError: false,
      missingColumns: [],
      extraColumns: [],
      duplicates: [],
      orderMismatch: !isOrderCorrect,
      rowErrors
    });

    setProcessedData(cleanedRows);
    setIsProcessed(true);
    setActiveStep(3);
  };

  // ==========================================
  // 4. STORAGE EXTRACTION READERS & WRITERS
  // ==========================================
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setFileName(file.name);

    const reader = new FileReader();
    reader.onload = (evt) => {
      const binaryData = evt.target.result;
      const workbook = XLSX.read(binaryData, { type: 'binary', cellDates: false });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const rawJson = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
      
      if (rawJson.length === 0) return;
      const headers = Object.keys(rawJson[0]);
      executeDataPipeline(headers, rawJson);
    };
    reader.readAsBinaryString(file);
  };

  const handleClipboardPaste = (e) => {
    const content = e.target.value;
    setPasteData(content);
    if (!content.trim()) return;

    const lines = content.split('\n').map(line => line.split('\t').map(c => c.replace(/\r/g, "").trim()));
    const headers = lines[0];
    const dataLines = lines.slice(1).filter(l => l.length > 0 && l.join('') !== '');

    const rows = dataLines.map(line => {
      let obj = {};
      headers.forEach((h, idx) => { obj[h] = line[idx] || ''; });
      return obj;
    });

    setFileName("Clipboard Ingestion Stream");
    executeDataPipeline(headers, rows);
  };

  const saveAsExcelFile = (dataArray, outputName) => {
    const worksheet = XLSX.utils.json_to_sheet(dataArray, { skipHeader: false });
    const workbook = XLSX.utils.book_new();
    
    Object.keys(worksheet).forEach((cellRef) => {
      if (cellRef[0] === '!') return; 
      const cell = worksheet[cellRef];
      if (cell && cell.v !== undefined) {
        if (typeof cell.v === 'number') {
          cell.t = 'n';    // Forces True Excel Numeric Format
          cell.z = '0.00'; // Standard numeric precision masking layout
        } else {
          cell.v = String(cell.v).trim();
          cell.t = 's';    // Locks column values as strict explicit Text fields
          cell.z = '@';
        }
      }
    });

    XLSX.utils.book_append_sheet(workbook, worksheet, "Processed Data");
    XLSX.writeFile(workbook, outputName, { bookType: 'xlsx', type: 'binary', cellStyles: true });
  };

  const executeDownloadAction = async (type) => {
    if (type === 'cleaned') {
      saveAsExcelFile(processedData, `${currentTemplate.templateId}_Template_Cleaned.xlsx`);
    } else if (type === 'report') {
      const blob = new Blob([JSON.stringify(validationSummary, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `${currentTemplate.templateId}_Error_Report.json`;
      link.click();
    } else if (type === 'zip') {
      // Package multi-split Excel blocks into a unified safe ZIP package bundle dynamically
      const JSZip = (await import('jszip')).default;
      const zip = new JSZip();

      for (let i = 0; i < processedData.length; i += 300) {
        const chunk = processedData.slice(i, i + 300);
        const filePartNumber = Math.floor(i / 300) + 1;
        
        const worksheet = XLSX.utils.json_to_sheet(chunk, { skipHeader: false });
        const workbook = XLSX.utils.book_new();

        Object.keys(worksheet).forEach((cellRef) => {
          if (cellRef[0] === '!') return;
          const cell = worksheet[cellRef];
          if (cell && cell.v !== undefined) {
            if (typeof cell.v === 'number') {
              cell.t = 'n';
              cell.z = '0.00';
            } else {
              cell.v = String(cell.v).trim();
              cell.t = 's';
              cell.z = '@';
            }
          }
        });

        XLSX.utils.book_append_sheet(workbook, worksheet, "Processed Data");
        const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array', cellStyles: true });
        zip.file(`${currentTemplate.templateId}_Part_${filePartNumber}.xlsx`, excelBuffer);
      }

      zip.generateAsync({ type: "blob" }).then((content) => {
        const url = URL.createObjectURL(content);
        const link = document.createElement('a');
        link.href = url;
        link.download = `${currentTemplate.templateId}_All_Batches.zip`;
        link.click();
      });
    }
  };

  const downloadSampleTemplate = () => {
    if (!currentTemplate) return;
    const emptyRowStructure = {};
    currentTemplate.columns.forEach(col => { emptyRowStructure[col] = ''; });
    saveAsExcelFile([emptyRowStructure], `${currentTemplate.templateId}_Global_Template.xlsx`);
  };

  const resetPipeline = () => {
    setValidationSummary(null);
    setProcessedData([]);
    setIsProcessed(false);
    setPasteData('');
    setFileName('');
    setActiveStep(1);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const totalChunks = Math.ceil(processedData.length / 300) || 1;

  // ==========================================
  // 5. TAILWIND DESIGN LAYOUT
  // ==========================================
  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 p-4 md:p-8 font-sans antialiased selection:bg-emerald-500 selection:text-slate-950">
      <div className="max-w-5xl mx-auto space-y-6">
        
        {/* Navigation Banner */}
        <header className="bg-slate-900 border border-slate-800 rounded-2xl p-6 flex flex-col md:flex-row justify-between items-start md:items-center gap-4 shadow-xl">
          <div className="flex items-center gap-4">
            <div className="p-3 bg-gradient-to-br from-emerald-400 to-teal-600 rounded-xl text-slate-950 shadow-lg">
              <FileSpreadsheet size={26} />
            </div>
            <div>
              <h1 className="text-xl md:text-2xl font-black tracking-tight text-white">
                Multi-Template Harmonization Hub
              </h1>
              <p className="text-slate-400 text-xs md:text-sm mt-0.5">
                Cleansing architecture targeting global master sheet structures.
              </p>
            </div>
          </div>
          
          <div className="flex items-center gap-3 w-full md:w-auto justify-between md:justify-end border-t md:border-t-0 border-slate-800/60 pt-3 md:pt-0">
            <span className="text-[11px] font-mono tracking-wider uppercase bg-slate-950 px-3 py-1.5 rounded-full border border-slate-800 text-slate-400 flex items-center gap-2">
              <span className="h-2 w-2 rounded-full bg-emerald-400 animate-ping" /> Engine Loaded
            </span>
            {isProcessed && (
              <button onClick={resetPipeline} className="bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-semibold py-1.5 px-3 rounded-lg flex items-center gap-1.5 transition border border-slate-700">
                <RefreshCw size={13} /> Reset
              </button>
            )}
          </div>
        </header>

        {/* Step Progression Flow */}
        <div className="grid grid-cols-3 gap-2 bg-slate-900/40 p-1.5 rounded-xl border border-slate-800/60 text-center">
          <div className={`py-2 px-1 rounded-lg text-xs font-bold transition ${activeStep === 1 ? 'bg-emerald-500/10 text-emerald-400 border border-emerald-500/20' : 'text-slate-500'}`}>
            1. Template Config Match
          </div>
          <div className={`py-2 px-1 rounded-lg text-xs font-bold transition ${activeStep === 2 ? 'bg-emerald-500/10 text-emerald-400 border border-emerald-500/20' : 'text-slate-500'}`}>
            2. Ingestion Stream
          </div>
          <div className={`py-2 px-1 rounded-lg text-xs font-bold transition ${activeStep === 3 ? 'bg-emerald-500/10 text-emerald-400 border border-emerald-500/20' : 'text-slate-500'}`}>
            3. Verification Export
          </div>
        </div>

        {/* STEP 1: Select Profile Module */}
        <section className="bg-slate-900 border border-slate-800 rounded-2xl p-6 shadow-xl relative overflow-hidden">
          <div className="flex flex-col lg:flex-row gap-6">
            <div className="w-full lg:w-1/3 space-y-2">
              <label className="text-xs font-extrabold uppercase tracking-widest text-slate-400 flex items-center gap-1.5">
                <Layers size={13} className="text-emerald-400" /> Target Profile Format
              </label>
              <select 
                value={selectedTemplateKey} 
                onChange={(e) => {
                  setSelectedTemplateKey(e.target.value);
                  setValidationSummary(null);
                  setIsProcessed(false);
                  setPasteData('');
                  setFileName('');
                  setActiveStep(e.target.value ? 2 : 1);
                }}
                className="w-full bg-slate-950 border border-slate-800 hover:border-slate-700 text-slate-200 rounded-xl px-4 py-3 text-sm font-semibold focus:outline-none focus:ring-2 focus:ring-emerald-500"
              >
                <option value="">-- Choose Profile Mapping Type --</option>
                {Object.keys(TEMPLATE_CONFIGS).map(key => (
                  <option key={key} value={key}>{TEMPLATE_CONFIGS[key].templateName}</option>
                ))}
              </select>
            </div>

            {currentTemplate ? (
              <div className="flex-1 bg-slate-950/60 rounded-xl p-5 border border-slate-800 flex flex-col md:flex-row gap-6 animate-fadeIn">
                <div className="flex-1 space-y-4">
                  <div>
                    <span className="bg-emerald-500/10 text-emerald-400 text-[10px] font-bold px-2 py-0.5 rounded border border-emerald-500/20 uppercase font-mono">Verified Format Blueprint</span>
                    <h3 className="text-base font-bold text-white mt-1">{currentTemplate.templateName}</h3>
                    <p className="text-xs text-slate-400 mt-1 leading-relaxed">{currentTemplate.description}</p>
                  </div>
                  <div>
                    <span className="text-slate-500 block text-[10px] font-bold uppercase tracking-wider mb-2">Required Heading Layout Sequence:</span>
                    <div className="flex flex-wrap gap-1.5 max-h-[110px] overflow-y-auto p-1 bg-slate-900/60 rounded-lg border border-slate-800/40">
                      {currentTemplate.columns.map((c, i) => (
                        <span key={i} className="text-[11px] font-mono bg-slate-950 text-slate-300 px-2.5 py-1 rounded border border-slate-800/80 shadow-sm flex items-center gap-1">
                          <span className="text-slate-600 font-bold">{i+1}.</span> {c}
                          {currentTemplate.mandatory.includes(c) && <span className="text-amber-500 font-bold">*</span>}
                        </span>
                      ))}
                    </div>
                  </div>
                </div>

                <div className="w-full md:w-52 flex flex-col justify-between bg-slate-900/60 p-4 rounded-xl border border-slate-800 space-y-4">
                  <div className="space-y-2 text-xs text-slate-400">
                    <div className="flex justify-between border-b border-slate-800 pb-1.5"><span>Columns:</span><span className="font-mono text-white font-bold">{currentTemplate.columns.length}</span></div>
                    <div className="flex justify-between border-b border-slate-800 pb-1.5"><span>Mandatory:</span><span className="font-mono text-amber-400 font-bold">{currentTemplate.mandatory.length}</span></div>
                    <div className="flex justify-between"><span>Amount Rules:</span><span className="font-mono text-cyan-400 font-bold">Numeric Float</span></div>
                  </div>
                  <button 
                    onClick={downloadSampleTemplate}
                    className="w-full bg-slate-800 hover:bg-slate-700 text-white text-xs font-bold py-2.5 px-3 rounded-lg flex items-center justify-center gap-2 transition border border-slate-700 shadow"
                  >
                    <Download size={14} className="text-emerald-400" /> Download Structure
                  </button>
                </div>
              </div>
            ) : (
              <div className="flex-1 border border-dashed border-slate-800 rounded-xl flex items-center justify-center p-6 text-center text-slate-500 text-xs">
                Awaiting structural profile selection to deploy extraction criteria matrix configurations...
              </div>
            )}
          </div>
        </section>

        {currentTemplate && (
          <>
            {/* STEP 2: Dual Option Data Ingestion */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              
              {/* Option A - Spreadsheet Drag Drops */}
              <div 
                onDragOver={(e) => { e.preventDefault(); setDragActive(true); }}
                onDragLeave={() => setDragActive(false)}
                onDrop={(e) => {
                  e.preventDefault();
                  setDragActive(false);
                  if (e.dataTransfer.files?.[0]) handleFileUpload({ target: { files: e.dataTransfer.files } });
                }}
                className={`bg-slate-900/60 border-2 border-dashed rounded-2xl p-6 flex flex-col items-center justify-center text-center transition ${
                  dragActive ? 'border-emerald-500 bg-slate-900/90 scale-[1.01]' : 'border-slate-800 hover:border-slate-700'
                }`}
              >
                <div className="p-3 bg-slate-950 rounded-xl text-emerald-400 border border-slate-800 mb-3 shadow-inner">
                  <Upload size={24} />
                </div>
                <h3 className="font-bold text-sm text-slate-100">Upload Native Excel Workbook</h3>
                <p className="text-xs text-slate-400 max-w-xs mt-1 mb-4 leading-relaxed">
                  Drop your template file target here. Accepted formats: <code className="font-mono text-slate-300">.xlsx</code> or <code className="font-mono text-slate-300">.xls</code>.
                </p>
                <button 
                  onClick={() => fileInputRef.current?.click()}
                  className="bg-emerald-600 hover:bg-emerald-500 text-slate-950 text-xs font-extrabold px-4 py-2.5 rounded-lg shadow-lg transition"
                >
                  Browse Local Files
                </button>
                <input type="file" ref={fileInputRef} accept=".xlsx, .xls" className="hidden" onChange={handleFileUpload} />
              </div>

              {/* Option B - Clipboard Area */}
              <div className="bg-slate-900/60 rounded-2xl border border-slate-800 p-6 flex flex-col relative">
                <div className="flex items-center justify-between mb-2">
                  <div className="flex items-center gap-2 text-slate-300 font-bold text-sm">
                    <Clipboard size={16} className="text-cyan-400" />
                    <h3>Inject Clipboard Selections</h3>
                  </div>
                  {pasteData && (
                    <button onClick={() => setPasteData('')} className="text-slate-500 hover:text-slate-400 text-xs flex items-center gap-1 transition">
                      <Trash2 size={12} /> Clear Box
                    </button>
                  )}
                </div>
                <p className="text-xs text-slate-400 mb-3 leading-relaxed">Copy multiple spreadsheet block rows directly out of Excel and paste them down below.</p>
                <textarea
                  value={pasteData}
                  onChange={handleClipboardPaste}
                  placeholder="Paste grid blocks cleanly here (include table row column headers in line #1)..."
                  className="w-full flex-1 min-h-[120px] bg-slate-950 text-slate-300 p-3 rounded-xl border border-slate-800 focus:border-emerald-500/40 focus:outline-none focus:ring-1 focus:ring-emerald-500/30 font-mono text-xs resize-none transition"
                />
              </div>

            </div>

            {/* STEP 3: Automated Validation Summary Board */}
            {validationSummary && (
              <section className="bg-slate-900 border border-slate-800 rounded-2xl p-6 space-y-4 shadow-xl">
                
                <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 border-b border-slate-800 pb-4">
                  <div className="flex items-center gap-3">
                    {validationSummary.isValid && !validationSummary.criticalHeaderError ? (
                      <span className="flex items-center gap-1.5 text-emerald-400 font-extrabold bg-emerald-500/10 border border-emerald-500/20 px-3 py-1 rounded-lg text-xs tracking-wider uppercase font-mono">
                        <CheckCircle size={14} /> Dataset Verified Clean
                      </span>
                    ) : (
                      <span className="flex items-center gap-1.5 text-rose-400 font-extrabold bg-rose-500/10 border border-rose-500/20 px-3 py-1 rounded-lg text-xs tracking-wider uppercase font-mono">
                        <XCircle size={14} /> Parsing Exceptions Found
                      </span>
                    )}
                    <div className="text-xs text-slate-400 font-mono truncate max-w-xs">
                      Active Buffer: <span className="text-slate-200 underline">{fileName || "Clipboard Ingestion"}</span>
                    </div>
                  </div>
                  {validationSummary.orderMismatch && (
                    <div className="text-[11px] text-amber-400 bg-amber-500/10 border border-amber-500/20 px-2.5 py-1 rounded-md flex items-center gap-1.5 font-semibold">
                      <AlertTriangle size={13} /> Heading Matrix Sequence Mismatch Warning
                    </div>
                  )}
                </div>

                {/* Critical Column Array Structure Incompatibilities */}
                {validationSummary.criticalHeaderError && (
                  <div className="bg-rose-950/15 border border-rose-900/50 rounded-xl p-4 space-y-2 text-xs">
                    <div className="flex items-center gap-2 text-rose-400 font-bold text-sm">
                      <AlertTriangle size={15} />
                      <h4>Structural Layout Schema Alignment Error</h4>
                    </div>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4 pt-1">
                      {validationSummary.missingColumns.length > 0 && (
                        <div className="bg-slate-950 p-3 rounded-xl border border-slate-900">
                          <span className="text-rose-400 block font-bold text-[10px] uppercase tracking-wider mb-1">Missing Layout Columns:</span>
                          <ul className="list-disc list-inside text-slate-300 font-mono space-y-0.5">
                            {validationSummary.missingColumns.map((c, idx) => <li key={idx}>{c}</li>)}
                          </ul>
                        </div>
                      )}
                      {validationSummary.extraColumns.length > 0 && (
                        <div className="bg-slate-950 p-3 rounded-xl border border-slate-900">
                          <span className="text-amber-400 block font-bold text-[10px] uppercase tracking-wider mb-1">Extraneous Header Elements:</span>
                          <ul className="list-disc list-inside text-slate-300 font-mono space-y-0.5">
                            {validationSummary.extraColumns.map((c, idx) => <li key={idx}>{c}</li>)}
                          </ul>
                        </div>
                      )}
                    </div>
                  </div>
                )}

                {/* Content Cell Exception Rows Reporting List */}
                {!validationSummary.criticalHeaderError && validationSummary.rowErrors.length > 0 && (
                  <div className="bg-amber-950/15 border border-amber-900/40 rounded-xl p-4 text-xs">
                    <h4 className="font-bold text-amber-400 flex items-center gap-1.5 mb-2 text-sm">
                      <AlertTriangle size={15} /> Missing Mandatory Field Rows ({validationSummary.rowErrors.length} Lines)
                    </h4>
                    <div className="max-h-[160px] overflow-y-auto space-y-1.5 pr-1 font-mono">
                      {validationSummary.rowErrors.map((err, idx) => (
                        <div key={idx} className="bg-slate-950 p-2.5 rounded-lg border border-amber-900/20 flex gap-4 items-start">
                          <span className="text-amber-500 font-bold bg-amber-500/10 border border-amber-500/10 px-2 py-0.5 rounded text-[10px]">ROW {err.rowNumber}</span>
                          <div className="text-slate-300 space-y-0.5 flex-1 text-slate-400">
                            {err.messages.map((m, mIdx) => <p key={mIdx}>• {m}</p>)}
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {/* Successful Process Output Download Hub */}
                {isProcessed && (
                  <div className="bg-slate-950/40 rounded-xl p-5 border border-slate-800 space-y-5 animate-fadeIn">
                    <div className="flex items-center gap-2 text-emerald-400 font-bold text-sm border-b border-slate-900 pb-3">
                      <ShieldCheck size={16} /> Dynamic Schema Processing Parameters Enforced Natively
                    </div>
                    
                    <div className="grid grid-cols-2 md:grid-cols-3 gap-4 text-center max-w-xl text-xs">
                      <div className="bg-slate-900 p-3 rounded-xl border border-slate-800">
                        <span className="text-slate-500 block text-[9px] uppercase font-bold tracking-widest">Processed Rows</span>
                        <span className="text-lg font-black text-white font-mono mt-0.5 block">{processedData.length}</span>
                      </div>
                      <div className="bg-slate-900 p-3 rounded-xl border border-slate-800">
                        <span className="text-slate-500 block text-[9px] uppercase font-bold tracking-widest">Split Output Files</span>
                        <span className="text-lg font-black text-cyan-400 font-mono mt-0.5 block">{totalChunks} Excel Pack{totalChunks > 1 ? 's' : ''}</span>
                      </div>
                      <div className="bg-slate-900 p-3 rounded-xl border border-slate-800 col-span-2 md:col-span-1">
                        <span className="text-slate-500 block text-[9px] uppercase font-bold tracking-widest">Zip Packaging Strategy</span>
                        <span className={`text-xs font-bold mt-1.5 block ${processedData.length > 300 ? 'text-indigo-400' : 'text-emerald-400'}`}>
                          {processedData.length > 300 ? 'Active ZIP Bundle (>300 Chunks)' : 'Direct Document Download'}
                        </span>
                      </div>
                    </div>

                    <div className="pt-2 flex flex-col sm:flex-row flex-wrap gap-3">
                      <button 
                        onClick={() => executeDownloadAction('cleaned')}
                        className="bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-slate-950 text-xs font-black px-5 py-3 rounded-xl flex items-center justify-center gap-2 transition shadow-lg shadow-emerald-600/10"
                      >
                        <Download size={14} /> Download Cleaned Excel File (.xlsx)
                      </button>

                      {processedData.length > 300 && (
                        <button 
                          onClick={() => executeDownloadAction('zip')}
                          className="bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-bold px-5 py-3 rounded-xl flex items-center justify-center gap-2 transition shadow-lg shadow-indigo-600/10"
                        >
                          <FileArchive size={14} /> Download Structured ZIP Package Bundle
                        </button>
                      )}
                    </div>
                  </div>
                )}
              </section>
            )}
          </>
        )}
      </div>
    </div>
  );
}