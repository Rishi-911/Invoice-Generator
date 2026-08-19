import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import confetti from 'canvas-confetti';
import axios from "axios"
import { 
  FileText, 
  FileSpreadsheet, 
  Check, 
  AlertCircle, 
  UploadCloud, 
  Trash2, 
  Download, 
  Sparkles, 
  CheckCircle2, 
  Loader2
} from 'lucide-react';

export default function App() {
  const [templateFile, setTemplateFile] = useState(null);
  const [dataFile, setDataFile] = useState(null);
  const [templateTags, setTemplateTags] = useState([]);
  const [sheetData, setSheetData] = useState({ headers: [], rows: [] });
  const [templateDragActive, setTemplateDragActive] = useState(false);
  const [dataDragActive, setDataDragActive] = useState(false);
  const [isValidating, setIsValidating] = useState(false);
  const [isGenerating, setIsGenerating] = useState(false);
  const [generatedInvoices, setGeneratedInvoices] = useState([]);
  const [downloadZipUrl, setDownloadZipUrl] = useState(null);
  const [zipFileName, setZipFileName] = useState('');

  const templateInputRef = useRef(null);
  const dataInputRef = useRef(null);

  const parseTemplate = async (file) => {
    try {
      const arrayBuffer = await file.arrayBuffer();
      const zip = new PizZip(arrayBuffer);
      
      for (const filename of Object.keys(zip.files)) {
        if (filename.startsWith("word/") && filename.endsWith(".xml")) {
          const fileObj = zip.file(filename);
          if (fileObj) {
            let xmlText = fileObj.asText();
            const normalizedXml = xmlText.replace(/«/g, "&lt;&lt;").replace(/»/g, "&gt;&gt;");
            zip.file(filename, normalizedXml);
          }
        }
      }

      const doc = new Docxtemplater(zip, {
        delimiters: { start: "<<", end: ">>" },
      });
      
      const placeholders = new Set();
      for (const filename of Object.keys(zip.files)) {
        if (filename.startsWith("word/") && filename.endsWith(".xml")) {
          const fileObj = zip.file(filename);
          if (fileObj) {
            let xmlText = fileObj.asText();
            let cleanText = xmlText.replace(/<[^>]+>/g, "")
              .replace(/&lt;/g, "<")
              .replace(/&gt;/g, ">")
              .replace(/&amp;/g, "&")
              .replace(/&quot;/g, '"')
              .replace(/&apos;/g, "'");

            const regex = /<<\s*([^>]+?)\s*>>/g;
            let match;
            while ((match = regex.exec(cleanText)) !== null) {
              let tag = match[1].trim();
              if (tag.startsWith('#') || tag.startsWith('/') || tag.startsWith('^')) {
                tag = tag.substring(1).trim();
              }
              if (tag) {
                placeholders.add(tag);
              }
            }
          }
        }
      }
      
      const tags = Array.from(placeholders);
      setTemplateTags(tags);
      return tags;
    } catch (error) {
      console.error("Error parsing Word template:", error);
      alert("Failed to parse the Word template. Please ensure it is a valid .docx file.");
      return [];
    }
  };

  const parseDataSheet = async (file) => {
    try {
      const arrayBuffer = await file.arrayBuffer();
      const data = new Uint8Array(arrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      
      const headers = (rawRows[0] || []).map(val => String(val || '').trim()).filter(Boolean);
      const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "" }).map(row => {
        const cleanRow = {};
        for (const key of Object.keys(row)) {
          cleanRow[key.trim()] = row[key];
        }
        return cleanRow;
      });

      const result = { headers, rows };
      setSheetData(result);
      return result;
    } catch (error) {
      console.error("Error parsing spreadsheet:", error);
      alert("Failed to parse the Excel spreadsheet. Please ensure it is a valid .xlsx file.");
      return { headers: [], rows: [] };
    }
  };

  const handleTemplateChange = async (file) => {
    if (!file) return;
    if (!file.name.endsWith('.docx')) {
      alert("Only Word templates (.docx) are supported.");
      return;
    }
     try {
    setIsValidating(true);

    const buffer = await file.arrayBuffer();

    const stableFile = new File(
      [buffer],
      file.name,
      {
        type: file.type || 
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      }
    );

    setTemplateFile(stableFile);
    setGeneratedInvoices([]);
    setDownloadZipUrl(null);

    await parseTemplate(stableFile);

  } catch (error) {
    console.error("Error reading template:", error);
    alert("Could not read the Word template.");
  } finally {
    setIsValidating(false);
  }
  };

  const handleDataChange = async (file) => {
    if (!file) return;
    if (!file.name.endsWith('.xlsx')) {
      alert("Only Excel spreadsheets (.xlsx) are supported.");
      return;
    }
     try {
    setIsValidating(true);

    const buffer = await file.arrayBuffer();

    const stableFile = new File(
      [buffer],
      file.name,
      {
        type: file.type ||
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      }
    );

    setDataFile(stableFile);
    setGeneratedInvoices([]);
    setDownloadZipUrl(null);

    await parseDataSheet(stableFile);

  } catch (error) {
    console.error("Error reading spreadsheet:", error);
    alert("Could not read the Excel file.");
  } finally {
    setIsValidating(false);
  }
  };

  const matchedTags = templateTags.filter(tag => 
    sheetData.headers.some(header => header.toLowerCase() === tag.toLowerCase())
  );
  const missingTags = templateTags.filter(tag => 
    !sheetData.headers.some(header => header.toLowerCase() === tag.toLowerCase())
  );
  
  const isReadyToGenerate = templateFile && dataFile && templateTags.length > 0 && sheetData.rows.length > 0;

  const generateInvoices = async () => {
    if (!isReadyToGenerate) return;
    setIsGenerating(true);
    
    try {
      const formData = new FormData();
      formData.append("template", templateFile);
      formData.append("data", dataFile);

      const response = await axios.post("http://127.0.0.1:8000/api/generate",
        formData,
        {
          responseType : "blob",
        }
      );

      const zipBlob = response.data;
     
      const zipUrl = URL.createObjectURL(zipBlob);
      const zipName = `FastInvoices_PDFs_${new Date().toISOString().slice(0, 10)}.zip`;
      
      setZipFileName(zipName);
      setDownloadZipUrl(zipUrl);

      const list = sheetData.rows.map(row => ({
        name: `${row.Name || 'Invoice'}.pdf`,
        url: ''
      }));
      setGeneratedInvoices(list);

      const link = document.createElement('a');
      link.href = zipUrl;
      link.download = zipName;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

      confetti({
        particleCount: 150,
        spread: 80,
        origin: { y: 0.6 }
      });
    } catch (error) {
      console.error("Error generating invoices:", error);
      alert("An error occurred during invoice generation: " + error.message);
    } finally {
      setIsGenerating(false);
    }
  };

  const removeTemplateFile = () => {
    setTemplateFile(null);
    setTemplateTags([]);
    setGeneratedInvoices([]);
    setDownloadZipUrl(null);
    if (templateInputRef.current) templateInputRef.current.value = '';
  };

  const removeDataFile = () => {
    setDataFile(null);
    setSheetData({ headers: [], rows: [] });
    setGeneratedInvoices([]);
    setDownloadZipUrl(null);
    if (dataInputRef.current) dataInputRef.current.value = '';
  };

  return (
    <div className="flex flex-col min-h-screen">
      <header className="header-banner">
        <h1>FASTINVOICE: Automate Your Billing</h1>
      </header>

      <main className="app-container">
        <section>
          <h2 className="section-title">How It Works</h2>
          <div className="how-it-works-grid">
            <div className="how-card card-docx">
              <h3>1. Prepare Template (DOCX)</h3>
              <div className="mockup-container">
                <div className="word-mockup">
                  <div className="doc-layout">
                    <div className="doc-line" style={{ width: '80%' }}></div>
                    <div className="doc-tag blue">&lt;&lt;Customer_Name&gt;&gt;</div>
                    <div className="doc-line"></div>
                    <div className="doc-tag purple">&lt;&lt;Invoice_Date&gt;&gt;</div>
                    <div className="doc-line" style={{ width: '90%' }}></div>
                    <div className="doc-tag orange">&lt;&lt;Amount&gt;&gt;</div>
                  </div>
                  <div className="callout-bubble">
                    <strong>MANUALLY MATCH TAGS EXACTLY!</strong>
                    <span>Wrap headers like &lt;&lt;Amount&gt;&gt; in double angle brackets.</span>
                    <span>Example: Column heading "Amount" must be `&lt;&lt;Amount&gt;&gt;` in DOCX.</span>
                  </div>
                </div>
              </div>
            </div>

            <div className="how-card card-xlsx">
              <h3>2. Upload Data (XLSX)</h3>
              <div className="mockup-container">
                <div className="excel-mockup">
                  <div className="excel-header">
                    <div className="excel-dot"></div>
                    <div className="excel-dot"></div>
                    <div className="excel-dot"></div>
                  </div>
                  <div className="excel-grid">
                    <div className="excel-cell header">A</div>
                    <div className="excel-cell header">B</div>
                    <div className="excel-cell header">C</div>
                    
                    <div className="excel-cell">Customer_Name</div>
                    <div className="excel-cell">Invoice_Date</div>
                    <div className="excel-cell">Amount</div>

                    <div className="excel-cell" style={{ color: '#94A3B8' }}>John Doe</div>
                    <div className="excel-cell" style={{ color: '#94A3B8' }}>2026-08-06</div>
                    <div className="excel-cell" style={{ color: '#94A3B8' }}>$450.00</div>
                  </div>
                  <button className="excel-upload-btn">
                    <UploadCloud size={8} /> Upload
                  </button>
                  <img 
                    src="https://img.icons8.com/color/48/pointer.png" 
                    alt="cursor" 
                    className="click-hand"
                  />
                </div>
              </div>
            </div>

            <div className="how-card card-generate">
              <h3>3. Generate Invoices (PDF/DOCX)</h3>
              <div className="mockup-container">
                <div className="generate-mockup">
                  <div className="file-stack">
                    <div className="file-icon docx">DOCX</div>
                    <div className="file-icon pdf">PDF</div>
                    <div className="file-icon green-action">
                      <Sparkles size={16} />
                      <span style={{ fontSize: '7px' }}>DONE</span>
                    </div>
                  </div>
                  <div className="download-pill-btn">
                    <Download size={12} /> Download
                  </div>
                </div>
              </div>
            </div>
          </div>

          <ul className="instructions-list" style={{ marginTop: '1.5rem' }}>
            <li>Use &lt;&lt;ColumnName&gt;&gt; syntax</li>
            <li>Match exact capitalization</li>
            <li>No extra spaces within brackets</li>
          </ul>
        </section>

        <section>
          <h2 className="section-title">Invoice Workspace</h2>
          <div className="workspace-grid">
            <div 
              className={`dropzone-container docx-accent ${templateDragActive ? 'drag-active' : ''} ${templateFile ? 'has-file' : ''}`}
              onDragOver={(e) => { e.preventDefault(); setTemplateDragActive(true); }}
              onDragLeave={() => setTemplateDragActive(false)}
              onDrop={(e) => {
                e.preventDefault();
                setTemplateDragActive(false);
                if (e.dataTransfer.files && e.dataTransfer.files[0]) {
                  handleTemplateChange(e.dataTransfer.files[0]);
                }
              }}
              onClick={() => { if (!templateFile) templateInputRef.current.click(); }}
            >
              <input 
                type="file" 
                ref={templateInputRef} 
                style={{ display: 'none' }} 
                accept=".docx" 
                onChange={(e) => handleTemplateChange(e.target.files[0])} 
              />
              
              {!templateFile ? (
                <>
                  <div className="icon-circle">
                    <FileText size={36} />
                  </div>
                  <h3 className="dropzone-title">Upload Invoice Template (DOCX)</h3>
                  <p className="dropzone-subtitle">Drag and drop your word document template here, or click to browse</p>
                </>
              ) : (
                <div className="file-details-card">
                  <div className="file-header-row">
                    <div className="invoice-item-info">
                      <FileText size={28} style={{ color: '#2196F3' }} />
                      <div className="file-info-text">
                        <span className="file-name">{templateFile.name}</span>
                        <span className="file-size">{(templateFile.size / 1024).toFixed(1)} KB</span>
                      </div>
                    </div>
                    <button 
                      className="remove-file-btn" 
                      onClick={(e) => { e.stopPropagation(); removeTemplateFile(); }}
                    >
                      <Trash2 size={16} />
                    </button>
                  </div>
                  
                  {templateTags.length > 0 ? (
                    <>
                      <div className="parsed-meta-badge">
                        <Sparkles size={14} />
                        <span>{templateTags.length} placeholder tags detected</span>
                      </div>
                      <div className="tags-scroll-container">
                        {templateTags.map((tag, idx) => {
                          const isMatched = sheetData.headers.some(h => h.toLowerCase() === tag.toLowerCase());
                          return (
                            <span 
                              key={idx} 
                              className={`tag-bubble ${dataFile ? (isMatched ? 'matched' : 'mismatched') : ''}`}
                            >
                              &lt;&lt;{tag}&gt;&gt;
                            </span>
                          );
                        })}
                      </div>
                    </>
                  ) : (
                    <div className="parsed-meta-badge" style={{ backgroundColor: '#FFF9C4', color: '#F57F17' }}>
                      <AlertCircle size={14} />
                      <span>No tags detected. Make sure to use &lt;&lt;TagName&gt;&gt; format.</span>
                    </div>
                  )}
                </div>
              )}
            </div>

            <div 
              className={`dropzone-container xlsx-accent ${dataDragActive ? 'drag-active' : ''} ${dataFile ? 'has-file' : ''}`}
              onDragOver={(e) => { e.preventDefault(); setDataDragActive(true); }}
              onDragLeave={() => setDataDragActive(false)}
              onDrop={(e) => {
                e.preventDefault();
                setDataDragActive(false);
                if (e.dataTransfer.files && e.dataTransfer.files[0]) {
                  handleDataChange(e.dataTransfer.files[0]);
                }
              }}
              onClick={() => { if (!dataFile) dataInputRef.current.click(); }}
            >
              <input 
                type="file" 
                ref={dataInputRef} 
                style={{ display: 'none' }} 
                accept=".xlsx" 
                onChange={(e) => handleDataChange(e.target.files[0])} 
              />
              
              {!dataFile ? (
                <>
                  <div className="icon-circle">
                    <FileSpreadsheet size={36} />
                  </div>
                  <h3 className="dropzone-title">Upload Customer Data (XLSX)</h3>
                  <p className="dropzone-subtitle">Drag and drop your spreadsheet data here, or click to browse</p>
                </>
              ) : (
                <div className="file-details-card">
                  <div className="file-header-row">
                    <div className="invoice-item-info">
                      <FileSpreadsheet size={28} style={{ color: '#2E7D32' }} />
                      <div className="file-info-text">
                        <span className="file-name">{dataFile.name}</span>
                        <span className="file-size">{(dataFile.size / 1024).toFixed(1)} KB</span>
                      </div>
                    </div>
                    <button 
                      className="remove-file-btn" 
                      onClick={(e) => { e.stopPropagation(); removeDataFile(); }}
                    >
                      <Trash2 size={16} />
                    </button>
                  </div>

                  {sheetData.rows.length > 0 ? (
                    <>
                      <div className="parsed-meta-badge">
                        <Check size={14} />
                        <span>{sheetData.rows.length} rows & {sheetData.headers.length} columns parsed</span>
                      </div>
                      <div className="tags-scroll-container">
                        {sheetData.headers.map((header, idx) => {
                          const isUsed = templateTags.some(t => t.toLowerCase() === header.toLowerCase());
                          return (
                            <span 
                              key={idx} 
                              className={`tag-bubble ${templateFile ? (isUsed ? 'matched' : '') : ''}`}
                            >
                              {header}
                            </span>
                          );
                        })}
                      </div>
                    </>
                  ) : (
                    <div className="parsed-meta-badge" style={{ backgroundColor: '#FFF9C4', color: '#F57F17' }}>
                      <AlertCircle size={14} />
                      <span>Empty spreadsheet or header row not detected.</span>
                    </div>
                  )}
                </div>
              )}
            </div>
          </div>
        </section>

        <section className="action-section-container">
          <div className={`validation-card ${isValidating ? 'validating' : isReadyToGenerate ? (missingTags.length === 0 ? 'success' : 'warning') : ''}`}>
            <div className="validation-icon">
              {isValidating ? (
                <Loader2 className="spin" size={28} style={{ color: '#2196F3' }} />
              ) : isReadyToGenerate ? (
                missingTags.length === 0 ? (
                  <CheckCircle2 size={28} style={{ color: '#10B981' }} />
                ) : (
                  <AlertCircle size={28} style={{ color: '#F59E0B' }} />
                )
              ) : (
                <AlertCircle size={28} style={{ color: '#CBD5E1' }} />
              )}
            </div>
            <div className="validation-content">
              <span className="validation-title">
                {isValidating ? (
                  "Validating..."
                ) : !templateFile && !dataFile ? (
                  "Waiting for files..."
                ) : templateFile && !dataFile ? (
                  "Waiting for customer data spreadsheet..."
                ) : !templateFile && dataFile ? (
                  "Waiting for Word invoice template..."
                ) : templateTags.length === 0 ? (
                  "No placeholder tags found in template!"
                ) : missingTags.length === 0 ? (
                  `${matchedTags.length} matching tags found!`
                ) : (
                  `${matchedTags.length} of ${templateTags.length} tags matching`
                )}
              </span>
              <span className="validation-subtitle">
                {isValidating ? (
                  "Analyzing template content and column fields..."
                ) : !templateFile || !dataFile ? (
                  "Upload both template and spreadsheet to begin validation."
                ) : templateTags.length === 0 ? (
                  "Ensure your Word doc contains fields like <<Name>>."
                ) : missingTags.length === 0 ? (
                  "All template placeholders match data columns perfectly."
                ) : (
                  `Missing fields in sheet: ${missingTags.map(t => `<<${t}>>`).join(', ')}`
                )}
              </span>
            </div>
          </div>

          <button 
            className="generate-btn"
            disabled={!isReadyToGenerate || isGenerating}
            onClick={generateInvoices}
          >
            {isGenerating ? (
              <>
                <Loader2 className="spin" size={20} />
                Generating...
              </>
            ) : (
              "Generate Invoices"
            )}
          </button>
        </section>

        {generatedInvoices.length > 0 && (
          <section className="result-section">
            <div className="result-header">
              <h3 className="result-title">Generated Invoices ({generatedInvoices.length})</h3>
              {downloadZipUrl && (
                <a 
                  href={downloadZipUrl} 
                  download={zipFileName}
                  className="download-pill-btn"
                  style={{ textDecoration: 'none', background: '#00E676', borderColor: '#00C853' }}
                >
                  <Download size={14} /> Download ZIP Archive
                </a>
              )}
            </div>
            <div className="invoices-grid">
              {generatedInvoices.map((inv, idx) => (
                <div className="invoice-item-card" key={idx}>
                  <div className="invoice-item-info">
                    <FileText size={18} style={{ color: '#2196F3' }} />
                    <span className="invoice-item-name" title={inv.name}>{inv.name}</span>
                  </div>
                  {inv.url && (
                    <a 
                      href={inv.url} 
                      download={inv.name} 
                      className="invoice-download-btn"
                      title="Download individual file"
                    >
                      <Download size={14} />
                    </a>
                  )}
                </div>
              ))}
            </div>
          </section>
        )}
      </main>
    </div>
  );
}
