import React, { useState, useCallback } from 'react';
import { FileText, Table, X, Download, Loader2, UploadCloud, CheckCircle, AlertCircle } from 'lucide-react';
import { motion, AnimatePresence } from 'framer-motion';
import * as XLSX from 'xlsx';
import mammoth from 'mammoth';
import html2canvas from 'html2canvas';

function App() {
  const [files, setFiles] = useState([]);
  const [isDragging, setIsDragging] = useState(false);

  // File Handling
  const onDragOver = (e) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const onDragLeave = (e) => {
    e.preventDefault();
    setIsDragging(false);
  };

  const onDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);
    const droppedFiles = Array.from(e.dataTransfer.files);
    handleFiles(droppedFiles);
  };

  const onFileInput = (e) => {
    const selectedFiles = Array.from(e.target.files);
    handleFiles(selectedFiles);
  };

  const handleFiles = (newFiles) => {
    const validFiles = newFiles.filter(file =>
      file.name.endsWith('.xlsx') ||
      file.name.endsWith('.xls') ||
      file.name.endsWith('.docx')
    );

    const formattedFiles = validFiles.map(file => ({
      id: Math.random().toString(36).substr(2, 9),
      file,
      name: file.name,
      type: (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) ? 'excel' : 'word',
      size: (file.size / 1024 / 1024).toFixed(2) + ' MB',
      status: 'idle', // idle, converting, success, error
      errorMessage: ''
    }));

    setFiles(prev => [...prev, ...formattedFiles]);
  };

  const removeFile = (id) => {
    setFiles(prev => prev.filter(f => f.id !== id));
  };

  const convertFile = async (fileItem) => {
    setFiles(prev => prev.map(f => f.id === fileItem.id ? { ...f, status: 'converting' } : f));

    try {
      if (fileItem.type === 'excel') {
        await convertExcelToImage(fileItem);
      } else {
        await convertWordToImage(fileItem);
      }

      setFiles(prev => prev.map(f => f.id === fileItem.id ? { ...f, status: 'success' } : f));
    } catch (error) {
      console.error(error);
      setFiles(prev => prev.map(f => f.id === fileItem.id ? { ...f, status: 'error', errorMessage: 'Conversion failed' } : f));
    }
  };

  const convertExcelToImage = (fileItem) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });

          // Convert first sheet to HTML
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const html = XLSX.utils.sheet_to_html(worksheet);

          generateImage(html, fileItem.name.replace(/\.[^/.]+$/, ""), resolve, reject);
        } catch (err) {
          reject(err);
        }
      };
      reader.readAsArrayBuffer(fileItem.file);
    });
  };

  const convertWordToImage = (fileItem) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        mammoth.convertToHtml({ arrayBuffer: e.target.result })
          .then(result => {
            const html = `<div class="word-content">${result.value}</div>`;
            generateImage(html, fileItem.name.replace(/\.[^/.]+$/, ""), resolve, reject);
          })
          .catch(err => reject(err));
      };
      reader.readAsArrayBuffer(fileItem.file);
    });
  };

  const generateImage = async (htmlContent, fileName, resolve, reject) => {
    // Create a temporary iframe to isolate styles from Tailwind v4 (oklch issue)
    const iframe = document.createElement('iframe');
    iframe.style.position = 'fixed';
    iframe.style.top = '0';
    iframe.style.left = '0';
    iframe.style.width = '800px';
    iframe.style.height = '1200px';
    iframe.style.zIndex = '-9999';
    iframe.style.opacity = '0';
    iframe.style.pointerEvents = 'none';
    iframe.style.border = 'none';

    document.body.appendChild(iframe);

    try {
      const doc = iframe.contentWindow.document;
      doc.open();
      doc.write(`
        <html>
          <head>
            <style>
              body {
                font-family: Arial, sans-serif;
                font-size: 12px;
                color: #000000;
                background-color: #ffffff;
                padding: 40px;
                margin: 0;
              }
              table {
                width: 100%;
                border-collapse: collapse;
                margin-bottom: 20px;
                border: 1px solid #dddddd;
              }
              td, th {
                border: 1px solid #dddddd;
                padding: 8px;
                background-color: #ffffff;
                color: #000000;
              }
              .word-content {
                color: #000000;
              }
            </style>
          </head>
          <body>
            ${htmlContent}
          </body>
        </html>
      `);
      doc.close();

      // Wait for content render
      await new Promise(r => setTimeout(r, 500));

      const canvas = await html2canvas(doc.body, {
        scale: 2,
        useCORS: true,
        backgroundColor: '#ffffff'
      });

      const image = canvas.toDataURL("image/png");

      // Trigger download
      const link = document.createElement('a');
      link.href = image;
      link.download = `${fileName}.png`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

      // Clean up extendedly to ensure no memory leak
      setTimeout(() => {
        if (document.body.contains(iframe)) {
          document.body.removeChild(iframe);
        }
      }, 100);

      resolve();
    } catch (err) {
      if (document.body.contains(iframe)) {
        document.body.removeChild(iframe);
      }
      reject(err);
    }
  };


  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 p-8 flex flex-col items-center">
      <header className="mb-12 text-center max-w-2xl mx-auto">
        <h1 className="text-4xl font-bold bg-gradient-to-r from-cyan-400 to-blue-500 bg-clip-text text-transparent mb-4">
          Document Converter
        </h1>
        <div className="space-y-4 text-slate-400 mb-8">
          <p className="text-lg">
            EXCELやWord (.docx) ファイルを、ブラウザ上で瞬時に<strong>画像 (PNG)</strong> へ変換するツールです。
          </p>
          <div className="bg-slate-900/50 p-4 rounded-xl border border-slate-800 text-sm text-left mx-auto max-w-lg space-y-2">
            <div className="flex items-start gap-2">
              <CheckCircle className="w-4 h-4 text-cyan-400 mt-0.5 shrink-0" />
              <span><strong className="text-slate-200">完全プライベート:</strong> ファイルはサーバーに送信されず、すべてお使いの端末内で安全に処理されます。機密文書も安心です。</span>
            </div>
            <div className="flex items-start gap-2">
              <CheckCircle className="w-4 h-4 text-cyan-400 mt-0.5 shrink-0" />
              <span><strong className="text-slate-200">対応フォーマット:</strong> Excel (.xlsx, .xls) および Word (.docx) に対応しています。</span>
            </div>
            <div className="flex items-start gap-2">
              <CheckCircle className="w-4 h-4 text-cyan-400 mt-0.5 shrink-0" />
              <span><strong className="text-slate-200">インストール不要:</strong> 特別なソフトをインストールすることなく、ブラウザのみで完結します。</span>
            </div>
          </div>
        </div>
      </header>

      <main className="w-full max-w-2xl">
        {/* Drop Zone */}
        <div
          className={`
            border-2 border-dashed rounded-2xl p-12 mb-8 transition-all duration-300 cursor-pointer
            flex flex-col items-center justify-center gap-4 group
            ${isDragging ? 'border-blue-500 bg-blue-500/10' : 'border-slate-700 hover:border-blue-400 hover:bg-slate-900'}
          `}
          onDragOver={onDragOver}
          onDragLeave={onDragLeave}
          onDrop={onDrop}
          onClick={() => document.getElementById('fileInput').click()}
        >
          <input
            type="file"
            id="fileInput"
            multiple
            accept=".xlsx,.xls,.docx"
            className="hidden"
            onChange={onFileInput}
          />
          <div className="p-4 rounded-full bg-slate-800 group-hover:bg-slate-700 transition-colors">
            <UploadCloud className={`w-10 h-10 ${isDragging ? 'text-blue-400' : 'text-slate-400'}`} />
          </div>
          <div className="text-center">
            <p className="text-lg font-medium text-slate-200">Click to upload or drag and drop</p>
            <p className="text-sm text-slate-500 mt-1">Excel (.xlsx, .xls) または Word (.docx) のみ対応</p>
          </div>
        </div>

        {/* File List */}
        <div className="space-y-4">
          <AnimatePresence>
            {files.map(file => (
              <motion.div
                key={file.id}
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, height: 0 }}
                className="bg-slate-900 border border-slate-800 rounded-xl p-4 flex items-center gap-4 shadow-sm"
              >
                <div className={`p-3 rounded-lg ${file.type === 'excel' ? 'bg-green-500/10 text-green-400' : 'bg-blue-500/10 text-blue-400'}`}>
                  {file.type === 'excel' ? <Table className="w-6 h-6" /> : <FileText className="w-6 h-6" />}
                </div>

                <div className="flex-1 min-w-0">
                  <h3 className="font-medium text-slate-200 truncate">{file.name}</h3>
                  <p className="text-sm text-slate-500">{file.size}</p>
                </div>

                <div className="flex items-center gap-3">
                  {file.status === 'idle' && (
                    <button
                      onClick={() => convertFile(file)}
                      className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white text-sm font-medium rounded-lg transition-colors flex items-center gap-2"
                    >
                      Convert
                    </button>
                  )}

                  {file.status === 'converting' && (
                    <div className="flex items-center gap-2 text-blue-400">
                      <Loader2 className="w-5 h-5 animate-spin" />
                      <span className="text-sm">Processing...</span>
                    </div>
                  )}

                  {file.status === 'success' && (
                    <div className="flex items-center gap-2 text-green-400">
                      <CheckCircle className="w-5 h-5" />
                      <span className="text-sm">Done</span>
                    </div>
                  )}

                  {file.status === 'error' && (
                    <div className="flex items-center gap-2 text-red-400" title={file.errorMessage}>
                      <AlertCircle className="w-5 h-5" />
                      <span className="text-sm">Failed</span>
                    </div>
                  )}

                  <button
                    onClick={() => removeFile(file.id)}
                    className="p-2 text-slate-500 hover:text-slate-300 hover:bg-slate-800 rounded-lg transition-colors"
                  >
                    <X className="w-5 h-5" />
                  </button>
                </div>
              </motion.div>
            ))}
          </AnimatePresence>

          {files.length === 0 && (
            <div className="text-center text-slate-600 py-8">
              No files uploaded yet.
            </div>
          )}
        </div>
      </main>
    </div>
  );
}

export default App;
