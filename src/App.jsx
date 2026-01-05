import React, { useState } from 'react';
import { FileText, Table, X, Loader2, UploadCloud, CheckCircle, AlertCircle } from 'lucide-react';
import { motion as Motion, AnimatePresence } from 'framer-motion';
import ExcelJS from 'exceljs';
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
    if (e.dataTransfer && e.dataTransfer.files) {
      const droppedFiles = Array.from(e.dataTransfer.files);
      handleFiles(droppedFiles);
    }
  };

  const onFileInput = (e) => {
    if (e.target && e.target.files) {
      const selectedFiles = Array.from(e.target.files);
      handleFiles(selectedFiles);
      // Reset input value to allow selecting the same file again
      e.target.value = '';
    }
  };

  const handleFiles = (newFiles) => {
    const validFiles = newFiles.filter(file => {
      const name = file.name.toLowerCase();
      return name.endsWith('.xlsx') ||
             name.endsWith('.xls') ||
             name.endsWith('.docx');
    });

    const formattedFiles = validFiles.map(file => {
      const name = file.name.toLowerCase();
      return {
        id: Math.random().toString(36).slice(2, 11),
        file,
        name: file.name,
        type: (name.endsWith('.xlsx') || name.endsWith('.xls')) ? 'excel' : 'word',
        size: (file.size / 1024 / 1024).toFixed(2) + ' MB',
        status: 'idle', // idle, converting, success, error
        errorMessage: ''
      };
    });

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

  const convertExcelToImage = async (fileItem) => {
    return new Promise(async (resolve, reject) => {
      try {
        const fileName = fileItem.name.toLowerCase();
        const isXlsx = fileName.endsWith('.xlsx');

        if (isXlsx) {
          // Use ExcelJS for .xlsx files (better style support)
          const arrayBuffer = await fileItem.file.arrayBuffer();
          const workbook = new ExcelJS.Workbook();
          await workbook.xlsx.load(arrayBuffer);

          let combinedHtml = '';

          workbook.eachSheet((worksheet, sheetId) => {
            const sheetHtml = generateSheetHtmlWithExcelJS(worksheet);

            if (sheetHtml) {
              combinedHtml += `
                <div class="sheet-container">
                  <h2>${worksheet.name}</h2>
                  ${sheetHtml}
                </div>
              `;
            }
          });

          if (!combinedHtml) {
            throw new Error("No content found in Excel file.");
          }

          generateImage(combinedHtml, fileItem.name.replace(/\.[^/.]+$/, ""), resolve, reject, 'excel');
        } else {
          // Use SheetJS for .xls files (legacy format)
          const arrayBuffer = await fileItem.file.arrayBuffer();
          const data = new Uint8Array(arrayBuffer);
          const workbook = XLSX.read(data, { type: 'array', cellStyles: true });

          let combinedHtml = '';

          workbook.SheetNames.forEach(sheetName => {
            const worksheet = workbook.Sheets[sheetName];

            if (!worksheet['!ref']) return;

            const sheetHtml = generateSheetHtmlWithSheetJS(worksheet);

            if (sheetHtml) {
              combinedHtml += `
                <div class="sheet-container">
                  <h2>${sheetName}</h2>
                  ${sheetHtml}
                </div>
              `;
            }
          });

          if (!combinedHtml) {
            throw new Error("No content found in Excel file.");
          }

          generateImage(combinedHtml, fileItem.name.replace(/\.[^/.]+$/, ""), resolve, reject, 'excel');
        }
      } catch (err) {
        reject(err);
      }
    });
  };

  const generateSheetHtmlWithSheetJS = (worksheet) => {
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    const merges = worksheet['!merges'] || [];
    const cols = worksheet['!cols'] || [];
    const rows = worksheet['!rows'] || [];

    let html = '<table>';

    // Generate colgroup for column widths
    html += '<colgroup>';
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const col = cols[C];
      let width = 80;
      if (col) {
        if (col.wpx) width = col.wpx;
        else if (col.wch) width = col.wch * 8;
        else if (col.width) width = col.width * 8;
      }
      html += `<col style="width: ${Math.max(width, 30)}px;">`;
    }
    html += '</colgroup>';

    // Pre-calculate merged cells map
    const skipMap = new Set();
    const mergeMap = {};

    merges.forEach(merge => {
      const startR = merge.s.r;
      const startC = merge.s.c;
      const endR = merge.e.r;
      const endC = merge.e.c;

      mergeMap[`${startR},${startC}`] = {
        rowspan: endR - startR + 1,
        colspan: endC - startC + 1
      };

      for (let r = startR; r <= endR; r++) {
        for (let c = startC; c <= endC; c++) {
          if (r === startR && c === startC) continue;
          skipMap.add(`${r},${c}`);
        }
      }
    });

    for (let R = range.s.r; R <= range.e.r; ++R) {
      const row = rows[R];
      let rowHeight = 20;
      if (row) {
        if (row.hpx) rowHeight = row.hpx;
        else if (row.hpt) rowHeight = row.hpt * 1.33;
      }

      html += `<tr style="height: ${rowHeight}px;">`;

      for (let C = range.s.c; C <= range.e.c; ++C) {
        if (skipMap.has(`${R},${C}`)) continue;

        const cellAddress = { c: C, r: R };
        const cellRef = XLSX.utils.encode_cell(cellAddress);
        const cell = worksheet[cellRef];

        let cellValue = '';
        if (cell) {
          cellValue = (cell.w || cell.v || '').toString();
          cellValue = cellValue
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;')
            .replace(/\n/g, '<br>');
        }

        const mergeInfo = mergeMap[`${R},${C}`];
        const rowspan = mergeInfo ? `rowspan="${mergeInfo.rowspan}"` : '';
        const colspan = mergeInfo ? `colspan="${mergeInfo.colspan}"` : '';

        // Basic styling for .xls files
        let styleStr = 'border: 1px solid #d0d0d0; padding: 4px 6px; vertical-align: middle; ';

        // Number alignment
        if (cell && cell.t === 'n') {
          styleStr += 'text-align: right; ';
        }

        // Header rows heuristic
        if (R <= 2) {
          styleStr += 'font-weight: bold; background-color: #e8e8e8; text-align: center; ';
        }

        html += `<td ${rowspan} ${colspan} style="${styleStr}">${cellValue || '&nbsp;'}</td>`;
      }
      html += '</tr>';
    }
    html += '</table>';

    return html;
  };

  const generateSheetHtmlWithExcelJS = (worksheet) => {
    // Helper to convert RGB object to hex
    const rgbToHex = (rgb) => {
      if (!rgb) return null;
      if (typeof rgb === 'string') return rgb.startsWith('#') ? rgb : `#${rgb}`;
      if (rgb.argb) return `#${rgb.argb.substring(2)}`;
      return null;
    };

    // Helper to get border style
    const getBorderStyle = (border) => {
      if (!border || !border.style) return '1px solid #d0d0d0';
      const width = border.style === 'thin' ? '1px' : border.style === 'medium' ? '2px' : border.style === 'thick' ? '3px' : '1px';
      const color = border.color ? rgbToHex(border.color) || '#000' : '#000';
      return `${width} solid ${color}`;
    };

    // Build map of merged cells
    const mergedCellsMap = new Map();
    const skipCells = new Set();

    if (worksheet._merges && Object.keys(worksheet._merges).length > 0) {
      Object.values(worksheet._merges).forEach(merge => {
        const { top, left, bottom, right } = merge;
        const masterKey = `${top},${left}`;
        mergedCellsMap.set(masterKey, {
          rowspan: bottom - top + 1,
          colspan: right - left + 1
        });

        // Mark cells to skip
        for (let r = top; r <= bottom; r++) {
          for (let c = left; c <= right; c++) {
            if (r !== top || c !== left) {
              skipCells.add(`${r},${c}`);
            }
          }
        }
      });
    }

    // Get actual dimensions
    const rowCount = worksheet.rowCount || 100;
    const colCount = worksheet.columnCount || 20;

    let html = '<table>';

    // Generate colgroup for column widths
    html += '<colgroup>';
    for (let c = 1; c <= colCount; c++) {
      const col = worksheet.getColumn(c);
      const width = col.width ? col.width * 7 : 80;
      html += `<col style="width: ${width}px;">`;
    }
    html += '</colgroup>';

    // Iterate through rows
    for (let r = 1; r <= rowCount; r++) {
      const row = worksheet.getRow(r);
      const rowHeight = row.height ? row.height * 1.33 : 20;
      html += `<tr style="height: ${rowHeight}px;">`;

      for (let c = 1; c <= colCount; c++) {
        // Skip if this cell is part of a merge (but not the master)
        if (skipCells.has(`${r},${c}`)) {
          continue;
        }

        const cell = worksheet.getCell(r, c);

        let cellValue = '';
        if (cell.value !== null && cell.value !== undefined) {
          if (cell.value.richText) {
            cellValue = cell.value.richText.map(rt => rt.text).join('');
          } else if (typeof cell.value === 'object' && cell.value.text) {
            cellValue = cell.value.text;
          } else {
            cellValue = cell.value.toString();
          }
        }

        // HTML escape
        cellValue = cellValue
          .replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;')
          .replace(/"/g, '&quot;')
          .replace(/'/g, '&#039;')
          .replace(/\n/g, '<br>');

        // Handle merged cells
        let rowspan = '';
        let colspan = '';
        const mergeKey = `${r},${c}`;
        if (mergedCellsMap.has(mergeKey)) {
          const merge = mergedCellsMap.get(mergeKey);
          rowspan = merge.rowspan > 1 ? `rowspan="${merge.rowspan}"` : '';
          colspan = merge.colspan > 1 ? `colspan="${merge.colspan}"` : '';
        }

        // Build cell style
        let styleStr = '';

        // Borders
        if (cell.border) {
          if (cell.border.top) styleStr += `border-top: ${getBorderStyle(cell.border.top)}; `;
          if (cell.border.bottom) styleStr += `border-bottom: ${getBorderStyle(cell.border.bottom)}; `;
          if (cell.border.left) styleStr += `border-left: ${getBorderStyle(cell.border.left)}; `;
          if (cell.border.right) styleStr += `border-right: ${getBorderStyle(cell.border.right)}; `;
        } else {
          styleStr += 'border: 1px solid #d0d0d0; ';
        }

        styleStr += 'padding: 4px 6px; ';

        // Background color
        if (cell.fill && cell.fill.type === 'pattern' && cell.fill.fgColor) {
          const bgColor = rgbToHex(cell.fill.fgColor);
          if (bgColor) styleStr += `background-color: ${bgColor}; `;
        }

        // Font styles
        if (cell.font) {
          if (cell.font.size) styleStr += `font-size: ${cell.font.size}px; `;
          if (cell.font.bold) styleStr += 'font-weight: bold; ';
          if (cell.font.italic) styleStr += 'font-style: italic; ';
          if (cell.font.underline) styleStr += 'text-decoration: underline; ';
          if (cell.font.color) {
            const fontColor = rgbToHex(cell.font.color);
            if (fontColor) styleStr += `color: ${fontColor}; `;
          }
          if (cell.font.name) styleStr += `font-family: "${cell.font.name}", sans-serif; `;
        }

        // Text alignment
        if (cell.alignment) {
          if (cell.alignment.horizontal) {
            const hAlign = cell.alignment.horizontal;
            styleStr += `text-align: ${hAlign === 'center' ? 'center' : hAlign === 'right' ? 'right' : 'left'}; `;
          }
          if (cell.alignment.vertical) {
            const vAlign = cell.alignment.vertical;
            styleStr += `vertical-align: ${vAlign === 'middle' ? 'middle' : vAlign === 'top' ? 'top' : 'bottom'}; `;
          }
          if (cell.alignment.wrapText) {
            styleStr += 'white-space: pre-wrap; word-wrap: break-word; ';
          } else {
            styleStr += 'white-space: nowrap; ';
          }
        }

        html += `<td ${rowspan} ${colspan} style="${styleStr}">${cellValue || '&nbsp;'}</td>`;
      }

      html += '</tr>';
    }

    html += '</table>';
    return html;
  };

  const convertWordToImage = (fileItem) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        mammoth.convertToHtml({ arrayBuffer: e.target.result })
          .then(result => {
            const html = `<div class="word-content">${result.value}</div>`;
            generateImage(html, fileItem.name.replace(/\.[^/.]+$/, ""), resolve, reject, 'word');
          })
          .catch(err => reject(err));
      };
      reader.readAsArrayBuffer(fileItem.file);
    });
  };

  const generateImage = async (htmlContent, fileName, resolve, reject, fileType) => {
    console.log('Generating image for:', fileName, 'Type:', fileType);

    // Create a temporary iframe for complete style isolation
    const iframe = document.createElement('iframe');
    iframe.style.position = 'absolute';
    iframe.style.left = '-9999px';
    iframe.style.top = '0';
    iframe.style.width = '1400px'; // Set a fixed width or make it dynamic
    iframe.style.height = 'auto';
    iframe.style.visibility = 'visible'; // Needed for html2canvas
    iframe.style.zIndex = '-9999';
    iframe.style.border = 'none';

    document.body.appendChild(iframe);

    try {
      const doc = iframe.contentWindow.document;
      doc.open();
      doc.write(`
        <!DOCTYPE html>
        <html>
          <head>
            <style>
              body {
                background-color: #ffffff;
                margin: 0;
                padding: 40px;
                font-family: "MS PGothic", "MS Gothic", "Meiryo", Arial, sans-serif;
                font-size: 11px;
                color: #000000;
              }
              table {
                border-collapse: collapse;
                background: #ffffff;
                margin-bottom: 30px;
                table-layout: fixed;
              }
              td {
                overflow: hidden;
                text-overflow: ellipsis;
              }
              .sheet-container {
                margin-bottom: 50px;
                background-color: #ffffff;
                page-break-after: always;
              }
              h2 {
                font-size: 16px;
                color: #000000;
                border-bottom: 2px solid #000000;
                padding-bottom: 8px;
                margin-bottom: 20px;
                margin-top: 0;
                font-weight: bold;
                font-family: "MS PGothic", "MS Gothic", "Meiryo", Arial, sans-serif;
              }
              .word-content {
                color: #000000;
                line-height: 1.6;
                max-width: 800px;
              }
              .word-content p {
                margin-bottom: 10px;
              }
            </style>
          </head>
          <body>
            <div id="render-content">${htmlContent}</div>
          </body>
        </html>
      `);
      doc.close();

      // Wait for content render
      await new Promise(r => setTimeout(r, 500));

      // Adjust iframe height to fit content
      const contentHeight = doc.body.scrollHeight;
      const contentWidth = doc.body.scrollWidth;
      iframe.style.height = (contentHeight + 100) + 'px';
      iframe.style.width = (Math.max(contentWidth, 1000)) + 'px';

      const canvas = await html2canvas(doc.body, {
        scale: 2,
        useCORS: true,
        backgroundColor: '#ffffff',
        logging: false,
        width: Math.max(contentWidth, 1000),
        height: contentHeight
      });

      canvas.toBlob((blob) => {
        if (!blob) {
          reject(new Error('Canvas to Blob conversion failed'));
          return;
        }

        const url = URL.createObjectURL(blob);

        // Clean up
        if (document.body.contains(iframe)) {
          document.body.removeChild(iframe);
        }

        // Trigger download
        const link = document.createElement('a');
        link.href = url;
        link.download = `${fileName}.png`;
        document.body.appendChild(link);
        link.click();

        // Clean up URL and link
        setTimeout(() => {
          if (document.body.contains(link)) {
            document.body.removeChild(link);
          }
          URL.revokeObjectURL(url);
        }, 100);

        resolve();
      }, 'image/png');

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
              <Motion.div
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
              </Motion.div>
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
