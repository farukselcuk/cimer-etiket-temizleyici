import React, { useState } from 'react';
import ExcelJS from 'exceljs';
import logo from './logo.png';
import './App.css';

function App() {
  const [file, setFile] = useState(null);
  const [processing, setProcessing] = useState(false);
  const [status, setStatus] = useState('');
  const [dragActive, setDragActive] = useState(false);

  const removeHtmlTags = (text) => {
    if (!text || typeof text !== 'string') return text;
    
    // First decode HTML entities
    let cleaned = text
      .replace(/&nbsp;/gi, ' ')
      .replace(/&lt;/gi, '<')
      .replace(/&gt;/gi, '>')
      .replace(/&amp;/gi, '&')
      .replace(/&quot;/gi, '"')
      .replace(/&#39;/gi, "'")
      .replace(/&#34;/gi, '"')
      .replace(/&rdquo;/gi, '"')
      .replace(/&ldquo;/gi, '"')
      .replace(/&rsquo;/gi, "'")
      .replace(/&lsquo;/gi, "'");
    
    // Then remove HTML tags
    cleaned = cleaned.replace(/<[^>]*>/g, '');
    
    // Clean up extra spaces and newlines
    cleaned = cleaned.replace(/\s+/g, ' ').trim();
    
    return cleaned;
  };

  const getCellValue = (cell) => {
    // Handle rich text objects
    if (cell.value && typeof cell.value === 'object') {
      if (cell.value.richText) {
        // Rich text format
        return cell.value.richText.map(rt => rt.text).join('');
      } else if (cell.value.text) {
        return cell.value.text;
      }
    }
    return cell.value;
  };

  const setCellValue = (cell, newValue) => {
    // Preserve cell type and set new value
    if (cell.value && typeof cell.value === 'object') {
      if (cell.value.richText) {
        // Set as plain text, not rich text
        cell.value = newValue;
      } else if (cell.value.text) {
        cell.value.text = newValue;
      } else {
        cell.value = newValue;
      }
    } else {
      cell.value = newValue;
    }
  };

  const handleFileChange = (e) => {
    const selectedFile = e.target.files[0];
    if (selectedFile) {
      setFile(selectedFile);
      setStatus('');
    }
  };

  const handleDrag = (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      setFile(e.dataTransfer.files[0]);
      setStatus('');
    }
  };

  const processExcel = async () => {
    if (!file) {
      setStatus('Lütfen bir dosya seçin!');
      return;
    }

    setProcessing(true);
    setStatus('İşleniyor...');

    try {
      // Read Excel file with ExcelJS
      const workbook = new ExcelJS.Workbook();
      const arrayBuffer = await file.arrayBuffer();
      await workbook.xlsx.load(arrayBuffer);
      
      let totalProcessed = 0;
      
      // Process each worksheet
      workbook.eachSheet((worksheet) => {
        // Try to find columns by header names first
        let basvuruCol = null;
        let metinCol = null;
        
        // Check first row for headers
        const headerRow = worksheet.getRow(1);
        
        // Check specific columns and also scan all columns
        for (let colNumber = 1; colNumber <= 20; colNumber++) {
          const cell = headerRow.getCell(colNumber);
          const headerText = String(cell.value || '').trim();
          if (headerText) {
            if (headerText.includes('Başvuru') && headerText.includes('İçeriği')) {
              basvuruCol = colNumber;
            } else if (headerText === 'Metin') {
              metinCol = colNumber;
            }
          }
        }
        
        // If not found by header, use H and J columns (8 and 10)
        if (!basvuruCol) {
          basvuruCol = 8; // H column
        }
        if (!metinCol) {
          metinCol = 10; // J column
        }
        
        // Process rows starting from row 2
        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          if (rowNumber === 1) return; // Skip header row
          
          // Process "Başvuru İçeriği" column
          if (basvuruCol) {
            const cell = row.getCell(basvuruCol);
            const cellValue = getCellValue(cell);
            
            if (cellValue && typeof cellValue === 'string') {
              const original = cellValue;
              const cleaned = removeHtmlTags(original);
              if (cleaned !== original) {
                setCellValue(cell, cleaned);
                totalProcessed++;
              }
            }
          }
          
          // Process "Metin" column
          if (metinCol) {
            const cell = row.getCell(metinCol);
            const cellValue = getCellValue(cell);
            
            if (cellValue && typeof cellValue === 'string') {
              const original = cellValue;
              const cleaned = removeHtmlTags(original);
              if (cleaned !== original) {
                setCellValue(cell, cleaned);
                totalProcessed++;
              }
            }
          }
        });
      });
      
      // Generate Excel file as buffer
      const buffer = await workbook.xlsx.writeBuffer();
      
      // Create blob and download
      const blob = new Blob([buffer], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      
      const fileName = file.name.replace(/\.[^/.]+$/, '_temizlenmis.xlsx');
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);

      setStatus(`✓ Başarılı! ${totalProcessed} hücre temizlendi. Format korundu.`);
      setProcessing(false);
    } catch (error) {
      console.error('Hata detayı:', error);
      setStatus('Hata: ' + error.message);
      setProcessing(false);
    }
  };

  return (
    <div className="App">
      <header className="App-header">
        <img src={logo} className="App-logo" alt="logo" />
        <h1>Excel HTML Tag Temizleyici</h1>
        <p className="subtitle">CİMER</p>
      </header>

      <main className="App-main">
        <div className="container">
          <div className="instructions">
            <h2>Nasıl Kullanılır?</h2>
            <ol>
              <li>Excel dosyanızı seçin veya sürükleyip bırakın</li>
              <li>"Başvuru İçeriği" ve "Metin" kolonlarındaki HTML etiketleri temizlenecek</li>
              <li>Orijinal format korunarak temizlenmiş dosya indirilecek</li>
            </ol>
          </div>

          <div 
            className={`upload-area ${dragActive ? 'drag-active' : ''}`}
            onDragEnter={handleDrag}
            onDragLeave={handleDrag}
            onDragOver={handleDrag}
            onDrop={handleDrop}
          >
            <div className="upload-icon">📄</div>
            <label htmlFor="file-upload" className="file-label">
              {file ? (
                <span className="file-name">{file.name}</span>
              ) : (
                <span>Dosya Seç veya Sürükle</span>
              )}
            </label>
            <input
              id="file-upload"
              type="file"
              accept=".xlsx,.xls"
              onChange={handleFileChange}
              className="file-input"
            />
            <p className="file-info">Excel dosyaları (.xlsx, .xls)</p>
          </div>

          {status && (
            <div className={`status ${status.includes('✓') ? 'success' : status.includes('Hata') ? 'error' : 'info'}`}>
              {status}
            </div>
          )}

          <button 
            className="process-button" 
            onClick={processExcel}
            disabled={!file || processing}
          >
            {processing ? (
              <>
                <span className="spinner"></span>
                İşleniyor...
              </>
            ) : (
              'HTML Etiketlerini Temizle'
            )}
          </button>
        </div>
      </main>

      <footer className="App-footer">
        <p>T.C. Cumhurbaşkanlığı İletişim Başkanlığı</p>
      </footer>
    </div>
  );
}

export default App;
