import React, { useState, useCallback, useRef } from 'react';
import { UploadCloud, FileSpreadsheet, X, Loader2, Files, Minimize, FileDown, Settings } from 'lucide-react';
import './App.css';
import {
  mergeMultipleFilesToMultipleSheets,
  mergeMultipleFilesToOneSheet,
  mergeSingleFileSheetsToOneSheet,
  downloadWorkbook
} from './utils/excelOperations';

function App() {
  const [files, setFiles] = useState([]);
  const [isDragActive, setIsDragActive] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [headerRows, setHeaderRows] = useState(7); // Default to 7 based on user's screenshot
  const fileInputRef = useRef(null);

  const handleDragEnter = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragActive(true);
  };

  const handleDragLeave = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragActive(false);
  };

  const handleDragOver = (e) => {
    e.preventDefault();
    e.stopPropagation();
  };

  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragActive(false);
    
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFiles(Array.from(e.dataTransfer.files));
    }
  };

  const handleFileInput = (e) => {
    if (e.target.files && e.target.files.length > 0) {
      handleFiles(Array.from(e.target.files));
    }
    // reset input so same file can be selected again if removed
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const handleFiles = (newFiles) => {
    // Filter to only accept excel files (xls, xlsx, csv)
    const validFiles = newFiles.filter(file => 
      file.name.match(/\.(xlsx|xls|csv)$/i)
    );
    
    if (validFiles.length !== newFiles.length) {
      alert("Chỉ chấp nhận các file Excel (.xlsx, .xls, .csv)");
    }

    setFiles(prev => [...prev, ...validFiles]);
  };

  const removeFile = (indexToRemove) => {
    setFiles(files.filter((_, index) => index !== indexToRemove));
  };

  const handleAction = async (actionType) => {
    if (files.length === 0) return;
    
    if (actionType === 'single-multi-to-one' && files.length !== 1) {
      alert("Chức năng này chỉ áp dụng cho 1 file duy nhất.");
      return;
    }
    if ((actionType === 'multi-to-multi' || actionType === 'multi-to-one') && files.length < 2) {
      alert("Vui lòng chọn ít nhất 2 file để sử dụng chức năng này.");
      return;
    }

    setIsLoading(true);

    try {
      let wb;
      let filename = 'Merged_Excel.xlsx';
      
      const skipRows = parseInt(headerRows) || 0;

      switch (actionType) {
        case 'multi-to-multi':
          wb = await mergeMultipleFilesToMultipleSheets(files);
          filename = 'Merged_Multiple_Sheets.xlsx';
          break;
        case 'multi-to-one':
          wb = await mergeMultipleFilesToOneSheet(files, skipRows);
          filename = 'Merged_Single_Sheet.xlsx';
          break;
        case 'single-multi-to-one':
          wb = await mergeSingleFileSheetsToOneSheet(files[0], skipRows);
          filename = `Merged_${files[0].name.replace(/\.[^/.]+$/, "")}_SingleSheet.xlsx`;
          break;
        default:
          return;
      }

      await downloadWorkbook(wb, filename);
    } catch (error) {
      console.error(error);
      alert(`Đã xảy ra lỗi khi xử lý file: ${error.message}\n\nVui lòng đảm bảo các file đều là định dạng .xlsx (exceljs không hỗ trợ file .xls cũ).`);
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className="app-container">
      {isLoading && (
        <div className="loading-overlay">
          <Loader2 size={48} className="spinner" />
          <h2>Đang xử lý file, vui lòng đợi...</h2>
          <p style={{ marginTop: '0.5rem', color: '#6B7280' }}>(Bảo toàn định dạng sẽ mất chút thời gian)</p>
        </div>
      )}

      <div className="header">
        <h1><FileSpreadsheet size={32} /> Excel Merger</h1>
        <p>Công cụ ghép file Excel nhanh chóng, giữ nguyên định dạng (in đậm, gộp ô).</p>
      </div>

      <div 
        className={`dropzone ${isDragActive ? 'drag-active' : ''}`}
        onDragEnter={handleDragEnter}
        onDragLeave={handleDragLeave}
        onDragOver={handleDragOver}
        onDrop={handleDrop}
      >
        <UploadCloud size={48} className="dropzone-icon" />
        <h3 className="dropzone-text">Kéo thả các file Excel của bạn vào đây</h3>
        <p className="dropzone-subtext">hoặc click để chọn file từ máy tính (.xlsx, .xls)</p>
        <input 
          type="file" 
          multiple 
          accept=".xlsx, .xls, .csv" 
          className="file-input"
          onChange={handleFileInput}
          ref={fileInputRef}
        />
      </div>

      {files.length > 0 && (
        <div className="file-list">
          <div className="file-list-title">
            Đã chọn {files.length} file
            <button className="file-list-clear" onClick={() => setFiles([])}>Xoá tất cả</button>
          </div>
          <div className="file-items">
            {files.map((file, index) => (
              <div key={`${file.name}-${index}`} className="file-item">
                <div className="file-item-name" title={file.name}>
                  <FileSpreadsheet size={16} color="#10B981" />
                  {file.name}
                </div>
                <button 
                  className="file-item-remove" 
                  onClick={() => removeFile(index)}
                  title="Xoá file này"
                >
                  <X size={16} />
                </button>
              </div>
            ))}
          </div>
        </div>
      )}

      <div className="settings-section">
        <div className="settings-card">
          <div className="settings-title">
            <Settings size={20} />
            <span>Cấu hình dòng tiêu đề (Dành cho gộp thành 1 sheet)</span>
          </div>
          <div className="settings-control">
            <label htmlFor="headerRows">Số dòng tiêu đề cần giữ (bỏ qua ở các file sau):</label>
            <input 
              type="number" 
              id="headerRows" 
              min="0" 
              max="100" 
              value={headerRows} 
              onChange={(e) => setHeaderRows(e.target.value)}
              className="number-input"
            />
          </div>
        </div>
      </div>

      <div className="actions-section">
        <h2 className="actions-title">Chọn chức năng ghép file</h2>
        
        <div className="action-cards">
          <button 
            className="action-card color-indigo" 
            disabled={files.length < 2}
            onClick={() => handleAction('multi-to-multi')}
          >
            <div className="action-icon-wrapper">
              <Files size={24} />
            </div>
            <h3>Nhiều file ➔ 1 file<br/>(Nhiều Sheet)</h3>
            <p>Gộp nhiều file Excel lại với nhau. Mỗi file hoặc mỗi sheet sẽ là một sheet riêng biệt trong file mới.</p>
            <div className="action-btn">Thực hiện</div>
          </button>

          <button 
            className="action-card color-emerald"
            disabled={files.length < 2}
            onClick={() => handleAction('multi-to-one')}
          >
            <div className="action-icon-wrapper">
              <Minimize size={24} />
            </div>
            <h3>Nhiều file ➔ 1 file<br/>(1 Sheet duy nhất)</h3>
            <p>Gộp toàn bộ dữ liệu từ tất cả các file và tất cả các sheet thành 1 sheet duy nhất.</p>
            <div className="action-btn">Thực hiện</div>
          </button>

          <button 
            className="action-card color-rose"
            disabled={files.length !== 1}
            onClick={() => handleAction('single-multi-to-one')}
          >
            <div className="action-icon-wrapper">
              <FileDown size={24} />
            </div>
            <h3>1 file (nhiều Sheet) ➔ 1 Sheet</h3>
            <p>Gộp toàn bộ các sheet có trong 1 file Excel thành 1 sheet duy nhất.</p>
            <div className="action-btn">Thực hiện</div>
          </button>
        </div>
      </div>
      
      <div className="copyright-footer">
        Made by Nguyễn Phi Hùng - Zalo 0938750424
      </div>
    </div>
  );
}

export default App;
