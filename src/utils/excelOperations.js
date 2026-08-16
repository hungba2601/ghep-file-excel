import ExcelJS from 'exceljs';
import JSZip from 'jszip';

export const readFileAsync = (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(e.target.result);
    reader.onerror = (e) => reject(e);
    reader.readAsArrayBuffer(file);
  });
};

export const cleanExcelBuffer = async (buffer) => {
  try {
    const zip = await JSZip.loadAsync(buffer);
    // Xóa các file metadata thường gây lỗi cho exceljs khi xuất từ các phần mềm thứ ba
    zip.remove('docProps/app.xml');
    zip.remove('docProps/core.xml');
    zip.remove('docProps/custom.xml');
    
    // Xử lý tiền tố namespace lạ (ví dụ: <x:sst>) làm sập parser của exceljs
    const filesToProcess = Object.keys(zip.files).filter(name => name.endsWith('.xml') || name.endsWith('.rels'));
    for (const fileName of filesToProcess) {
      let content = await zip.file(fileName).async('string');
      if (content.includes('<x:') || content.includes('</x:')) {
        content = content.replace(/<x:/g, '<').replace(/<\/x:/g, '</');
        zip.file(fileName, content);
      }
    }
    
    const cleanBuffer = await zip.generateAsync({ type: 'arraybuffer' });
    return cleanBuffer;
  } catch (error) {
    console.warn("Could not clean zip buffer, proceeding with original", error);
    return buffer;
  }
};

const copySheet = (sourceSheet, targetSheet, skipRows = 0, targetStartRow = 1) => {
  let sourceRowStart = 1 + skipRows;
  let sourceRowEnd = sourceSheet.rowCount;

  if (targetStartRow === 1 && sourceSheet.columns) {
    // Copy column properties only if this is the first time writing to target sheet
    targetSheet.columns = sourceSheet.columns.map(col => {
      if (!col) return {};
      return {
        width: col.width,
        hidden: col.hidden,
        style: col.style
      };
    });
  }

  let targetCurrentRow = targetStartRow;

  for (let r = sourceRowStart; r <= sourceRowEnd; r++) {
    const sourceRow = sourceSheet.getRow(r);
    const targetRow = targetSheet.getRow(targetCurrentRow);

    if (sourceRow.height) targetRow.height = sourceRow.height;
    if (sourceRow.hidden) targetRow.hidden = sourceRow.hidden;

    sourceRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const targetCell = targetRow.getCell(colNumber);
      targetCell.value = cell.value;
      targetCell.style = cell.style;
    });

    targetCurrentRow++;
  }

  // Handle merges
  (sourceSheet.model?.merges || []).forEach(mergeRange => {
    const match = mergeRange.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (match) {
        const startCol = match[1];
        const startRow = parseInt(match[2], 10);
        const endCol = match[3];
        const endRow = parseInt(match[4], 10);
        
        // Only apply merge if it falls within the rows we copied
        if (startRow >= sourceRowStart) {
            const rowOffset = targetStartRow - sourceRowStart;
            const newStartRow = startRow + rowOffset;
            const newEndRow = endRow + rowOffset;
            try {
              targetSheet.mergeCells(`${startCol}${newStartRow}:${endCol}${newEndRow}`);
            } catch (err) {
              // Ignore merge overlaps
            }
        }
    }
  });

  return targetCurrentRow; // Return next available row
};

export const mergeMultipleFilesToMultipleSheets = async (files) => {
  const wbTarget = new ExcelJS.Workbook();
  
  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    const data = await readFileAsync(file);
    const cleanData = await cleanExcelBuffer(data);
    const wbSource = new ExcelJS.Workbook();
    await wbSource.xlsx.load(cleanData);
    
    wbSource.eachSheet((sourceSheet, sheetId) => {
      let newSheetName = sourceSheet.name;
      let counter = 1;
      
      let safeSheetName = newSheetName.substring(0, 31);
      
      while (wbTarget.getWorksheet(safeSheetName)) {
        const suffix = `(${counter})`;
        safeSheetName = `${newSheetName.substring(0, 31 - suffix.length)}${suffix}`;
        counter++;
      }
      
      const targetSheet = wbTarget.addWorksheet(safeSheetName);
      copySheet(sourceSheet, targetSheet, 0, 1);
    });
  }
  return wbTarget;
};

export const mergeMultipleFilesToOneSheet = async (files, headerRows = 0) => {
  const wbTarget = new ExcelJS.Workbook();
  const targetSheet = wbTarget.addWorksheet('MergedData');
  
  let targetStartRow = 1;

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    const data = await readFileAsync(file);
    const cleanData = await cleanExcelBuffer(data);
    const wbSource = new ExcelJS.Workbook();
    await wbSource.xlsx.load(cleanData);
    
    wbSource.eachSheet((sourceSheet, sheetId) => {
      let skipRows = (targetStartRow === 1) ? 0 : headerRows;
      targetStartRow = copySheet(sourceSheet, targetSheet, skipRows, targetStartRow);
    });
  }
  return wbTarget;
};

export const mergeSingleFileSheetsToOneSheet = async (file, headerRows = 0) => {
  const wbTarget = new ExcelJS.Workbook();
  const targetSheet = wbTarget.addWorksheet('MergedData');
  
  let targetStartRow = 1;

  const data = await readFileAsync(file);
  const cleanData = await cleanExcelBuffer(data);
  const wbSource = new ExcelJS.Workbook();
  await wbSource.xlsx.load(cleanData);
  
  wbSource.eachSheet((sourceSheet, sheetId) => {
    let skipRows = (targetStartRow === 1) ? 0 : headerRows;
    targetStartRow = copySheet(sourceSheet, targetSheet, skipRows, targetStartRow);
  });
  
  return wbTarget;
};

export const downloadWorkbook = async (wb, filename) => {
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  window.URL.revokeObjectURL(url);
};
