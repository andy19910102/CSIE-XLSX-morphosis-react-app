"use client";
import Image from "next/image";
import { useState, useCallback } from "react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";

export default function Home() {
  const [isDragging, setIsDragging] = useState(false);
  const [processedData, setProcessedData] = useState(null);

  const handleDrag = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
  }, []);

  const handleDragIn = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.dataTransfer.items && e.dataTransfer.items.length > 0) {
      setIsDragging(true);
    }
  }, []);

  const handleDragOut = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
  }, []);

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);

    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const file = e.dataTransfer.files[0];
      processFile(file);
      e.dataTransfer.clearData();
    }
  }, []);

  const handleFileSelect = useCallback((e) => {
    if (e.target.files && e.target.files.length > 0) {
      const file = e.target.files[0];
      processFile(file);
    }
  }, []);

  const processFile = async (file) => {
    if (file.type === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const journalSheet = workbook.Sheets["日記帳"];
        if (!journalSheet) {
          alert("Could not find '日記帳' worksheet in the uploaded file");
          return;
        }
        // Parse with date formatting
        // Normalize header names by trimming whitespace
        const range = XLSX.utils.decode_range(journalSheet['!ref']);
        const headers = {};
        for (let C = range.s.c; C <= range.e.c; ++C) {
          const cell = journalSheet[XLSX.utils.encode_cell({r: range.s.r, c: C})];
          if (!cell) continue;
          let header = cell.v.toString().trim();
          headers[XLSX.utils.encode_col(C)] = header;
        }

        // Apply normalized headers
        for (let C = range.s.c; C <= range.e.c; ++C) {
          const col = XLSX.utils.encode_col(C);
          if (headers[col]) {
            const firstCell = XLSX.utils.encode_cell({r: range.s.r, c: C});
            journalSheet[firstCell].v = headers[col];
          }
        }

        const jsonData = XLSX.utils.sheet_to_json(journalSheet, {
          raw: false,
          dateNF: 'm/d',
          defval: ''
        });

        // Normalize data values by trimming whitespace
        const normalizedData = jsonData.map(row => ({
          ...row,
          '收入': typeof row['收入'] === 'string' ? row['收入'].trim() : row['收入'],
          '支出': typeof row['支出'] === 'string' ? row['支出'].trim() : row['支出']
        }));

        console.log('Parsed XLSX Data:', normalizedData);

        // Group data by 科代, excluding empty account codes
        const groupedData = {};
        jsonData.forEach(row => {
          const accountCode = row['科代']?.toString();
          if (accountCode && accountCode.trim() !== '') {
            // Normalize column names and their values
            const normalizedRow = {};
            Object.entries(row).forEach(([key, value]) => {
              // Remove extra spaces from column names
              const normalizedKey = key.trim();
              // Handle special cases for income and expense columns
              if (normalizedKey === '支出' || normalizedKey === ' 支出 ') {
                normalizedRow['支出'] = value?.toString().trim() || '';
              } else if (normalizedKey === '收入' || normalizedKey === ' 收入 ') {
                normalizedRow['收入'] = value?.toString().trim() || '';
              } else {
                normalizedRow[normalizedKey] = value;
              }
            });

            if (!groupedData[accountCode]) {
              groupedData[accountCode] = [];
            }
            groupedData[accountCode].push(normalizedRow);
          }
        });

        console.log('Grouped Data:', groupedData);
        setProcessedData(groupedData);
      };
      reader.readAsArrayBuffer(file);
    } else {
      alert("Please upload an XLSX file");
    }
  };

  const [showModal, setShowModal] = useState(false);
  const [selectedAccounts, setSelectedAccounts] = useState({});

  const handleDownload = () => {
    if (!processedData) return;

    const workbook = XLSX.utils.book_new();

    // Only process selected accounts
    Object.entries(processedData).forEach(([accountCode, data]) => {
      if (selectedAccounts[accountCode]) {
        // Add account code and name as first row
        const accountName = data[0]['會計科目'] || '';
        const accountInfoRow = [{
          '日期': accountCode,
          '經費名稱': accountName,
          '經辦人': '',
          '粘存單': '',
          '摘要': '',
          '受款人': '',
          '收入': '',
          '支出': ''
        }];

        // Add header row with column names
        const headerRow = [{
          '日期': '日期',
          '經費名稱': '經費名稱',
          '經辦人': '經辦人',
          '粘存單': '粘存單',
          '摘要': '摘要',
          '受款人': '受款人',
          '收入': '收入',
          '支出': '支出'
        }];

        // Reorder and ensure all columns exist
        const orderedData = accountInfoRow.concat(headerRow).concat(data.map(row => ({
          '日期': row['日期'] || '',
          '經費名稱': row['經費名稱'] || '',
          '經辦人': row['經辦人'] || '',
          '粘存單': row['粘存單'] || '',
          '摘要': row['摘要'] || '',
          '受款人': row['受款人'] || '',
          '收入': row['收入'] || '',
          '支出': row['支出'] || ''
        })));
        const worksheet = XLSX.utils.json_to_sheet(orderedData, { skipHeader: true });
        // Set date format for the date column
        const dateCol = XLSX.utils.decode_col('A');
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        for (let row = range.s.r + 1; row <= range.e.r; ++row) {
          const cell = worksheet[XLSX.utils.encode_cell({r: row, c: dateCol})];
          if (cell && cell.t === 'd') {
            cell.z = 'm/d';
          }
        }
        XLSX.utils.book_append_sheet(workbook, worksheet, accountCode);
      }
    });

    // Generate XLSX file
    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const data = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(data, 'processed_accounts.xlsx');
    setShowModal(false);
  };

  const toggleAccount = (accountCode) => {
    setSelectedAccounts(prev => ({
      ...prev,
      [accountCode]: !prev[accountCode]
    }));
  };

  const selectAll = () => {
    const allAccounts = {};
    Object.keys(processedData).forEach(code => {
      allAccounts[code] = true;
    });
    setSelectedAccounts(allAccounts);
  };

  const deselectAll = () => {
    setSelectedAccounts({});
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-gray-50 to-gray-100 p-4">
      <div
        className={`w-full max-w-2xl p-8 ${isDragging ? 'bg-blue-50' : 'bg-white'} rounded-xl shadow-lg transition-all duration-300 ease-in-out`}
      >
        <div
          className={`relative border-2 border-dashed rounded-lg p-12 ${isDragging ? 'border-blue-400 bg-blue-50' : 'border-gray-300 hover:border-blue-400'} transition-colors duration-300 ease-in-out`}
          onDragEnter={handleDragIn}
          onDragLeave={handleDragOut}
          onDragOver={handleDrag}
          onDrop={handleDrop}
          onClick={() => document.getElementById('fileInput').click()}
        >
          <input
            type="file"
            id="fileInput"
            className="hidden"
            accept=".xlsx"
            onChange={handleFileSelect}
          />
          <div className="text-center">
            <Image
              src="/excel-png-office-xlsx-icon-3.png"
              alt="Upload XLSX"
              width={64}
              height={64}
              className="mx-auto mb-4"
            />
            <h2 className="text-xl font-semibold text-gray-700 mb-2">
              將您的 XLSX 檔案拖曳至此
            </h2>
            <p className="text-sm text-gray-500">
              或點擊此處選擇檔案
            </p>
          </div>
        </div>
        {processedData && (
          <button
            onClick={() => {
              const allAccounts = {};
              Object.keys(processedData).forEach(code => {
                allAccounts[code] = true;
              });
              setSelectedAccounts(allAccounts);
              setShowModal(true);
            }}
            className="mt-4 px-4 py-2 bg-green-500 text-white rounded hover:bg-green-600 transition-colors duration-300 w-full"
          >
            選擇要下載的科目
          </button>
        )}
      </div>
      {showModal && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-lg p-6 max-w-md w-full">
            <h3 className="text-lg font-semibold mb-4">選擇要下載的科目</h3>
            <div className="mb-4 flex gap-2">
              <button
                onClick={selectAll}
                className="px-3 py-1 bg-blue-500 text-white rounded text-sm hover:bg-blue-600"
              >
                全選
              </button>
              <button
                onClick={deselectAll}
                className="px-3 py-1 bg-gray-500 text-white rounded text-sm hover:bg-gray-600"
              >
                取消全選
              </button>
            </div>
            <div className="max-h-60 overflow-y-auto">
              {Object.keys(processedData).map(accountCode => (
                <label key={accountCode} className="flex items-center p-2 hover:bg-gray-100 rounded">
                  <input
                    type="checkbox"
                    checked={!!selectedAccounts[accountCode]}
                    onChange={() => toggleAccount(accountCode)}
                    className="mr-2"
                  />
                  {accountCode}
                </label>
              ))}
            </div>
            <div className="mt-4 flex justify-end gap-2">
              <button
                onClick={() => setShowModal(false)}
                className="px-4 py-2 bg-gray-300 text-gray-700 rounded hover:bg-gray-400"
              >
                取消
              </button>
              <button
                onClick={handleDownload}
                className="px-4 py-2 bg-green-500 text-white rounded hover:bg-green-600"
                disabled={Object.values(selectedAccounts).filter(Boolean).length === 0}
              >
                下載選擇的科目
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );

}
