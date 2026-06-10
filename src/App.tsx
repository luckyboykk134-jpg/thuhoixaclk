import React, { useState, useEffect, useRef, useMemo, useCallback } from 'react';
import { Search, CheckCircle, XCircle, AlertTriangle, FileSpreadsheet, Trash2, Settings, BarChart2, ScanLine, Ban, ClipboardList, X, Download, Save, UploadCloud, Loader2, ExternalLink, RefreshCw } from 'lucide-react';

// ----------------------------------------------------------------------
// HELPER FUNCTIONS (Pure Functions)
// ----------------------------------------------------------------------
const getRemark = (row) => {
  if (!row) return "";
  let remarks = [];
  if (String(row["KMH"]) === "1") remarks.push("KMH");
  if (String(row["Demo"]) === "1") remarks.push("Demo");
  return remarks.join(", ");
};

const createRecord = (rawScan, status, rowData = null, extraData = {}) => {
  return {
    id: Date.now() + Math.random(),
    rawScan: String(rawScan || ''),
    status: String(status || ''),
    sp: rowData ? String(rowData["SP"] || "").trim() : '', 
    scCode: rowData ? String(rowData["SC Code"] || rowData["SC code"] || "").trim() : '', 
    ttbh: rowData ? String(rowData["Warehouse Name"] || "").trim() : '',
    soRO: rowData ? String(rowData["After-sales work order No."] || "").trim() : '',
    maLK: rowData ? String(rowData["Defective material code"] || "").trim() : '',
    tenLK: rowData ? String(rowData["Product Name"] || "").trim() : '',
    model: rowData ? String(rowData["Product Model"] || rowData["Model"] || "").trim() : '', 
    phanLoai: rowData ? String(rowData["Type"] || "").trim() : '',
    bhDv: rowData ? String(rowData["Repair Type"] || "").trim() : '',
    slg: rowData ? String(rowData["Consumed quantity"] || rowData["Consumed"] || "").trim() : '', 
    remark: String(getRemark(rowData) || ""),
    ...extraData
  };
};

const renderProgressBar = (scannedCount, totalCount, label, type) => {
  const percent = totalCount > 0 ? Math.round((scannedCount / totalCount) * 100) : 0;
  const isIW = type === 'IW';
  const colorBg = isIW ? 'bg-emerald-100' : 'bg-blue-100';
  const colorFg = isIW ? 'bg-emerald-500' : 'bg-blue-500';
  const colorText = isIW ? 'text-emerald-700' : 'text-blue-700';

  return (
    <div className="w-full mb-3 last:mb-0">
      <div className="flex justify-between items-end mb-1">
        <span className={`font-bold text-sm uppercase tracking-wider ${colorText}`}>{String(label)}</span>
        <span className={`font-bold text-sm ${colorText}`}>
          {percent}% <span className="text-xs opacity-70 font-medium ml-0.5">({Number(scannedCount)}/{Number(totalCount)})</span>
        </span>
      </div>
      <div className={`w-full h-3 rounded-full ${colorBg} overflow-hidden shadow-inner`}>
        <div className={`h-full rounded-full ${colorFg} transition-all duration-700 ease-out`} style={{ width: `${percent}%` }}></div>
      </div>
    </div>
  );
};

const generateWorkbookData = (selectedWarehouse, currentDisplayedRecords, detailedProgressList) => {
  const workbook = window.XLSX.utils.book_new();

  const scannedKMH_SPs = new Set();
  currentDisplayedRecords.forEach(record => {
      if ((record.status === "Khớp, Trả Xác LK về" || record.status === "Không xác LK") && String(record.remark).includes("KMH")) {
          scannedKMH_SPs.add(String(record.sp || '').trim().toUpperCase());
      }
  });

  const isKMHChecked = (row) => {
      const sp = String(row["SP"] || '').trim().toUpperCase();
      return String(row["KMH"]) === "1" || String(row["KMH"]).includes("KMH") || scannedKMH_SPs.has(sp);
  };

  const statsData = [
    {
      "Nhóm Phân Loại": "Bảo Hành (IW)",
      "Tổng Cần Scan": detailedProgressList.filter(r => String(r["Repair Type"]).toUpperCase() === 'IW').length,
      "Đã Scan Khớp": detailedProgressList.filter(r => String(r["Repair Type"]).toUpperCase() === 'IW' && r.isScanned).length,
      "Còn Thiếu": detailedProgressList.filter(r => String(r["Repair Type"]).toUpperCase() === 'IW' && !r.isScanned).length,
    },
    {
      "Nhóm Phân Loại": "Dịch Vụ (OOW)",
      "Tổng Cần Scan": detailedProgressList.filter(r => String(r["Repair Type"]).toUpperCase() === 'OOW' && !isKMHChecked(r)).length,
      "Đã Scan Khớp": detailedProgressList.filter(r => String(r["Repair Type"]).toUpperCase() === 'OOW' && r.isScanned && !isKMHChecked(r)).length,
      "Còn Thiếu": detailedProgressList.filter(r => String(r["Repair Type"]).toUpperCase() === 'OOW' && !r.isScanned && !isKMHChecked(r)).length,
    },
    {
      "Nhóm Phân Loại": "Khác",
      "Tổng Cần Scan": detailedProgressList.filter(r => !['IW', 'OOW'].includes(String(r["Repair Type"]).toUpperCase())).length,
      "Đã Scan Khớp": detailedProgressList.filter(r => !['IW', 'OOW'].includes(String(r["Repair Type"]).toUpperCase()) && r.isScanned).length,
      "Còn Thiếu": detailedProgressList.filter(r => !['IW', 'OOW'].includes(String(r["Repair Type"]).toUpperCase()) && !r.isScanned).length,
    }
  ];
  const wsStats = window.XLSX.utils.json_to_sheet(statsData);
  wsStats['!cols'] = [{ wch: 20 }, { wch: 15 }, { wch: 15 }, { wch: 20 }];
  window.XLSX.utils.book_append_sheet(workbook, wsStats, "ThongKe");

  const exportScannedData = currentDisplayedRecords.map(record => ({
    "Trạng thái": String(record.status), 
    "Cột SP": String(record.sp || record.rawScan), 
    "SC Code": String(record.scCode), 
    "Warehouse Name": String(record.ttbh), 
    "Số RO": String(record.soRO), 
    "BH/DV": String(record.bhDv), 
    "Mã LK": String(record.maLK), 
    "Product Name": String(record.tenLK), 
    "Model": String(record.model),
    "Type": String(record.phanLoai), 
    "Slg": String(record.slg), 
    "Remark": String(record.remark)
  }));
  const wsScanned = window.XLSX.utils.json_to_sheet(exportScannedData);
  wsScanned['!cols'] = [{ wch: 25 }, { wch: 45 }, { wch: 15 }, { wch: 30 }, { wch: 25 }, { wch: 10 }, { wch: 15 }, { wch: 40 }, { wch: 20 }, { wch: 10 }, { wch: 10 }, { wch: 15 }];
  window.XLSX.utils.book_append_sheet(workbook, wsScanned, "LichSu_DaScan");

  const unscannedData = detailedProgressList.filter(r => !r.isScanned).map(record => ({
    "Trạng thái": "Chưa Scan", 
    "Cột SP": String(record["SP"] || ""), 
    "SC Code": String(record["SC Code"] || record["SC code"] || ''), 
    "Warehouse Name": String(record["Warehouse Name"] || ""), 
    "Số RO": String(record["After-sales work order No."] || ""), 
    "BH/DV": String(record["Repair Type"] || ""), 
    "Mã LK": String(record["Defective material code"] || ""), 
    "Product Name": String(record["Product Name"] || ""), 
    "Model": String(record["Product Model"] || record["Model"] || ''), 
    "Type": String(record["Type"] || ""), 
    "Slg": String(record["Consumed quantity"] || record["Consumed"] || ''), 
    "Remark": String(getRemark(record))
  }));
  const wsUnscanned = window.XLSX.utils.json_to_sheet(unscannedData);
  wsUnscanned['!cols'] = [{ wch: 15 }, { wch: 45 }, { wch: 15 }, { wch: 30 }, { wch: 25 }, { wch: 10 }, { wch: 15 }, { wch: 40 }, { wch: 20 }, { wch: 10 }, { wch: 10 }, { wch: 15 }];
  window.XLSX.utils.book_append_sheet(workbook, wsUnscanned, "DanhSach_ChuaScan");

  if (!workbook.Workbook) workbook.Workbook = {};
  workbook.Workbook.Sheets = [{ Hidden: 1 }, { Hidden: 0 }, { Hidden: 0 }];

  return workbook;
};

// ----------------------------------------------------------------------
// MAIN COMPONENT
// ----------------------------------------------------------------------
export default function App() {
  const [excelData, setExcelData] = useState([]);
  const [warehouses, setWarehouses] = useState([]);
  const [selectedWarehouse, setSelectedWarehouse] = useState('');
  const [scanInput, setScanInput] = useState('');
  const [scannedRecords, setScannedRecords] = useState([]);
  const [isLoadingData, setIsLoadingData] = useState(false);
  
  const [isWhDropdownOpen, setIsWhDropdownOpen] = useState(false);
  const [whSearchQuery, setWhSearchQuery] = useState('');
  const whDropdownRef = useRef(null);

  const gasUrl = 'https://script.google.com/macros/s/AKfycbyo8B2gUNmdLgcXx5YHrUmcUdRy-5X9sdr8voVKWCcqS2FiIFgNkBwHK9d4x7ozKyFH/exec';
  const [isUploading, setIsUploading] = useState(false);

  const [showPasswordModal, setShowPasswordModal] = useState(false);
  const [passwordInput, setPasswordInput] = useState('');
  const passwordInputRef = useRef(null);

  const [toast, setToast] = useState(null);
  const [confirmDialog, setConfirmDialog] = useState(null);
  const [warningAlert, setWarningAlert] = useState(null); 
  const [pendingQRCheck, setPendingQRCheck] = useState(null); 
  const [secondScanInput, setSecondScanInput] = useState('');
  const [showDetailedProgressModal, setShowDetailedProgressModal] = useState(false);
  
  const [lockedMaLK, setLockedMaLK] = useState(null);
  const [violationCount, setViolationCount] = useState(0);
  const [showForceRemarkModal, setShowForceRemarkModal] = useState(false);
  const [missingRowsForLockedLK, setMissingRowsForLockedLK] = useState([]);
  const [rowRemarks, setRowRemarks] = useState({}); 
  const [selectedMissingRows, setSelectedMissingRows] = useState({}); 
  
  const [showOverrideInput, setShowOverrideInput] = useState(false);
  const [overridePassword, setOverridePassword] = useState('');
  const [isGlobalOverrideActive, setIsGlobalOverrideActive] = useState(false);

  const [quickNoXacRow, setQuickNoXacRow] = useState(null);
  const [quickNoXacReason, setQuickNoXacReason] = useState('');

  const [showOOWOverrideInput, setShowOOWOverrideInput] = useState(false);
  const [oowOverridePassword, setOowOverridePassword] = useState('');
  const [isOOWOverrideActive, setIsOOWOverrideActive] = useState(false);
  const [oowScanInput, setOowScanInput] = useState('');

  const mainInputRef = useRef(null);
  const secondInputRef = useRef(null);
  const oowScanInputRef = useRef(null);

  const hasErrorLock = useMemo(() => {
    return scannedRecords.some(r => String(r.status) !== "Khớp, Trả Xác LK về" && String(r.status) !== "Không xác LK");
  }, [scannedRecords]);

  const currentDisplayedRecords = useMemo(() => {
    return scannedRecords.filter(r => String(r.ttbh) === String(selectedWarehouse) || String(r.ttbh) === '');
  }, [scannedRecords, selectedWarehouse]);

  const conflictRecordsSPs = useMemo(() => {
    const spMap = new Map(); 
    currentDisplayedRecords.forEach(r => {
      const key = String(r.sp || "").trim().toUpperCase();
      if (!key) return;
      if (!spMap.has(key)) {
        spMap.set(key, { hasKhop: false, hasNoXac: false });
      }
      const info = spMap.get(key);
      if (String(r.status) === "Khớp, Trả Xác LK về") info.hasKhop = true;
      if (String(r.status) === "Không xác LK") info.hasNoXac = true;
    });

    const conflicting = new Set();
    spMap.forEach((val, key) => {
      if (val.hasKhop && val.hasNoXac) {
        conflicting.add(key);
      }
    });
    return conflicting;
  }, [currentDisplayedRecords]);

  const hasConflictLock = useMemo(() => {
    return conflictRecordsSPs.size > 0;
  }, [conflictRecordsSPs]);

  const sortedDisplayedRecords = useMemo(() => {
    return [...currentDisplayedRecords].sort((a, b) => {
      const aSp = String(a.sp || "").trim().toUpperCase();
      const bSp = String(b.sp || "").trim().toUpperCase();
      const aIsConflictNoXac = conflictRecordsSPs.has(aSp) && String(a.status) === "Không xác LK";
      const bIsConflictNoXac = conflictRecordsSPs.has(bSp) && String(b.status) === "Không xác LK";
      
      if (aIsConflictNoXac && !bIsConflictNoXac) return -1;
      if (!aIsConflictNoXac && bIsConflictNoXac) return 1;
      return 0; 
    });
  }, [currentDisplayedRecords, conflictRecordsSPs]);

  const currentScCode = useMemo(() => {
    const match = excelData.find(row => String(row["Warehouse Name"]).trim() === selectedWarehouse);
    return match ? String(match["SC Code"] || match["SC code"] || "").trim() : '';
  }, [excelData, selectedWarehouse]);

  const filteredWarehouses = useMemo(() => {
    return warehouses.filter(wh => 
      String(wh).toLowerCase().includes(whSearchQuery.toLowerCase())
    );
  }, [warehouses, whSearchQuery]);

  const detailedProgressList = useMemo(() => {
    if (!selectedWarehouse || excelData.length === 0) return [];
    
    const scannedSPs = scannedRecords
      .filter(r => (String(r.status) === "Khớp, Trả Xác LK về" || String(r.status) === "Không xác LK") && String(r.ttbh) === selectedWarehouse)
      .map(r => String(r.sp || r.rawScan).split('_').join('').split(' ').join('').toUpperCase());
      
    const result = excelData
      .filter(row => String(row["Warehouse Name"]).trim() === selectedWarehouse)
      .map(row => {
        const isScanned = scannedSPs.includes(String(row["SP"]).trim());
        return { ...row, isScanned };
      });
      
    return result.sort((a, b) => (a.isScanned === b.isScanned ? 0 : a.isScanned ? 1 : -1));
  }, [excelData, selectedWarehouse, scannedRecords]);

  const canSkipAnyMissing = useMemo(() => {
    return missingRowsForLockedLK.some(row => {
        const repairType = String(row["Repair Type"] || '').trim().toUpperCase();
        const typeLK = String(row["Type"] || '').trim().toUpperCase();
        return !(repairType === 'IW' && typeLK !== 'PIN');
    });
  }, [missingRowsForLockedLK]);

  useEffect(() => {
    if (isGlobalOverrideActive || !selectedWarehouse || excelData.length === 0 || scannedRecords.length === 0) {
      setLockedMaLK(null);
      setViolationCount(0);
      return;
    }

    const lastValidRecord = scannedRecords.find(r => 
      String(r.ttbh) === String(selectedWarehouse) && 
      (String(r.status) === "Khớp, Trả Xác LK về" || String(r.status) === "Không xác LK")
    );

    if (lastValidRecord) {
      const currentMaLK = String(lastValidRecord.maLK);
      const repairType = String(lastValidRecord.bhDv).trim().toUpperCase();

      if (repairType === 'IW' || repairType === 'OOW') {
         const totalRowsForMaLK = excelData.filter(row => 
           String(row["Warehouse Name"]).trim() === selectedWarehouse && 
           String(row["Defective material code"]).trim() === currentMaLK &&
           String(row["Repair Type"]).trim().toUpperCase() === repairType
         );

         const scannedRowsForMaLK = scannedRecords.filter(r => 
           String(r.ttbh) === String(selectedWarehouse) && 
           String(r.maLK) === currentMaLK && 
           String(r.bhDv).trim().toUpperCase() === repairType &&
           (String(r.status) === "Khớp, Trả Xác LK về" || String(r.status) === "Không xác LK")
         );

         if (scannedRowsForMaLK.length < totalRowsForMaLK.length && totalRowsForMaLK.length > 0) {
           setLockedMaLK(currentMaLK);
           const scannedSPs = scannedRowsForMaLK.map(r => String(r.sp || r.rawScan).split('_').join('').split(' ').join('').toUpperCase());
           const missing = totalRowsForMaLK.filter(row => !scannedSPs.includes(String(row["SP"]).trim()));
           setMissingRowsForLockedLK(missing);
         } else {
           setLockedMaLK(null);
           setViolationCount(0);
         }
      } else {
         setLockedMaLK(null);
         setViolationCount(0);
      }
    } else {
      setLockedMaLK(null);
      setViolationCount(0);
    }
  }, [scannedRecords, selectedWarehouse, excelData, isGlobalOverrideActive]);

  const stats = useMemo(() => {
    let totals = { IW: { LCD: 0, MAIN: 0, OTHERS: 0, total: 0 }, OOW: { LCD: 0, MAIN: 0, OTHERS: 0, total: 0 }, KMH: { LCD: 0, MAIN: 0, OTHERS: 0 } };
    let scanned = { IW: { LCD: 0, MAIN: 0, OTHERS: 0, total: 0 }, OOW: { LCD: 0, MAIN: 0, OTHERS: 0, total: 0 }, KMH: { LCD: 0, MAIN: 0, OTHERS: 0 } };

    if (!selectedWarehouse) return { totals, scanned };

    const processItem = (row, targetObj, qtyStr, repairStr, typeStr, kmhStr) => {
      const qty = parseInt(qtyStr) || 1;
      const rType = repairStr === 'IW' ? 'IW' : (repairStr === 'OOW' ? 'OOW' : null);
      const tType = (typeStr === 'LCD' || typeStr === 'MAIN') ? typeStr : 'OTHERS';
      const isKMH = String(kmhStr) === "1" || String(kmhStr).includes("KMH");
      
      if (isKMH) {
        targetObj.KMH[tType] += qty;
      }
      
      if (rType && targetObj[rType]) {
        if (rType === 'OOW' && isKMH) return;
        targetObj[rType][tType] += qty;
        targetObj[rType].total += qty;
      }
    };

    const scannedKMH_SPs = new Set();
    currentDisplayedRecords.forEach(record => {
       if ((String(record.status) === "Khớp, Trả Xác LK về" || String(record.status) === "Không xác LK") && String(record.remark).includes("KMH")) {
           scannedKMH_SPs.add(String(record.sp || '').trim().toUpperCase());
       }
    });

    excelData.forEach(row => {
      if (String(row["Warehouse Name"]).trim() === selectedWarehouse) {
        const sp = String(row["SP"] || '').trim().toUpperCase();
        let kmhStr = String(row["KMH"] || '');
        if (scannedKMH_SPs.has(sp)) kmhStr = "KMH";
        processItem(row, totals, String(row["Consumed quantity"] || row["Consumed"] || '1'), String(row["Repair Type"] || '').trim().toUpperCase(), String(row["Type"] || '').trim().toUpperCase(), kmhStr);
      }
    });

    currentDisplayedRecords.forEach(record => {
      if (String(record.status) === "Khớp, Trả Xác LK về" || String(record.status) === "Không xác LK") {
        processItem(record, scanned, String(record.slg || '1'), String(record.bhDv || '').trim().toUpperCase(), String(record.phanLoai || '').trim().toUpperCase(), String(record.remark || ''));
      }
    });

    return { totals, scanned };
  }, [excelData, selectedWarehouse, currentDisplayedRecords]);

  const showToast = useCallback((message, type = 'info') => {
    setToast({ message: String(message), type });
    setTimeout(() => setToast(null), 4000);
  }, []);

  const fetchGoogleSheetData = useCallback(async (isBackgroundUpdate = false) => {
    if (!isBackgroundUpdate) setIsLoadingData(true);
    await new Promise(resolve => setTimeout(resolve, 50)); 

    try {
      const sheetId = "10zbQkv7f_7nwVawF5lFDdqfpbC23UT1YBjgJ3aZyNWU";
      const url = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx`;
      const response = await fetch(url);
      if (!response.ok) throw new Error("Network response was not ok");
      const data = await response.arrayBuffer();

      if (!window.XLSX) {
        if (!isBackgroundUpdate) {
          showToast("Thư viện xử lý chưa sẵn sàng.", "warning");
          setIsLoadingData(false);
        }
        return;
      }

      const workbook = window.XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames.includes("DataXacLK") ? "DataXacLK" : workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const rawJsonData = window.XLSX.utils.sheet_to_json(worksheet, { defval: "" });

      if (rawJsonData.length === 0) {
        if (!isBackgroundUpdate) showToast("Dữ liệu trống!", "error");
        if (!isBackgroundUpdate) setIsLoadingData(false);
        return;
      }

      const jsonData = rawJsonData.map(row => {
        const cleanRow = {};
        Object.keys(row).forEach(key => cleanRow[String(key).trim()] = row[key]);
        return {
          ...cleanRow,
          "SP": String(cleanRow["SP"] || "").split(' ').join('').toUpperCase(),
          "Warehouse Name": String(cleanRow["Warehouse Name"] || "").trim()
        };
      });

      setExcelData(jsonData);
      const whList = [...new Set(jsonData.map(item => item["Warehouse Name"]).filter(Boolean))].sort((a, b) => String(a).localeCompare(String(b), 'vi'));
      setWarehouses(whList);
      
      try {
         localStorage.setItem('oppo_excel_data', JSON.stringify(jsonData));
         localStorage.setItem('oppo_warehouses', JSON.stringify(whList));
      } catch (e) {
         console.warn("Không thể lưu cache: ", e);
      }

      setSelectedWarehouse(prev => prev || (whList.length > 0 ? String(whList[0]) : ''));
      
      if (isBackgroundUpdate) {
        showToast("Dữ liệu gốc đã được làm mới ở chế độ nền!", "success");
      }
      setTimeout(() => { if (mainInputRef.current) mainInputRef.current.focus(); }, 100);
    } catch (error) {
      if (!isBackgroundUpdate) showToast("Lỗi kết nối tải dữ liệu!", "error");
    } finally {
      if (!isBackgroundUpdate) setIsLoadingData(false);
    }
  }, [showToast]);

  useEffect(() => {
    function handleClickOutside(event) {
      if (whDropdownRef.current && !whDropdownRef.current.contains(event.target)) {
        setIsWhDropdownOpen(false);
      }
    }
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  useEffect(() => {
    if (showPasswordModal && passwordInputRef.current) {
      passwordInputRef.current.focus();
    }
  }, [showPasswordModal]);

  useEffect(() => {
    if (!window.XLSX) {
      const script = document.createElement('script');
      script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
      script.async = true;
      script.onload = () => {
         const cachedData = localStorage.getItem('oppo_excel_data');
         const cachedWh = localStorage.getItem('oppo_warehouses');
         
         if (cachedData && cachedWh) {
             try {
                 const pData = JSON.parse(cachedData);
                 const pWh = JSON.parse(cachedWh);
                 setExcelData(pData);
                 setWarehouses(pWh);
                 if (pWh.length > 0) setSelectedWarehouse(String(pWh[0]));
                 fetchGoogleSheetData(true);
             } catch (e) {
                 fetchGoogleSheetData(false);
             }
         } else {
             fetchGoogleSheetData(false);
         }
      };
      document.body.appendChild(script);
    } else {
       const cachedData = localStorage.getItem('oppo_excel_data');
       const cachedWh = localStorage.getItem('oppo_warehouses');
       if (cachedData && cachedWh) {
           try {
               setExcelData(JSON.parse(cachedData));
               setWarehouses(JSON.parse(cachedWh));
               const pWh = JSON.parse(cachedWh);
               if (pWh.length > 0) setSelectedWarehouse(String(pWh[0]));
               fetchGoogleSheetData(true);
           } catch (e) {
               fetchGoogleSheetData(false);
           }
       } else {
           fetchGoogleSheetData(false);
       }
    }
  }, [fetchGoogleSheetData]);

  useEffect(() => {
    const isModalOpen = pendingQRCheck || showDetailedProgressModal || confirmDialog || isWhDropdownOpen || showForceRemarkModal || warningAlert || showPasswordModal || quickNoXacRow || isUploading || showOOWOverrideInput;
    if (!isModalOpen && !hasErrorLock && !hasConflictLock && !isLoadingData) {
      if (isOOWOverrideActive && oowScanInputRef.current) {
         oowScanInputRef.current.focus();
         oowScanInputRef.current.select();
      } else if (mainInputRef.current) {
         mainInputRef.current.focus();
         mainInputRef.current.select();
      }
    }
  }, [pendingQRCheck, hasErrorLock, hasConflictLock, scannedRecords.length, showDetailedProgressModal, confirmDialog, selectedWarehouse, isWhDropdownOpen, showForceRemarkModal, warningAlert, showPasswordModal, quickNoXacRow, isUploading, showOOWOverrideInput, isOOWOverrideActive, isLoadingData]);

  const handleMainScan = useCallback((e) => {
    if (e.key !== 'Enter') return;
    const rawScan = String(scanInput).trim();
    if (!rawScan) return;

    if (hasErrorLock || hasConflictLock) return;
    if (excelData.length === 0 || !selectedWarehouse) return;

    let isQRFirstScan = false;
    let extractedSPString = String(rawScan).split('_').join('').split(' ').join('').toUpperCase();
    
    let matchedRow = excelData.find(row => String(row["SP"] || '') === extractedSPString && String(row["Warehouse Name"] || '') === selectedWarehouse);

    if (!matchedRow) {
       matchedRow = excelData.find(row => String(row["QRCode"] || '').trim() === rawScan && String(row["Warehouse Name"] || '') === selectedWarehouse && rawScan !== '');
       if (matchedRow) {
           extractedSPString = String(matchedRow["SP"] || '').trim();
           isQRFirstScan = true;
       }
    }

    if (!matchedRow) {
       setScannedRecords(prev => [createRecord(rawScan, "Không đúng Xác LK hoặc QR ko tồn tại"), ...prev]);
       setScanInput('');
       return;
    }

    if (lockedMaLK && String(matchedRow["Defective material code"] || '') !== lockedMaLK) {
        const newVCount = violationCount + 1;
        setViolationCount(newVCount);
        if (newVCount >= 3) {
           setShowForceRemarkModal(true);
        } else {
           setWarningAlert(`Bạn phải quét cho xong toàn bộ Mã LK: ${lockedMaLK}.\nCòn thiếu ${missingRowsForLockedLK.length} số phiếu chưa scan.\n(Cảnh báo lần ${newVCount})`);
        }
        setScanInput('');
        return;
    }

    const isDuplicate = currentDisplayedRecords.some(r => String(r.status) === "Khớp, Trả Xác LK về" && (String(r.sp || r.rawScan).split('_').join('').split(' ').join('').toUpperCase()) === extractedSPString);
    if (isDuplicate) {
       setScannedRecords(prev => [createRecord(rawScan, "Bị trùng, SP hoặc QR này đã được scan"), ...prev]);
       setScanInput('');
       return;
    }

    const repairType = String(matchedRow["Repair Type"] || '').trim().toUpperCase();
    const type = String(matchedRow["Type"] || '').trim().toUpperCase();

    if (repairType === 'IW' && (type === 'MAIN' || type === 'LCD')) {
      setPendingQRCheck(isQRFirstScan ? { type: 'WAITING_FOR_SP', rowData: matchedRow, qrScan: rawScan } : { type: 'WAITING_FOR_QR', rowData: matchedRow, spScan: rawScan });
    } else {
      setScannedRecords(prev => [createRecord(rawScan, "Khớp, Trả Xác LK về", matchedRow), ...prev]);
    }
    setScanInput('');
  }, [scanInput, hasErrorLock, hasConflictLock, excelData, selectedWarehouse, currentDisplayedRecords, lockedMaLK, violationCount, missingRowsForLockedLK]);

  const handleSecondScan = useCallback((e) => {
    if (e.key !== 'Enter') return;
    const secondScan = String(secondScanInput).trim();
    if (!secondScan || !pendingQRCheck) return;

    const { type, rowData, spScan, qrScan } = pendingQRCheck;
    if (type === 'WAITING_FOR_QR') {
      const isMatched = secondScan === String(rowData["QRCode"] || '').trim();
      setScannedRecords(prev => [createRecord(spScan, isMatched ? "Khớp, Trả Xác LK về" : "QR Tem Xác và LK ko khớp nhau", rowData, { secondQRScanned: secondScan }), ...prev]);
    } else {
      const spString = String(secondScan).split('_').join('').split(' ').join('').toUpperCase();
      const isMatched = spString === String(rowData["SP"] || '').trim();
      if (isMatched) {
        const isDuplicate = currentDisplayedRecords.some(r => String(r.status) === "Khớp, Trả Xác LK về" && (String(r.sp || r.rawScan).split('_').join('').split(' ').join('').toUpperCase()) === spString);
        setScannedRecords(prev => [createRecord(secondScan, isDuplicate ? "Bị trùng, phải xóa dòng này" : "Khớp, Trả Xác LK về", rowData, { qrScan }), ...prev]);
      } else {
        setScannedRecords(prev => [createRecord(secondScan, "Số phiếu/Mã LK không khớp với QR Tem Xác", rowData), ...prev]);
      }
    }

    setPendingQRCheck(null);
    setSecondScanInput('');
    setTimeout(() => {
       if (isOOWOverrideActive && oowScanInputRef.current) { oowScanInputRef.current.focus(); oowScanInputRef.current.select(); }
       else if (mainInputRef.current) { mainInputRef.current.focus(); mainInputRef.current.select(); }
    }, 150);
  }, [secondScanInput, pendingQRCheck, isOOWOverrideActive, currentDisplayedRecords]);

  const handleDeleteRecord = useCallback((id) => {
    setScannedRecords(prev => prev.filter(record => record.id !== id));
  }, []);

  const handleBulkForceRemark = () => {
    const newRecords = [];
    let hasError = false;

    Object.keys(selectedMissingRows).forEach(idx => {
       if (selectedMissingRows[idx]) {
          const remarkText = rowRemarks[idx];
          if (!remarkText || !String(remarkText).trim()) {
             hasError = true;
          } else {
             const selectedRow = missingRowsForLockedLK[idx];
             const rawSP = String(selectedRow["SP"] || '');
             const newRecord = createRecord(rawSP, "Không xác LK", selectedRow);
             newRecord.remark = String(remarkText).trim();
             newRecords.push(newRecord);
          }
       }
    });

    if (hasError) { showToast("Vui lòng nhập đầy đủ lý do cho các phiếu đã được tick chọn!", "warning"); return; }
    if (newRecords.length === 0) { showToast("Vui lòng tick chọn ít nhất 1 phiếu để xác nhận!", "warning"); return; }

    setScannedRecords(prev => [...newRecords, ...prev]);
    setShowForceRemarkModal(false);
    setViolationCount(0);
    setRowRemarks({});
    setSelectedMissingRows({});
    showToast(`Đã ghi nhận ${newRecords.length} phiếu Không xác LK thành công.`, "success");
  };

  const handleQuickNoXacSubmit = () => {
    if (!quickNoXacReason || !String(quickNoXacReason).trim()) return;
    const newRecord = createRecord(String(quickNoXacRow["SP"] || ''), "Không xác LK", quickNoXacRow);
    newRecord.remark = String(quickNoXacReason).trim();
    setScannedRecords(prev => [newRecord, ...prev]);
    setQuickNoXacRow(null);
    setQuickNoXacReason('');
    showToast(`Đã báo mất xác LK số phiếu: ${newRecord.soRO}`, "success");
  };

  const executeUploadToDrive = async (workbook, isSilent = false) => {
    if (!gasUrl || !String(gasUrl).trim()) return false;
    if (!isSilent) setIsUploading(true);
    
    try {
      const sheetsData = {};
      workbook.SheetNames.forEach(sheetName => {
        sheetsData[sheetName] = window.XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "" });
      });

      const folderName = `${currentScCode || 'Unknown'} - ${selectedWarehouse || 'Unknown'}`;
      const now = new Date();
      const filename = `BaoCao_XacLK_Thang_${String(now.getMonth() + 1).padStart(2, '0')}_${now.getFullYear()}`;

      const response = await fetch(String(gasUrl), {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ folderName, fileName: filename, sheets: sheetsData })
      });

      const textRes = await response.text();
      try {
        const res = JSON.parse(textRes);
        return !!res.success;
      } catch (parseError) {
        return false;
      }
    } catch (err) {
      return false;
    } finally {
      if (!isSilent) setIsUploading(false);
    }
  };

  const canExecuteCenterAction = useCallback(() => {
    if (!selectedWarehouse) { showToast("Vui lòng chọn Trung Tâm Bảo Hành!", "warning"); return false; }
    if (hasErrorLock) { showToast("Hệ thống đang khóa do có dòng lỗi! Vui lòng xóa các dòng lỗi (màu đỏ) trước khi thực hiện.", "error"); return false; }
    if (hasConflictLock) { showToast("Vui lòng XÓA dòng 'Không xác LK' bị trùng ở đầu lịch sử trước!", "error"); return false; }
    if (currentDisplayedRecords.length === 0) { showToast("Không có lịch sử quét nào!", "warning"); return false; }
    if (!currentDisplayedRecords.some(r => String(r.status) === "Khớp, Trả Xác LK về")) { showToast("Lịch sử chỉ có phiếu 'Không xác LK'. Quét ít nhất 1 linh kiện thực tế!", "error"); return false; }
    return true;
  }, [selectedWarehouse, currentDisplayedRecords, hasErrorLock, hasConflictLock, showToast]);

  const handleUploadToDrive = async () => {
    if (!canExecuteCenterAction()) return;
    if (!window.XLSX) { showToast("Thư viện Excel đang được tải.", "warning"); return; }
    
    const workbook = generateWorkbookData(selectedWarehouse, currentDisplayedRecords, detailedProgressList);
    const success = await executeUploadToDrive(workbook, false);
    if (success) showToast(`Đẩy lên Google Drive thành công!`, "success");
    else showToast("Đồng bộ thất bại! Hãy kiểm tra quyền truy cập Apps Script.", "error");
  };

  const handleSaveSession = async () => {
    if (!canExecuteCenterAction()) return;
    const dataStr = JSON.stringify(scannedRecords, null, 2);
    const blob = new Blob([dataStr], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `PhienScan_${selectedWarehouse ? String(selectedWarehouse).split(' ').join('') : 'All'}_${new Date().toISOString().slice(0, 10).split('-').join('')}.json`;
    document.body.appendChild(link); link.click(); document.body.removeChild(link); URL.revokeObjectURL(url);
    showToast("Đã tải xuống file sao lưu phiên quét JSON.", "success");

    if (gasUrl && String(gasUrl).trim()) {
      if (!window.XLSX) return;
      const workbook = generateWorkbookData(selectedWarehouse, currentDisplayedRecords, detailedProgressList);
      const uploadSuccess = await executeUploadToDrive(workbook, true);
      if (uploadSuccess) showToast("Đã tự động sao lưu Google Sheet lên Drive thành công!", "success");
    }
  };

  const handleExportExcel = useCallback(async () => {
    if (!canExecuteCenterAction()) return;
    const workbook = generateWorkbookData(selectedWarehouse, currentDisplayedRecords, detailedProgressList);
    window.XLSX.writeFile(workbook, `BaoCao_ScanLK_${selectedWarehouse ? String(selectedWarehouse).split(' ').join('') : 'All'}_${new Date().toISOString().slice(0, 10).split('-').join('')}.xlsx`);
    showToast("Đã xuất báo cáo Excel thành công.", "success");

    if (gasUrl && String(gasUrl).trim()) {
      const uploadSuccess = await executeUploadToDrive(workbook, true);
      if (uploadSuccess) showToast("Đã tự động cập nhật báo cáo lên Google Drive!", "success");
    }
  }, [currentDisplayedRecords, detailedProgressList, selectedWarehouse, gasUrl, currentScCode, canExecuteCenterAction]);

  const handleLoadSession = useCallback((e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const loadedData = JSON.parse(evt.target.result);
        if (!Array.isArray(loadedData)) throw new Error();
        
        const loadedWarehouses = [...new Set(loadedData.map(r => String(r.ttbh || "").trim()).filter(Boolean))];
        setWarehouses(prev => {
          const combined = [...new Set([...prev, ...loadedWarehouses])];
          return combined.sort((a, b) => String(a).localeCompare(String(b), 'vi'));
        });
        if (loadedWarehouses.length === 1 && selectedWarehouse !== loadedWarehouses[0]) setSelectedWarehouse(String(loadedWarehouses[0]));

        setScannedRecords(prevOnScreenRecords => {
          const onScreenMap = new Map();
          prevOnScreenRecords.forEach(r => onScreenMap.set(`${String(r.rawScan).toUpperCase()}|${String(r.ttbh).trim()}`, r));

          let addCount = 0, skipCount = 0, conflictCount = 0;
          const mergedList = [...prevOnScreenRecords];

          loadedData.forEach(loadedItem => {
            const cleanItem = { ...loadedItem, sp: loadedItem.sp ? String(loadedItem.sp) : String(loadedItem.rawScan || '').split('_').join('').split(' ').join('').toUpperCase() };
            const uniqueKey = `${String(cleanItem.rawScan).toUpperCase()}|${String(cleanItem.ttbh).trim()}`;
            if (onScreenMap.has(uniqueKey)) {
              if (String(onScreenMap.get(uniqueKey).status) === String(cleanItem.status)) skipCount++;
              else { mergedList.push(cleanItem); addCount++; conflictCount++; }
            } else { mergedList.push(cleanItem); addCount++; }
          });

          if (conflictCount > 0) showToast(`Phát hiện ${conflictCount} phiếu xung đột trạng thái! Hệ thống KHÓA CHỜ xử lý xóa.`, "error");
          else if (addCount > 0) showToast(`Đã gộp thành công ${addCount} dòng mới.`, "success");
          else showToast(`Toàn bộ dữ liệu đều đã có sẵn trên màn hình.`, "warning");

          return mergedList;
        });
      } catch (err) { showToast("Lỗi định dạng tệp file phiên!", "error"); }
      finally { e.target.value = null; }
    };
    reader.readAsText(file);
  }, [selectedWarehouse, showToast]);

  const handleOOWOverrideSubmit = useCallback(() => {
    if (oowOverridePassword === 'kk134') {
       setIsOOWOverrideActive(true); setShowOOWOverrideInput(false); setOowOverridePassword('');
       showToast("Đã kích hoạt ô quét KMH chuyên dụng cho nhóm OOW-LCD!", "success");
    } else { setShowOOWOverrideInput(false); setOowOverridePassword(''); }
  }, [oowOverridePassword, showToast]);

  const handleOOWScan = useCallback((e) => {
    if (e.key !== 'Enter') return;
    const rawScan = String(oowScanInput).trim();
    if (!rawScan || hasErrorLock || hasConflictLock) return;

    let extracted = String(rawScan).split('_').join('').split(' ').join('').toUpperCase();
    let row = excelData.find(r => String(r["SP"] || '') === extracted && String(r["Warehouse Name"] || '') === selectedWarehouse);
    if (!row) {
      row = excelData.find(r => String(r["QRCode"] || '').trim() === rawScan && String(r["Warehouse Name"] || '') === selectedWarehouse && rawScan !== '');
    }

    if (!row || String(row["Repair Type"]).toUpperCase() !== 'OOW' || String(row["Type"]).toUpperCase() !== 'LCD') {
      showToast("Lỗi: Chỉ chấp nhận mã linh kiện OOW và phân nhóm LCD!", "error");
      setOowScanInput('');
      return;
    }

    if (currentDisplayedRecords.some(r => String(r.status) === "Khớp, Trả Xác LK về" && (String(r.sp || r.rawScan).split('_').join('').split(' ').join('').toUpperCase()) === String(row["SP"] || '').trim().toUpperCase())) {
       showToast("Lỗi: Mã này đã được scan rồi!", "error");
       setOowScanInput('');
       return;
    }

    const rec = createRecord(rawScan, "Khớp, Trả Xác LK về", row);
    rec.remark = "KMH";
    setScannedRecords(prev => [rec, ...prev]);
    setOowScanInput('');
    showToast(`Bắn mã KMH thành công số phiếu: ${rec.soRO}`, "success");
  }, [oowScanInput, excelData, selectedWarehouse, hasErrorLock, hasConflictLock, showToast, currentDisplayedRecords]);

  const handleOverrideSubmit = () => {
    if (overridePassword === 'kk134') {
      setIsGlobalOverrideActive(true); setLockedMaLK(null); setViolationCount(0); setShowForceRemarkModal(false); setOverridePassword(''); setShowOverrideInput(false);
      showToast("Đã mở khóa vượt cấp thành công! Trạm có thể tự do quét mã.", "success");
    } else { setOverridePassword(''); setShowOverrideInput(false); }
  };

  const executeExportAllExcel = useCallback(() => {
    const workbook = window.XLSX.utils.book_new();
    const globalScannedSPs = scannedRecords.filter(r => String(r.status) === "Khớp, Trả Xác LK về" || String(r.status) === "Không xác LK").map(r => String(r.sp || r.rawScan).split('_').join('').split(' ').join('').toUpperCase());
    const globalScannedKMH_SPs = new Set();
    scannedRecords.forEach(r => { if ((String(r.status) === "Khớp, Trả Xác LK về" || String(r.status) === "Không xác LK") && String(r.remark).includes("KMH")) globalScannedKMH_SPs.add(String(r.sp || r.rawScan).split('_').join('').split(' ').join('').toUpperCase()); });

    const globalDetailedList = excelData.map(row => ({ ...row, isScanned: globalScannedSPs.includes(String(row["SP"] || "").trim()) }));
    const statsData = [];
    
    warehouses.forEach(wh => {
      const whData = globalDetailedList.filter(r => String(r["Warehouse Name"] || "").trim() === String(wh));
      if (whData.length === 0) return;

      const iwData = whData.filter(r => String(r["Repair Type"]).toUpperCase() === 'IW');
      const iwGroups = {};
      iwData.forEach(r => {
        const type = (String(r["Type"]).toUpperCase() === 'LCD' || String(r["Type"]).toUpperCase() === 'MAIN') ? String(r["Type"]).toUpperCase() : 'OTHERS';
        const key = `${type}|${String(r["Defective material code"])}|${String(r["Product Name"])}`;
        if (!iwGroups[key]) iwGroups[key] = { type, maLK: String(r["Defective material code"] || ''), tenLK: String(r["Product Name"] || ''), total: 0, scanned: 0 };
        iwGroups[key].total += 1;
        if (r.isScanned) iwGroups[key].scanned += 1;
      });
      Object.values(iwGroups).forEach(g => statsData.push({ "TTBH": String(wh), "Phân Loại": "IW", "Nhóm LK": g.type, "Mã LK": g.maLK, "Tên LK": g.tenLK, "Cần Scan": g.total, "Đã Scan": g.scanned, "Còn Thiếu": g.total - g.scanned }));

      const isRowKMH = (r) => {
          const sp = String(r["SP"] || '').trim().toUpperCase();
          return String(r["KMH"]) === "1" || String(r["KMH"]).includes("KMH") || globalScannedKMH_SPs.has(sp);
      };

      const oowData = whData.filter(r => String(r["Repair Type"]).toUpperCase() === 'OOW' && !isRowKMH(r));
      const oowGroups = { 'LCD': { total: 0, scanned: 0 }, 'MAIN': { total: 0, scanned: 0 }, 'OTHERS': { total: 0, scanned: 0 } };
      oowData.forEach(r => {
        const type = (String(r["Type"]).toUpperCase() === 'LCD' || String(r["Type"]).toUpperCase() === 'MAIN') ? String(r["Type"]).toUpperCase() : 'OTHERS';
        oowGroups[type].total += 1;
        if (r.isScanned) oowGroups[type].scanned += 1;
      });
      ['LCD', 'MAIN', 'OTHERS'].forEach(type => { if (oowGroups[type].total > 0) statsData.push({ "TTBH": String(wh), "Phân Loại": "OOW", "Nhóm LK": type, "Mã LK": "-", "Tên LK": `Tổng số lượng ${type} (OOW)`, "Cần Scan": oowGroups[type].total, "Đã Scan": oowGroups[type].scanned, "Còn Thiếu": oowGroups[type].total - oowGroups[type].scanned }); });

      const otherData = whData.filter(r => !['IW', 'OOW'].includes(String(r["Repair Type"]).toUpperCase()));
      const otherGroups = { 'LCD': { total: 0, scanned: 0 }, 'MAIN': { total: 0, scanned: 0 }, 'OTHERS': { total: 0, scanned: 0 } };
      otherData.forEach(r => {
        const type = (String(r["Type"]).toUpperCase() === 'LCD' || String(r["Type"]).toUpperCase() === 'MAIN') ? String(r["Type"]).toUpperCase() : 'OTHERS';
        otherGroups[type].total += 1;
        if (r.isScanned) otherGroups[type].scanned += 1;
      });
      ['LCD', 'MAIN', 'OTHERS'].forEach(type => { if (otherGroups[type].total > 0) statsData.push({ "TTBH": String(wh), "Phân Loại": "Khác", "Nhóm LK": type, "Mã LK": "-", "Tên LK": `Tổng số lượng ${type} (Khác)`, "Cần Scan": otherGroups[type].total, "Đã Scan": otherGroups[type].scanned, "Còn Thiếu": otherGroups[type].total - otherGroups[type].scanned }); });
    });

    window.XLSX.utils.book_append_sheet(workbook, window.XLSX.utils.json_to_sheet(statsData), "ThongKe_TongHop");

    const modelGroups = {};
    globalDetailedList.forEach(r => {
        const repairType = String(r["Repair Type"] || 'Unknown').trim();
        const sp = String(r["SP"] || '').trim().toUpperCase();
        if (repairType.toUpperCase() === 'OOW' && (String(r["KMH"]) === "1" || String(r["KMH"]).includes("KMH") || globalScannedKMH_SPs.has(sp))) return;

        const key = `${String(r["Product Model"] || r["Model"] || 'Unknown').trim()}|${String(r["Type"] || 'Unknown').trim()}|${repairType}`;
        if (!modelGroups[key]) modelGroups[key] = { model: String(r["Product Model"] || r["Model"] || 'Unknown').trim(), type: String(r["Type"] || 'Unknown').trim(), repairType, totalSlg: 0, scannedSlg: 0 };
        modelGroups[key].totalSlg += parseInt(r["Consumed quantity"] || r["Consumed"] || 1);
        if (r.isScanned) modelGroups[key].scannedSlg += parseInt(r["Consumed quantity"] || r["Consumed"] || 1);
    });

    const modelStatsData = Object.values(modelGroups).map(g => ({ "Model": g.model, "Type": g.type, "BH/DV": g.repairType, "Tổng Cần Scan (Slg)": g.totalSlg, "Đã Scan Khớp (Slg)": g.scannedSlg, "Còn Thiếu (Slg)": g.totalSlg - g.scannedSlg })).sort((a, b) => String(a.Model).localeCompare(String(b.Model)));
    window.XLSX.utils.book_append_sheet(workbook, window.XLSX.utils.json_to_sheet(modelStatsData), "ThongKe_TheoModel");

    const exportScannedData = scannedRecords.map(record => ({ "Trạng thái": String(record.status), "Cột SP": String(record.sp || record.rawScan), "SC Code": String(record.scCode), "Warehouse Name": String(record.ttbh), "Số RO": String(record.soRO), "BH/DV": String(record.bhDv), "Mã LK": String(record.maLK), "Product Name": String(record.tenLK), "Model": String(record.model), "Type": String(record.phanLoai), "Slg": String(record.slg), "Remark": String(record.remark) }));
    window.XLSX.utils.book_append_sheet(workbook, window.XLSX.utils.json_to_sheet(exportScannedData), "LichSu_DaScan_TatCa");

    const wsUnscannedData = globalDetailedList.filter(r => !r.isScanned).map(record => ({ "Trạng thái": "Chưa Scan", "Cột SP": String(record["SP"] || ""), "SC Code": String(record["SC Code"] || record["SC code"] || ''), "Warehouse Name": String(record["Warehouse Name"] || ""), "Số RO": String(record["After-sales work order No."] || ""), "BH/DV": String(record["Repair Type"] || ""), "Mã LK": String(record["Defective material code"] || ""), "Product Name": String(record["Product Name"] || ""), "Model": String(record["Product Model"] || record["Model"] || ''), "Type": String(record["Type"] || ""), "Slg": String(record["Consumed quantity"] || record["Consumed"] || ''), "Remark": String(getRemark(record)) }));
    window.XLSX.utils.book_append_sheet(workbook, window.XLSX.utils.json_to_sheet(wsUnscannedData), "DanhSach_ChuaScan_TatCa");

    window.XLSX.writeFile(workbook, `BaoCao_ScanLK_TongHop_${new Date().toISOString().slice(0, 10).split('-').join('')}.xlsx`);
    showToast("Đã tải xuống file báo cáo TỔNG HỢP thành công.", "success");
  }, [excelData, scannedRecords, warehouses, showToast]);

  const handleExportAllExcel = useCallback(() => {
    if (excelData.length === 0 && scannedRecords.length === 0) { showToast("Không có dữ liệu để xuất Excel!", "warning"); return; }
    if (!window.XLSX) { showToast("Thư viện Excel đang tải.", "warning"); return; }
    setShowPasswordModal(true);
  }, [excelData, scannedRecords, showToast]);

  const handlePasswordSubmit = () => { if (passwordInput === '11221122a') executeExportAllExcel(); setPasswordInput(''); setShowPasswordModal(false); };

  const renderStatusHTML = (status) => {
    const s = String(status);
    if (s === "Khớp, Trả Xác LK về") return <span className="inline-flex items-center text-emerald-600 font-bold text-sm"><CheckCircle className="w-5 h-5 mr-1.5" /> Khớp, Trả Xác</span>;
    if (s === "Không xác LK") return <span className="inline-flex items-center text-gray-600 font-bold bg-gray-100 border border-gray-300 px-3 py-1.5 rounded-md text-sm"><AlertTriangle className="w-5 h-5 mr-1.5" /> Không xác LK</span>;
    if (s.includes("trùng")) return <span className="inline-flex items-center text-red-600 font-bold bg-red-100 px-3 py-1.5 rounded-md text-sm"><Ban className="w-5 h-5 mr-1.5" /> Lỗi trùng lặp</span>;
    return <span className="inline-flex items-center text-rose-500 font-bold bg-rose-50 px-3 py-1.5 rounded-md text-sm"><XCircle className="w-5 h-5 mr-1.5" /> Sai linh kiện / QR</span>;
  };

  const renderDetailedTableGroup = useCallback((title, dataList, titleColorClass) => {
    if (dataList.length === 0) return null;
    const isOOWGroup = String(title).includes("OOW");

    return (
      <div className="mb-8">
        <div className="flex flex-col lg:flex-row lg:items-center gap-3 mb-4">
          <h3 className={`font-bold text-base flex items-center uppercase ${titleColorClass} m-0`}>
            <span className="w-2 h-6 bg-current mr-2 rounded-sm opacity-80"></span>
            Nhóm phiếu {String(title)} 
            <span className="ml-2 bg-gray-100 text-gray-700 px-3 py-0.5 rounded-full text-sm border border-gray-200">{Number(dataList.length)}</span>
          </h3>
          
          {isOOWGroup && (
            <div className="flex items-center ml-2">
              {!isOOWOverrideActive ? (
                !showOOWOverrideInput ? (
                  <button onClick={() => setShowOOWOverrideInput(true)} className="text-xs text-gray-400 hover:text-blue-600 font-semibold underline decoration-dotted underline-offset-2 transition-colors cursor-pointer">
                    Mở khóa vượt cấp
                  </button>
                ) : (
                  <div className="flex items-center gap-2">
                    <input type="password" value={oowOverridePassword} onChange={e => setOowOverridePassword(e.target.value)} onKeyDown={e => { if (e.key === 'Enter') handleOOWOverrideSubmit(); }} placeholder="Nhập pass..." className="w-28 border border-gray-300 rounded-lg px-2 py-1 text-sm focus:outline-none focus:border-blue-500 bg-white" autoFocus />
                    <button onClick={handleOOWOverrideSubmit} className="text-xs bg-blue-600 text-white px-3 py-1.5 rounded-lg font-bold shadow-sm">OK</button>
                    <button onClick={() => { setShowOOWOverrideInput(false); setOowOverridePassword(''); }} className="text-xs text-gray-500 font-bold px-1 hover:text-gray-700">Hủy</button>
                  </div>
                )
              ) : (
                <div className="relative flex items-center ml-2">
                  <ScanLine className="absolute left-3 w-4 h-4 text-blue-600" />
                  <input ref={oowScanInputRef} type="text" value={oowScanInput} onChange={e => setOowScanInput(e.target.value)} onKeyDown={handleOOWScan} placeholder="Scan trực tiếp mã KMH (OOW-LCD)..." className="border-2 border-blue-400 bg-blue-50 rounded-xl pl-9 pr-4 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-300 text-blue-900 font-semibold w-72 shadow-inner" />
                </div>
              )}
            </div>
          )}
        </div>
        <table className="w-full text-left border-collapse whitespace-nowrap bg-white rounded-lg overflow-hidden shadow-sm border border-gray-200">
          <thead className="bg-gray-100 font-bold text-sm uppercase tracking-wider text-gray-600 border-b border-gray-200">
            <tr>
              <th className="p-3 px-4">Trạng Thái</th>
              <th className="p-3 px-4">Số RO</th>
              <th className="p-3 px-4">Mã LK</th>
              <th className="p-3 px-4">Tên LK</th>
              <th className="p-3 px-4">Model</th>
              <th className="p-3 px-4">Loại</th>
              <th className="p-3 px-4 text-center">Slg</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {dataList.map((row, idx) => {
              const repairType = String(row["Repair Type"] || '').trim().toUpperCase();
              const typeLK = String(row["Type"] || '').trim().toUpperCase();
              const allowQuickNoXac = repairType === 'OOW' || (repairType === 'IW' && typeLK === 'PIN');

              return (
                <tr key={idx} className={`hover:bg-gray-50/80 transition-colors ${row.isScanned ? 'bg-gray-50/50 opacity-60' : ''}`}>
                  <td className="p-3 px-4 w-[300px]">
                    <div className="flex items-center gap-2">
                      {row.isScanned ? (
                        <span className="inline-flex items-center text-gray-500 font-bold bg-gray-100 px-3 py-1.5 rounded text-sm border border-gray-200"><CheckCircle className="w-4 h-4 mr-1.5" /> Đã Scan</span>
                      ) : (
                        <>
                          <span className="inline-flex items-center text-red-500 font-bold bg-red-50 px-3 py-1.5 rounded text-sm border border-red-100"><AlertTriangle className="w-4 h-4 mr-1.5" /> Chưa Scan</span>
                          {allowQuickNoXac && (
                            <button onClick={() => { setQuickNoXacRow(row); setQuickNoXacReason(''); }} className="text-xs font-bold bg-red-100 text-red-700 hover:bg-red-200 px-2.5 py-1.5 rounded border border-red-200 transition-colors shadow-sm shrink-0">Không xác LK</button>
                          )}
                        </>
                      )}
                    </div>
                  </td>
                  <td className="p-3 px-4">
                    <div className="flex items-center gap-2">
                      <span className="font-mono text-[15px] text-gray-700">{String(row["After-sales work order No."] || "")}</span>
                      {row["After-sales work order No."] && (
                        <a href={`https://gcsm-sg.oppoit.com/order/order-management/after-sales-order/${String(row["After-sales work order No."])}/detail`} target="_blank" rel="noopener noreferrer" className="inline-flex items-center text-xs font-bold bg-blue-50 text-blue-600 hover:bg-blue-100 hover:text-blue-700 px-2 py-1 rounded border border-blue-200 transition-colors" title="Kiểm tra hệ thống GCSM"><ExternalLink className="w-3.5 h-3.5 mr-1" /> Check GCSM</a>
                      )}
                    </div>
                  </td>
                  <td className="p-3 px-4 font-mono text-[15px] text-gray-700">{String(row["Defective material code"] || "")}</td>
                  <td className="p-3 px-4 text-gray-700 max-w-[280px] truncate font-semibold text-[15px]" title={String(row["Product Name"] || "")}>{String(row["Product Name"] || "")}</td>
                  <td className="p-3 px-4 text-gray-600 text-[15px]">{String(row["Product Model"] || row["Model"] || '')}</td>
                  <td className="p-3 px-4 text-gray-500 text-sm font-medium">{String(row["Type"] || "")}</td>
                  <td className="p-3 px-4 font-bold text-gray-800 text-center text-[15px]">{String(row["Consumed quantity"] || row["Consumed"] || '1')}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    );
  }, [quickNoXacRow, quickNoXacReason, showToast, showOOWOverrideInput, oowOverridePassword, isOOWOverrideActive, oowScanInput, handleOOWOverrideSubmit, handleOOWScan]);

  return (
    <div className="h-screen flex flex-col bg-[#f8f9fa] text-base font-sans text-gray-800 overflow-hidden">
      
      {(isLoadingData || isUploading) && (
        <div className="fixed inset-0 bg-gray-900/50 backdrop-blur-sm flex items-center justify-center z-[9999] transition-all duration-300">
          <div className="bg-white px-8 py-5 rounded-2xl shadow-2xl flex items-center gap-4 border border-gray-100">
            <Loader2 className="w-8 h-8 text-blue-600 animate-spin" />
            <span className="font-bold text-xl text-gray-800 tracking-tight">Đang tải dữ liệu...</span>
          </div>
        </div>
      )}

      {toast && (
        <div className={`fixed bottom-8 right-8 z-[300] px-6 py-4 rounded-xl shadow-2xl text-white font-bold text-base transition-all duration-300 transform translate-y-0 opacity-100 ${String(toast.type) === 'error' ? 'bg-red-500' : String(toast.type) === 'success' ? 'bg-emerald-500' : String(toast.type) === 'warning' ? 'bg-amber-500' : 'bg-blue-500'}`}>
          {String(toast.message)}
        </div>
      )}

      {warningAlert && (
        <div className="fixed inset-0 bg-gray-900/60 backdrop-blur-sm flex items-center justify-center z-[250] p-4">
          <div className="bg-white rounded-xl shadow-2xl p-6 max-w-sm w-full border border-red-100 text-center relative z-50">
            <AlertTriangle className="w-16 h-16 text-amber-500 mx-auto mb-5 animate-bounce" />
            <h3 className="font-bold text-xl text-amber-600 mb-2">CẢNH BÁO (Lần {Number(violationCount)})</h3>
            <p className="text-gray-700 whitespace-pre-line mb-6 font-medium text-base">{String(warningAlert)}</p>
            <button onClick={() => setWarningAlert(null)} className="w-full py-3 text-base bg-amber-500 text-white rounded-xl hover:bg-amber-600 font-bold transition-colors shadow-md">Tôi đã hiểu</button>
          </div>
        </div>
      )}

      {confirmDialog && (
        <div className="fixed inset-0 bg-gray-900/40 backdrop-blur-sm flex items-center justify-center z-[150] p-4">
          <div className="bg-white rounded-2xl shadow-2xl p-8 max-w-md w-full border border-gray-100 relative z-50">
            <h3 className="font-bold text-xl mb-3 text-gray-800">{String(confirmDialog.title)}</h3>
            <p className="text-gray-600 mb-8 text-base">{String(confirmDialog.message)}</p>
            <div className="flex justify-end gap-4">
              <button onClick={() => setConfirmDialog(null)} className="px-5 py-2.5 bg-gray-100 text-base rounded-xl hover:bg-gray-200 font-bold transition-colors">Hủy</button>
              <button onClick={() => { confirmDialog.onConfirm(); setConfirmDialog(null); }} className="px-5 py-2.5 bg-red-500 text-white text-base rounded-xl hover:bg-red-600 font-bold transition-colors shadow-md">Đồng ý</button>
            </div>
          </div>
        </div>
      )}

      {showPasswordModal && (
        <div className="fixed inset-0 bg-gray-900/60 backdrop-blur-sm flex items-center justify-center z-[250] p-4">
          <div className="bg-white rounded-3xl shadow-2xl p-8 max-w-sm w-full border border-gray-200 relative z-50 text-center">
            <Settings className="w-14 h-14 text-blue-600 mx-auto mb-4 animate-spin" style={{ animationDuration: '4s' }} />
            <h3 className="font-bold text-lg text-gray-800 mb-2">Xác thực xuất dữ liệu</h3>
            <p className="text-sm text-gray-500 mb-6">Vui lòng nhập mật khẩu xác nhận quyền xuất báo cáo TỔNG HỢP toàn quốc.</p>
            <input 
              ref={passwordInputRef}
              type="password"
              placeholder="Nhập mật khẩu..."
              value={passwordInput}
              onChange={(e) => setPasswordInput(e.target.value)}
              onKeyDown={(e) => { if (e.key === 'Enter') handlePasswordSubmit(); }}
              className="w-full text-center border-2 border-gray-200 rounded-xl px-4 py-3 text-base focus:outline-none focus:border-blue-500 bg-gray-50/50 font-mono mb-6"
            />
            <div className="flex gap-3">
              <button onClick={() => { setShowPasswordModal(false); setPasswordInput(''); }} className="flex-1 py-3 text-sm bg-gray-100 text-gray-600 rounded-xl hover:bg-gray-200 font-bold transition-colors">Hủy bỏ</button>
              <button onClick={handlePasswordSubmit} className="flex-1 py-3 text-sm bg-blue-600 text-white rounded-xl hover:bg-blue-700 font-bold transition-colors shadow-lg">Xác nhận</button>
            </div>
          </div>
        </div>
      )}

      {quickNoXacRow && (
        <div className="fixed inset-0 bg-gray-900/80 backdrop-blur-md flex items-center justify-center z-[240] p-4">
          <div className="bg-white rounded-3xl shadow-2xl p-8 max-w-lg w-full border border-red-100 relative z-[250]">
            <div className="flex items-center text-red-600 mb-5 border-b border-gray-100 pb-4">
              <AlertTriangle className="w-8 h-8 mr-3 animate-pulse" />
              <h2 className="text-xl font-bold text-gray-800">Khai Báo Không Xác Linh Kiện</h2>
            </div>
            <div className="bg-red-50 border border-red-200 rounded-2xl p-5 mb-6 space-y-2 text-sm text-red-800">
              <p><strong>Số RO:</strong> {String(quickNoXacRow["After-sales work order No."] || "")}</p>
              <p><strong>Mã LK:</strong> {String(quickNoXacRow["Defective material code"] || "")}</p>
              <p><strong>Tên LK:</strong> {String(quickNoXacRow["Product Name"] || "")}</p>
              <p><strong>Phân Loại:</strong> {String(quickNoXacRow["Repair Type"] || "")} - {String(quickNoXacRow["Type"] || "")}</p>
            </div>
            <div className="mb-8">
              <label className="block text-gray-700 text-sm font-extrabold uppercase mb-3">Nhập lý do báo mất linh kiện (Bắt buộc):</label>
              <input
                type="text"
                value={quickNoXacReason}
                onChange={(e) => setQuickNoXacReason(e.target.value)}
                placeholder="Ví dụ: TTBH báo mất, Khách không trả..."
                className="w-full border-2 border-gray-200 rounded-xl px-5 py-3.5 text-base focus:outline-none focus:border-red-400 bg-gray-50/30 font-semibold"
                autoFocus
              />
            </div>
            <div className="flex justify-end gap-4">
              <button onClick={() => { setQuickNoXacRow(null); setQuickNoXacReason(''); }} className="px-6 py-3 text-sm bg-gray-100 text-gray-600 rounded-xl font-bold transition-colors">Hủy bỏ</button>
              <button onClick={handleQuickNoXacSubmit} className="px-6 py-3 text-sm bg-red-600 text-white rounded-xl font-bold transition-colors shadow-lg">Xác nhận báo mất</button>
            </div>
          </div>
        </div>
      )}

      {showForceRemarkModal && (
        <div className="fixed inset-0 bg-gray-900/85 backdrop-blur-md flex items-center justify-center z-[220] p-4">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-5xl border border-gray-100 flex flex-col max-h-[90vh] relative z-50">
            <div className="p-6 border-b border-gray-100 shrink-0 bg-rose-50 rounded-t-2xl">
               <h3 className="font-bold text-2xl text-rose-700 flex items-center"><AlertTriangle className="w-8 h-8 mr-3" /> Cảnh báo: Phải Scan xong toàn bộ 1 mã LK của IW hoặc OOW</h3>
               <p className="text-rose-600 mt-3 text-base font-medium">Bạn đã cố tình quét mã linh kiện khác khi <strong>Mã LK: {String(lockedMaLK)}</strong> vẫn chưa được quét xong!</p>
            </div>
            <div className="p-6 flex-1 overflow-y-auto">
               <div className="border border-gray-200 rounded-xl overflow-hidden">
                 <table className="w-full text-left text-base whitespace-nowrap">
                   <thead className="bg-gray-50 border-b border-gray-200">
                     <tr>
                       <th className="p-4 px-5 font-bold text-gray-600 w-16 text-center text-sm uppercase">Chọn</th>
                       <th className="p-4 px-5 font-bold text-gray-600 text-sm uppercase">Số RO</th>
                       <th className="p-4 px-5 font-bold text-gray-600 text-sm uppercase">Tên LK</th>
                       <th className="p-4 px-5 font-bold text-gray-600 text-sm uppercase">Nhập Lý do (Bắt buộc)</th>
                     </tr>
                   </thead>
                   <tbody className="divide-y divide-gray-100">
                     {missingRowsForLockedLK.map((row, idx) => {
                        const isStrictIW = String(row["Repair Type"]).toUpperCase() === 'IW' && String(row["Type"]).toUpperCase() !== 'PIN';
                        return (
                          <tr key={idx} className={`transition-colors ${isStrictIW ? 'bg-gray-50/50 opacity-70' : 'hover:bg-gray-50'}`}>
                             <td className="p-3 px-5 text-center">
                               <input type="checkbox" disabled={isStrictIW} checked={!!selectedMissingRows[idx]} onChange={(e) => setSelectedMissingRows({...selectedMissingRows, [idx]: e.target.checked})} className="w-5 h-5 text-blue-600 rounded border-gray-300" />
                             </td>
                             <td className="p-3 px-5 font-mono font-bold text-[15px] text-gray-700">{String(row["After-sales work order No."] || "")}</td>
                             <td className="p-3 px-5 text-gray-700 truncate max-w-[240px] font-medium">{String(row["Product Name"] || "")}</td>
                             <td className="p-3 px-5">
                               {isStrictIW ? <span className="text-red-500 text-sm font-bold">Bắt buộc có xác</span> : <input type="text" placeholder="Lý do..." disabled={!selectedMissingRows[idx]} value={rowRemarks[idx] || ''} onChange={(e) => setRowRemarks({...rowRemarks, [idx]: e.target.value})} className="border rounded-lg px-3 py-2 text-sm w-full bg-white" />}
                             </td>
                          </tr>
                        );
                     })}
                   </tbody>
                 </table>
               </div>
            </div>
            <div className="p-5 border-t border-gray-100 shrink-0 bg-gray-50 rounded-b-2xl flex justify-between items-center">
              <div className="flex flex-col items-start gap-1">
                <span className="text-sm text-gray-500 italic font-medium">Hệ thống sẽ mở khóa quét mã tiếp theo khi hoàn tất.</span>
                {!showOverrideInput ? (
                  <button onClick={() => setShowOverrideInput(true)} className="text-xs text-gray-400 hover:text-blue-600 font-semibold underline decoration-dotted underline-offset-2 transition-colors">Mở khóa vượt cấp</button>
                ) : (
                  <div className="flex items-center gap-2 mt-1">
                    <input type="password" value={overridePassword} onChange={(e) => setOverridePassword(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') handleOverrideSubmit(); }} placeholder="Mã..." className="w-28 border rounded-lg px-2 py-1 text-sm bg-white" autoFocus />
                    <button onClick={handleOverrideSubmit} className="text-xs bg-blue-600 text-white px-3 py-1.5 rounded-lg font-bold">OK</button>
                    <button onClick={() => { setShowOverrideInput(false); setOverridePassword(''); }} className="text-xs text-gray-500 font-bold px-1">Hủy</button>
                  </div>
                )}
              </div>
              <div className="flex gap-4">
                 <button onClick={() => { setShowForceRemarkModal(false); setViolationCount(0); setSelectedMissingRows({}); setRowRemarks({}); setShowOverrideInput(false); setOverridePassword(''); }} className="px-6 py-2.5 bg-gray-200 text-gray-700 rounded-xl font-bold">Đóng tạm thời</button>
                 {canSkipAnyMissing && <button onClick={handleBulkForceRemark} className="px-6 py-2.5 bg-red-500 text-white rounded-xl font-bold shadow-md">Xác nhận Không Xác LK</button>}
              </div>
            </div>
          </div>
        </div>
      )}

      {pendingQRCheck && (
        <div className="fixed inset-0 bg-gray-900/50 backdrop-blur-md flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-3xl shadow-2xl p-8 max-w-lg w-full border border-gray-100">
            <div className="flex items-center text-amber-600 mb-5 border-b border-gray-100 pb-4">
              <AlertTriangle className="w-8 h-8 mr-3" />
              <h2 className="text-xl font-bold text-gray-800">Xác Thực 2 Bước (SP & QR)</h2>
            </div>
            <div className="bg-amber-50 border border-amber-200 rounded-2xl p-5 mb-6">
              <p className="font-bold text-amber-800 text-base">
                {pendingQRCheck.type === 'WAITING_FOR_QR' ? `Yêu cầu Scan QR Tem Xác LK!` : `Đã nhận QR. Vui lòng Scan Số phiếu/Mã LK!`}
              </p>
              <p className="text-sm text-gray-700 mt-2 font-mono">Mã Phiếu: <strong className="bg-white px-2 py-0.5 rounded border border-amber-200 text-[15px]">{String(pendingQRCheck.rowData["After-sales work order No."] || "")}</strong></p>
            </div>
            <div className="mb-6">
              <label className="block text-gray-700 text-base font-bold mb-3">
                {pendingQRCheck.type === 'WAITING_FOR_QR' ? 'Quét mã QR trên tem xác:' : 'Quét mã SP (Số Phiếu_Mã LK):'}
              </label>
              <input ref={secondInputRef} autoFocus type="text" maxLength={100} value={secondScanInput} onChange={(e) => setSecondScanInput(e.target.value)} onKeyDown={handleSecondScan} placeholder="Bắn mã vào đây..." className="w-full border-2 border-amber-300 rounded-xl px-5 py-4 text-lg font-mono focus:outline-none focus:border-amber-500 bg-amber-50/30 text-gray-800 font-semibold" />
            </div>
            <div className="flex justify-end">
              <button onClick={() => { setPendingQRCheck(null); setSecondScanInput(''); }} className="px-6 py-3 text-base bg-gray-100 text-gray-600 rounded-xl font-bold">Hủy bỏ</button>
            </div>
          </div>
        </div>
      )}

      {showDetailedProgressModal && (
        <div className="fixed inset-0 bg-gray-900/60 backdrop-blur-md flex items-center justify-center z-[100] p-4">
          <div className="bg-white rounded-3xl shadow-2xl w-full max-w-6xl max-h-[90vh] flex flex-col border border-gray-100 relative z-50">
            <div className="flex justify-between items-center p-5 border-b border-gray-100 shrink-0 bg-gray-50 rounded-t-3xl">
              <h2 className="text-xl font-extrabold text-gray-800 flex items-center">
                <ClipboardList className="w-6 h-6 mr-3 text-blue-600" /> Danh sách phiếu CẦN SCAN - <span className="text-blue-700 mx-2">{String(selectedWarehouse)}</span> 
                <span className="text-sm bg-gray-200 text-gray-600 px-3 py-1 rounded-full ml-3 font-bold">Tổng: {Number(detailedProgressList.length)}</span>
              </h2>
              <button onClick={() => setShowDetailedProgressModal(false)} className="p-2 text-gray-400 hover:text-red-500 rounded-xl bg-white border border-gray-200 hover:bg-red-50 transition-colors"><X className="w-6 h-6" /></button>
            </div>
            <div className="overflow-auto flex-1 p-6">
              {renderDetailedTableGroup("Bảo Hành (IW)", detailedProgressList.filter(r => String(r["Repair Type"]).toUpperCase() === 'IW'), "text-emerald-700")}
              {renderDetailedTableGroup("Dịch Vụ (OOW)", detailedProgressList.filter(r => String(r["Repair Type"]).toUpperCase() === 'OOW'), "text-blue-700")}
              {renderDetailedTableGroup("Khác", detailedProgressList.filter(r => !['IW', 'OOW'].includes(String(r["Repair Type"]).toUpperCase())), "text-gray-700")}
            </div>
          </div>
        </div>
      )}

      <header className="bg-white border-b border-gray-200 px-6 py-4 shadow-sm flex justify-between items-center shrink-0 z-20 mt-[2px]">
        <h1 className="text-2xl font-black flex items-center text-gray-800 tracking-tight">
          <FileSpreadsheet className="mr-3 w-7 h-7 text-blue-600" /> Hệ Thống Đối Chiếu Xác Linh Kiện
        </h1>
        <div className="flex items-center gap-3">
          {isLoadingData ? (
            <span className="bg-blue-50 text-blue-700 px-4 py-2 rounded-xl text-sm border border-blue-200 flex items-center font-bold shadow-sm">
              <Loader2 className="w-4 h-4 mr-2 animate-spin" /> Đang tải dữ liệu...
            </span>
          ) : (
            <>
              <span className="bg-gray-100 text-gray-600 px-4 py-2 rounded-xl text-sm border border-gray-200 flex items-center font-bold">
                Tổng dữ liệu: <strong className="text-gray-900 ml-1.5">{Number(excelData.length)}</strong> dòng
              </span>
              <button onClick={() => fetchGoogleSheetData(false)} className="bg-white border border-gray-300 text-gray-700 hover:bg-blue-50 hover:text-blue-700 hover:border-blue-300 px-4 py-2 rounded-xl text-sm font-bold shadow-sm transition-colors flex items-center cursor-pointer active:scale-95" title="Tải lại dữ liệu mới nhất từ Google Sheets"><RefreshCw className="w-4 h-4 mr-2" /> Làm mới dữ liệu</button>
            </>
          )}
        </div>
      </header>

      <main className="flex-1 flex flex-col p-4 gap-4 overflow-hidden w-full max-w-[1700px] mx-auto z-10">
        {hasConflictLock && (
          <div className="bg-amber-100 border-l-8 border-red-500 p-4 rounded-xl shadow-md flex items-center justify-between animate-pulse">
            <div className="flex items-center gap-4">
              <AlertTriangle className="w-8 h-8 text-red-600" />
              <div>
                <h4 className="font-black text-red-800 text-base">PHÁT HIỆN PHIẾU XUNG ĐỘT TRẠNG THÁI PHIÊN!</h4>
                <p className="text-sm text-red-700 font-semibold mt-1">Vui lòng XÓA dòng 'Không xác LK' bị trùng ở đầu lịch sử để mở khóa hệ thống quét tiếp.</p>
              </div>
            </div>
          </div>
        )}

        <div className="grid grid-cols-1 lg:grid-cols-12 gap-4 shrink-0">
          <div className="lg:col-span-8 bg-white p-4 rounded-2xl shadow-sm border border-gray-200 flex flex-col relative z-20">
            <div className="flex justify-between items-center mb-3 border-b border-gray-100 pb-3">
              <div className="flex items-center gap-4">
                <h2 className="font-bold text-gray-800 flex items-center text-base"><BarChart2 className="w-5 h-5 mr-2 text-blue-500" /> Thống kê:</h2>
                <div className="relative" ref={whDropdownRef}>
                  <button type="button" onClick={() => setIsWhDropdownOpen(!isWhDropdownOpen)} className="border-2 border-blue-200 rounded-xl px-4 py-2 text-sm font-bold bg-blue-50 text-blue-800 min-w-[280px] flex items-center justify-between cursor-pointer">
                    <span className="truncate">{String(selectedWarehouse) || "-- Chọn TTBH --"}</span>
                    <Search className="w-4 h-4 ml-2 text-blue-600 opacity-70 shrink-0" />
                  </button>
                  {isWhDropdownOpen && (
                    <div className="absolute left-0 mt-2 w-full bg-white border-2 border-gray-200 rounded-xl shadow-2xl z-50 p-2.5 max-h-[300px] flex flex-col min-w-[300px]">
                      <div className="relative mb-2 flex items-center">
                        <Search className="absolute left-3 w-4 h-4 text-gray-400 pointer-events-none" />
                        <input type="text" placeholder="Tìm trạm nhanh..." value={whSearchQuery} onChange={(e) => setWhSearchQuery(e.target.value)} className="w-full text-sm border border-gray-300 rounded-lg pl-9 pr-3 py-2 font-semibold" autoFocus />
                        {whSearchQuery && <button type="button" onClick={() => setWhSearchQuery('')} className="absolute right-3 text-gray-400"><X className="w-4 h-4" /></button>}
                      </div>
                      <div className="overflow-y-auto flex-1 divide-y divide-gray-50">
                        {filteredWarehouses.map((wh, idx) => (
                          <button key={idx} type="button" onClick={() => { setSelectedWarehouse(String(wh)); setIsWhDropdownOpen(false); setWhSearchQuery(''); }} className={`w-full text-left px-3 py-2.5 text-sm font-bold flex items-center justify-between ${String(wh) === String(selectedWarehouse) ? 'bg-blue-100 text-blue-800' : 'hover:bg-gray-100'}`}>
                            <span>{String(wh)}</span>
                            {String(wh) === String(selectedWarehouse) && <CheckCircle className="w-4 h-4 text-blue-600" />}
                          </button>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              </div>
              {selectedWarehouse && <button onClick={() => setShowDetailedProgressModal(true)} className="text-sm font-bold bg-gray-50 text-gray-700 px-4 py-2 rounded-xl border border-gray-300 hover:bg-blue-50 flex items-center shadow-sm cursor-pointer active:scale-95"><ClipboardList className="w-4 h-4 mr-2" /> DS Chưa Scan</button>}
            </div>

            <div className="flex flex-wrap items-center justify-between gap-6 mt-1">
              <div className="flex-1 min-w-[250px] max-w-lg">
                <table className="w-full border-collapse text-center text-sm rounded-xl overflow-hidden border border-gray-200">
                  <thead>
                    <tr className="bg-gray-100">
                      <th className="p-2 font-extrabold border-b text-xs">Loại</th>
                      <th className="p-2 text-emerald-800 font-extrabold border-b text-xs">IW</th>
                      <th className="p-2 text-blue-800 font-extrabold border-b text-xs">OOW</th>
                      <th className="p-2 text-rose-700 font-extrabold border-b text-xs">"KMH"</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y bg-white">
                    {['LCD', 'MAIN', 'OTHERS'].map(type => (
                      <tr key={type} className="hover:bg-gray-50">
                        <td className="p-2 font-bold border-r">{type}</td>
                        <td className="p-2 font-bold text-emerald-700 border-r">{Number(stats.scanned.IW[type])} / {Number(stats.totals.IW[type])}</td>
                        <td className="p-2 font-bold text-blue-700 border-r">{Number(stats.scanned.OOW[type])} / {Number(stats.totals.OOW[type])}</td>
                        <td className="p-2 font-bold text-rose-600">{Number(stats.totals.KMH[type]) > 0 ? `${Number(stats.scanned.KMH[type])} / ${Number(stats.totals.KMH[type])}` : '-'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="flex-1 min-w-[200px] px-5 border-l flex flex-col justify-center">
                {renderProgressBar(stats.scanned.IW.total, stats.totals.IW.total, 'Tiến độ IW', 'IW')}
                {renderProgressBar(stats.scanned.OOW.total, stats.totals.OOW.total, 'Tiến độ OOW', 'OOW')}
              </div>
            </div>
          </div>

          <div className="lg:col-span-4 bg-white p-4 rounded-2xl shadow-sm border border-gray-200 flex flex-col gap-4 justify-center z-10">
            <div className="flex-1 flex flex-col justify-center bg-gray-50/80 p-3.5 rounded-xl border">
              <h2 className="font-extrabold text-gray-800 mb-3 flex items-center text-xs uppercase tracking-widest"><Save className="w-4 h-4 mr-2 text-indigo-500" /> Quản lý Phiên</h2>
              <div className="grid grid-cols-2 gap-3">
                <button onClick={handleSaveSession} disabled={hasConflictLock || hasErrorLock} className={`text-sm font-bold py-2.5 rounded-lg flex items-center justify-center border-2 ${hasConflictLock || hasErrorLock ? 'text-gray-400 cursor-not-allowed' : 'bg-white border-indigo-300 text-indigo-800 hover:bg-indigo-50 active:scale-95'}`}><Save className="w-4 h-4 mr-2" /> Lưu Phiên</button>
                <label className="text-sm font-bold bg-white border-2 border-amber-300 text-amber-800 py-2.5 rounded-lg flex items-center justify-center cursor-pointer mb-0 active:scale-95"><UploadCloud className="w-4 h-4 mr-2" /> Nạp Phiên<input type="file" accept=".json" onChange={handleLoadSession} className="hidden" /></label>
              </div>
            </div>
            <div className="flex-1 flex flex-col justify-center bg-gray-50/80 p-3.5 rounded-xl border">
              <h2 className="font-extrabold text-gray-800 mb-3 flex items-center text-xs uppercase tracking-widest"><Download className="w-4 h-4 mr-2 text-emerald-500" /> Xuất Dữ Liệu</h2>
              <div className="grid grid-cols-2 gap-3 mb-3">
                <button onClick={handleExportExcel} disabled={hasConflictLock || hasErrorLock} className={`text-sm font-bold py-2.5 rounded-lg flex items-center justify-center border-2 ${hasConflictLock || hasErrorLock ? 'text-gray-400 cursor-not-allowed' : 'bg-white border-emerald-300 text-emerald-700 active:scale-95'}`}><Download className="w-4 h-4 mr-2" /> Tải TTBH</button>
                <button onClick={handleExportAllExcel} className="text-sm font-bold bg-white border-2 border-blue-300 text-blue-700 py-2.5 rounded-lg flex items-center justify-center active:scale-95"><Download className="w-4 h-4 mr-2" /> Tải TẤT CẢ</button>
              </div>
              <button onClick={handleUploadToDrive} disabled={isUploading || hasConflictLock || hasErrorLock} className="text-sm font-black bg-emerald-600 text-white border-2 border-emerald-700 py-3 rounded-xl flex items-center justify-center uppercase tracking-wide shadow-md active:scale-95 disabled:active:scale-100"><UploadCloud className="w-5 h-5 mr-2" /> Đẩy báo cáo lên Drive</button>
            </div>
          </div>
        </div>

        <div className={`bg-white rounded-2xl shadow-sm border flex flex-col flex-1 min-h-0 overflow-hidden transition-all duration-300 ${hasErrorLock || hasConflictLock ? 'border-red-400 ring-4 ring-red-100' : 'border-gray-200'}`}>
          <div className={`p-4 border-b flex gap-5 items-center shrink-0 ${hasErrorLock || hasConflictLock ? 'bg-red-50' : 'bg-gray-50'}`}>
            <div className="font-extrabold text-gray-800 text-base flex items-center">Lịch sử đối chiếu<span className="bg-white border border-gray-300 text-gray-700 px-3 py-1 rounded-full text-sm ml-3 font-mono">{Number(currentDisplayedRecords.length)}</span></div>
            <div className="flex-1 relative flex items-center max-w-2xl ml-5">
               {hasErrorLock || hasConflictLock ? <span className="text-red-700 font-black text-sm bg-red-200 px-3 py-1 rounded flex items-center border border-red-300 shadow-sm"><AlertTriangle className="w-4 h-4 mr-2" /> HỆ THỐNG ĐANG KHÓA</span> : <span className="absolute left-4 text-gray-400"><ScanLine className="w-5 h-5" /></span>}
               <input ref={mainInputRef} type="text" maxLength={100} value={scanInput} onChange={(e) => setScanInput(e.target.value)} onKeyDown={handleMainScan} disabled={!selectedWarehouse || hasErrorLock || hasConflictLock || isLoadingData} className={`w-full rounded-xl pl-11 pr-5 py-3 text-base font-mono border-2 ${hasErrorLock || hasConflictLock ? 'border-red-400 bg-red-100 text-red-800 font-bold' : 'border-blue-300 bg-white text-gray-900 font-semibold'}`} placeholder="Bắn mã vạch vào đây..." />
            </div>
            <button onClick={() => setConfirmDialog({ title: "Làm mới danh sách", message: "Xóa toàn bộ lịch sử quét của trạm này?", onConfirm: () => setScannedRecords(prev => prev.filter(r => String(r.ttbh) !== String(selectedWarehouse))) })} className="text-sm font-bold bg-white border border-gray-300 text-gray-600 px-4 py-2 rounded-xl flex items-center shadow-sm cursor-pointer active:scale-95"><Trash2 className="w-4 h-4 mr-2" /> Làm mới danh sách</button>
          </div>
          
          <div className="flex-1 overflow-y-auto">
            <table className="w-full text-left border-collapse whitespace-nowrap relative">
              <thead className="bg-white sticky top-0 z-10 shadow-sm border-b">
                <tr className="bg-gray-50 text-sm font-extrabold text-gray-700 uppercase">
                  <th className="p-3 px-5 w-[140px] max-w-[140px]">Cột Scan QR</th>
                  <th className="p-3 px-5 w-[220px]">Trạng Thái</th>
                  <th className="p-3 px-5">Số RO</th>
                  <th className="p-3 px-5">Mã LK</th>
                  <th className="p-3 px-5">Tên LK</th>
                  <th className="p-3 px-5">Model</th>
                  <th className="p-3 px-3">Loại</th>
                  <th className="p-3 px-5">BH/DV</th>
                  <th className="p-3 px-5 text-center">Slg</th>
                  <th className="p-3 px-5">Remark</th>
                  <th className="p-3 px-5 text-center bg-gray-100">Xóa</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100">
                {sortedDisplayedRecords.length === 0 ? (
                  <tr>
                    <td colSpan="11" className="p-16 text-center text-gray-400 bg-gray-50/50">
                      <ScanLine className="w-10 h-10 mx-auto mb-4 opacity-30" /> 
                      <span className="font-semibold text-base">Chưa có dữ liệu scan nào ở TTBH này.</span>
                    </td>
                  </tr>
                ) : (
                  sortedDisplayedRecords.map((record) => {
                    const recordSp = String(record.sp || "").trim().toUpperCase();
                    const isConflictNoXac = conflictRecordsSPs.has(recordSp) && String(record.status) === "Không xác LK";
                    const isErrorRow = String(record.status) !== "Khớp, Trả Xác LK về" && String(record.status) !== "Không xác LK";
                    
                    return (
                      <tr key={record.id} className={`hover:bg-gray-50 ${isConflictNoXac ? 'bg-orange-100/90 font-bold border-y-2 border-red-500 animate-pulse' : isErrorRow ? 'bg-red-50/80' : ''}`}>
                        <td className="p-3 px-5 font-mono text-[14px] font-semibold text-gray-800">
                          <div className="w-[140px] overflow-hidden text-ellipsis whitespace-nowrap" title={String(record.rawScan)}>{String(record.rawScan)}</div>
                        </td>
                        <td className="p-3 px-5">{isConflictNoXac ? <span className="text-red-800 bg-red-200 border border-red-400 px-3 py-1.5 rounded-lg font-black text-sm flex items-center"><AlertTriangle className="w-4 h-4 mr-1.5" /> Trùng trạng thái</span> : renderStatusHTML(record.status)}</td>
                        <td className="p-3 px-5 font-mono font-bold text-blue-700">{record.soRO ? <a href={`https://gcsm-sg.oppoit.com/order/order-management/after-sales-order/${String(record.soRO)}/detail`} target="_blank" rel="noopener noreferrer">{String(record.soRO)}</a> : '-'}</td>
                        <td className="p-3 px-5 font-mono text-gray-600">{String(record.maLK)}</td>
                        <td className="p-3 px-5 max-w-[300px] truncate font-bold text-base">{String(record.tenLK)}</td>
                        <td className="p-3 px-5 text-gray-700 font-medium">{String(record.model)}</td>
                        <td className="p-3 px-3 text-sm font-bold text-gray-500">{String(record.phanLoai)}</td>
                        <td className="p-3 px-5 font-black text-sm">{String(record.bhDv)}</td>
                        <td className="p-3 px-5 font-black text-center text-base">{String(record.slg)}</td>
                        <td className="p-3 px-5 font-bold text-orange-600 text-sm">{String(record.remark)}</td>
                        <td className="p-2 text-center border-l">
                          <button 
                            type="button"
                            onClick={(e) => {
                              e.preventDefault();
                              e.stopPropagation();
                              handleDeleteRecord(record.id);
                            }} 
                            className={`p-2 rounded-lg cursor-pointer transition-all active:scale-95 ${
                              isConflictNoXac || isErrorRow 
                                ? 'bg-red-600 text-white hover:bg-red-700 shadow-md animate-pulse' 
                                : 'text-gray-400 hover:text-red-600 hover:bg-red-50'
                            }`}
                            title="Xóa dòng này"
                          >
                            <Trash2 className="w-5 h-5 mx-auto pointer-events-none" />
                          </button>
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
        </div>
      </main>
    </div>
  );
}