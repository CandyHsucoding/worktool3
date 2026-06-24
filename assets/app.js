const {
  useState,
  useEffect
} = React;
const Icon = ({
  name,
  size = 18,
  className = ""
}) => {
  const iconKey = name.split('-').map(part => part.charAt(0).toUpperCase() + part.slice(1)).join('');
  const iconSet = window.lucide && window.lucide.icons ? window.lucide.icons : {};
  const iconDefinition = iconSet[iconKey] || iconSet[name];
  const iconNode = Array.isArray(iconDefinition) ? iconDefinition : iconDefinition && Array.isArray(iconDefinition.iconNode) ? iconDefinition.iconNode : null;
  if (!iconNode) {
    return React.createElement("span", {
      className: className,
      style: {
        width: size,
        height: size,
        display: 'inline-block'
      }
    });
  }
  return React.createElement("svg", {
    xmlns: "http://www.w3.org/2000/svg",
    width: size,
    height: size,
    viewBox: "0 0 24 24",
    fill: "none",
    stroke: "currentColor",
    strokeWidth: "2",
    strokeLinecap: "round",
    strokeLinejoin: "round",
    className: className
  }, iconNode.map(([tag, attrs], index) => React.createElement(tag, {
    ...attrs,
    key: index
  })));
};
const TimesheetApp = () => {
  const today = new Date();
  const [year, setYear] = useState(today.getFullYear());
  const [month, setMonth] = useState(today.getMonth() + 1);
  const [showNames, setShowNames] = useState(true);
  const [showDownloadMenu, setShowDownloadMenu] = useState(false);
  const [showBackupModal, setShowBackupModal] = useState(false);
  const [backupString, setBackupString] = useState('');
  const defaultStaff = [{
    id: 1,
    name: '',
    hours: {}
  }];
  const [staffList, setStaffList] = useState(() => {
    const savedData = localStorage.getItem('timesheet_data_v2');
    return savedData ? JSON.parse(savedData) : defaultStaff;
  });
  useEffect(() => {
    localStorage.setItem('timesheet_data_v2', JSON.stringify(staffList));
  }, [staffList]);
  const getHolidayData = (y, m, d) => {
    const dateStr = `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const mmdd = `${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const fixedHolidays = {
      '01-01': '元旦',
      '02-28': '和平紀念',
      '04-04': '兒童節',
      '04-05': '清明節',
      '05-01': '勞動節',
      '10-10': '國慶日',
      '12-25': '聖誕節'
    };
    if (fixedHolidays[mmdd]) return fixedHolidays[mmdd];
    const holidays2026 = {
      '2026-02-14': '春節連假',
      '2026-02-15': '春節連假',
      '2026-02-16': '除夕',
      '2026-02-17': '初一',
      '2026-02-18': '初二',
      '2026-02-19': '初三',
      '2026-02-20': '初四',
      '2026-02-21': '初五',
      '2026-02-22': '春節連假',
      '2026-04-03': '補假',
      '2026-04-06': '補假',
      '2026-06-19': '端午節',
      '2026-09-25': '中秋節',
      '2026-10-09': '補假'
    };
    if (y === 2026) return holidays2026[dateStr] || null;
    return null;
  };
  const getDaysInMonth = (y, m) => {
    const days = new Date(y, m, 0).getDate();
    const result = [];
    for (let i = 1; i <= days; i++) {
      const date = new Date(y, m - 1, i);
      const dayOfWeek = date.getDay();
      result.push({
        date: i,
        dayOfWeek: dayOfWeek,
        isWeekend: dayOfWeek === 0 || dayOfWeek === 6,
        holiday: getHolidayData(y, m, i)
      });
    }
    return result;
  };
  const daysInMonth = getDaysInMonth(year, month);
  const weekDayNames = ['日', '一', '二', '三', '四', '五', '六'];
  const handleHoursChange = (staffId, day, value) => {
    const numValue = value === '' ? '' : parseFloat(value);
    setStaffList(staffList.map(staff => {
      if (staff.id === staffId) {
        const newHours = {
          ...staff.hours
        };
        const key = `${year}-${month}-${day}`;
        if (value === '') delete newHours[key];else newHours[key] = numValue;
        return {
          ...staff,
          hours: newHours
        };
      }
      return staff;
    }));
  };
  const calculateTotal = staff => {
    let total = 0;
    daysInMonth.forEach(dayObj => {
      const key = `${year}-${month}-${dayObj.date}`;
      if (staff.hours[key]) total += staff.hours[key];
    });
    return total;
  };
  const getProcessedContent = (isForExcel = false) => {
    const printableArea = document.getElementById('printable-area').cloneNode(true);
    printableArea.querySelectorAll('input').forEach(input => {
      const span = document.createElement('span');
      span.className = 'print-text-center';
      span.textContent = input.value;
      input.parentNode.appendChild(span);
      input.classList.add('print-hidden');
    });
    printableArea.querySelectorAll('.print-hidden-col').forEach(el => el.remove());
    printableArea.querySelectorAll('.print-hidden').forEach(el => el.remove());
    if (isForExcel) {
      printableArea.querySelectorAll('.w-name').forEach(el => {
        el.style.width = '80px';
        el.style.minWidth = '80px';
        el.removeAttribute('class');
      });
      printableArea.querySelectorAll('.w-day').forEach(el => {
        el.style.width = '25px';
        el.style.minWidth = '25px';
      });
      printableArea.querySelectorAll('.w-total').forEach(el => {
        el.style.width = '50px';
        el.style.minWidth = '50px';
      });
    }
    return printableArea.innerHTML;
  };
  const handlePrint = () => {
    const content = getProcessedContent(false);
    const printWindow = window.open('', '', 'height=800,width=1200');
    if (!printWindow) {
      alert("瀏覽器封鎖了列印視窗，請允許此網站開啟彈出視窗後再試一次。");
      return;
    }
    printWindow.document.write(`
                    <html>
                        <head>
                            <title>${year}年${month}月工時報表</title>
                            <style>
                                @page { size: landscape; margin: 0.5cm; }
                                body { font-family: sans-serif; padding: 10px; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
                                h2 { text-align: center; margin-bottom: 15px; }
                                table { width: 100%; border-collapse: collapse; table-layout: fixed; border: 1.5px solid black; }
                                th, td { border: 1px solid black; text-align: center; font-size: 11px; height: 32px; word-break: break-all; }
                                th { background-color: #f3f4f6 !important; }
                                .bg-red-50 { background-color: #fef2f2 !important; }
                                .text-red-600 { color: #dc2626 !important; }
                                .w-name { width: 150px; }
                                .w-total { width: 60px; }
                                .print-text-center { display: block; font-weight: bold; }
                            </style>
                        </head>
                        <body>
                            <h2>${year} 年 ${month} 月 臨時人員工時統計表</h2>
                            ${content}
                        </body>
                    </html>
                `);
    printWindow.document.close();
    printWindow.focus();
    setTimeout(() => {
      printWindow.print();
      printWindow.close();
    }, 600);
  };
  const exportFile = type => {
    const fileName = `${year}年${month}月工時報表_${new Date().getTime()}`;
    const isExcel = type === 'excel';
    if (type === 'pdf') {
      handlePrint();
      return;
    }
    const content = getProcessedContent(isExcel);
    const blobContent = `
                    <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:${isExcel ? 'excel' : 'word'}' xmlns='http://www.w3.org/TR/REC-html40'>
                    <head>
                        <meta charset='utf-8'>
                        <!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>工時報表</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->
                        <style>
                            table { border-collapse: collapse; table-layout: fixed; }
                            th, td { border: 1px solid black; text-align: center; font-size: 10pt; height: 25pt; }
                            th { background-color: #D3D3D3; }
                            .print-text-center { font-weight: bold; }
                        </style>
                    </head>
                    <body>
                        <h2>${year}年${month}月工時報表</h2>
                        ${content}
                    </body></html>`;
    const blob = new Blob([blobContent], {
      type: isExcel ? 'application/vnd.ms-excel' : 'application/msword'
    });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.href = url;
    link.download = isExcel ? `${fileName}.xls` : `${fileName}.doc`;
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    setTimeout(() => {
      URL.revokeObjectURL(url);
      link.remove();
    }, 0);
    setShowDownloadMenu(false);
  };
  const handleBackup = () => {
    const data = JSON.stringify(staffList);
    setBackupString(btoa(unescape(encodeURIComponent(data))));
    setShowBackupModal(true);
  };
  const handleRestore = () => {
    try {
      const data = decodeURIComponent(escape(atob(backupString)));
      const parsed = JSON.parse(data);
      if (Array.isArray(parsed)) {
        setStaffList(parsed);
        alert("資料還原成功！");
        setShowBackupModal(false);
      } else {
        throw new Error();
      }
    } catch (e) {
      alert("無效的備份字串，請檢查是否複製完整。");
    }
  };
  const copyBackupToClipboard = () => {
    const textArea = document.createElement("textarea");
    textArea.value = backupString;
    document.body.appendChild(textArea);
    textArea.select();
    document.execCommand('copy');
    document.body.removeChild(textArea);
    alert("備份碼已複製到剪貼簿！");
  };
  return React.createElement("div", {
    className: "min-h-screen p-4"
  }, React.createElement("div", {
    className: "max-w-[1700px] mx-auto bg-white shadow-2xl rounded-xl overflow-hidden border border-slate-200"
  }, React.createElement("div", {
    className: "p-6 bg-slate-900 text-white flex flex-col md:flex-row justify-between items-center gap-4 print-hidden"
  }, React.createElement("div", {
    className: "flex items-center gap-3"
  }, React.createElement("div", {
    className: "bg-indigo-500 p-2 rounded-lg"
  }, React.createElement(Icon, {
    name: "calendar",
    className: "text-white"
  })), React.createElement("h1", {
    className: "text-xl font-bold"
  }, "臨時人員工時管理系統")), React.createElement("div", {
    className: "flex items-center gap-3"
  }, React.createElement("button", {
    onClick: handleBackup,
    className: "flex items-center gap-1.5 bg-slate-700 hover:bg-slate-600 text-white px-3 py-1.5 rounded-lg text-sm transition border border-slate-600"
  }, React.createElement(Icon, {
    name: "database",
    size: 16
  }), " 資料備份/遷移"), React.createElement("div", {
    className: "flex items-center gap-2 bg-slate-800 p-1.5 rounded-lg border border-slate-700"
  }, React.createElement("select", {
    value: year,
    onChange: e => setYear(parseInt(e.target.value)),
    className: "bg-transparent px-3 py-1 font-bold outline-none text-white"
  }, [2025, 2026, 2027].map(y => React.createElement("option", {
    key: y,
    value: y,
    className: "text-black"
  }, y, " 年"))), React.createElement("select", {
    value: month,
    onChange: e => setMonth(parseInt(e.target.value)),
    className: "bg-transparent px-3 py-1 font-bold outline-none text-white"
  }, Array.from({
    length: 12
  }, (_, i) => i + 1).map(m => React.createElement("option", {
    key: m,
    value: m,
    className: "text-black"
  }, m, " 月")))))), React.createElement("div", {
    className: "p-4 flex flex-wrap justify-between items-center gap-3 border-b bg-white print-hidden"
  }, React.createElement("div", {
    className: "flex gap-2 relative"
  }, React.createElement("button", {
    onClick: () => setStaffList([...staffList, {
      id: Date.now(),
      name: '',
      hours: {}
    }]),
    className: "flex items-center gap-1.5 bg-emerald-600 text-white px-4 py-2 rounded-lg hover:bg-emerald-700 shadow-sm transition active:scale-95"
  }, React.createElement(Icon, {
    name: "plus"
  }), " 新增人員"), React.createElement("div", {
    className: "relative"
  }, React.createElement("button", {
    onClick: () => setShowDownloadMenu(!showDownloadMenu),
    className: "flex items-center gap-1.5 bg-blue-600 text-white px-4 py-2 rounded-lg hover:bg-blue-700 shadow-sm transition active:scale-95"
  }, React.createElement(Icon, {
    name: "download"
  }), " 下載匯出"), showDownloadMenu && React.createElement("div", {
    className: "absolute top-full left-0 mt-2 w-40 bg-white border rounded-lg shadow-xl z-[100] p-1"
  }, React.createElement("button", {
    onClick: () => exportFile('excel'),
    className: "w-full text-left px-4 py-2 hover:bg-slate-100 rounded flex items-center gap-2 text-sm text-slate-700"
  }, React.createElement(Icon, {
    name: "file-spreadsheet",
    size: 16,
    className: "text-emerald-600"
  }), " Excel 報表"), React.createElement("button", {
    onClick: () => exportFile('word'),
    className: "w-full text-left px-4 py-2 hover:bg-slate-100 rounded flex items-center gap-2 text-sm text-slate-700"
  }, React.createElement(Icon, {
    name: "file-text",
    size: 16,
    className: "text-blue-600"
  }), " Word 文件"), React.createElement("button", {
    onClick: () => exportFile('pdf'),
    className: "w-full text-left px-4 py-2 hover:bg-slate-100 rounded flex items-center gap-2 text-sm text-slate-700"
  }, React.createElement(Icon, {
    name: "file-type-2",
    size: 16,
    className: "text-red-600"
  }), " PDF 檔案"))), React.createElement("button", {
    onClick: () => setShowNames(!showNames),
    className: "flex items-center gap-1.5 bg-slate-100 text-slate-700 px-4 py-2 rounded-lg hover:bg-slate-200 border border-slate-200"
  }, React.createElement(Icon, {
    name: showNames ? "eye-off" : "eye"
  }), " ", showNames ? "隱藏姓名" : "顯示姓名"), React.createElement("button", {
    onClick: handlePrint,
    className: "flex items-center gap-1.5 bg-indigo-600 text-white px-4 py-2 rounded-lg hover:bg-indigo-700 shadow-md transition active:scale-95"
  }, React.createElement(Icon, {
    name: "printer"
  }), " 列印彩色表格")), React.createElement("div", {
    className: "text-xs text-slate-400 font-medium flex gap-4"
  }, React.createElement("span", {
    className: "flex items-center gap-1.5"
  }, React.createElement("div", {
    className: "w-3 h-3 bg-red-100 border border-red-200 rounded"
  }), " 國定假日/週末"), React.createElement("span", {
    className: "flex items-center gap-1.5"
  }, React.createElement("div", {
    className: "w-3 h-3 bg-white border border-slate-200 rounded"
  }), " 工作日"))), React.createElement("div", {
    id: "printable-area",
    className: "overflow-x-auto"
  }, React.createElement("table", {
    className: "w-full text-sm border-collapse table-fixed border-l-0 border-r-0"
  }, React.createElement("thead", {
    className: "bg-slate-50"
  }, React.createElement("tr", null, React.createElement("th", {
    className: "p-3 border border-slate-300 text-center w-[200px] sticky left-0 bg-slate-50 z-20 w-name"
  }, "姓名"), daysInMonth.map(day => React.createElement("th", {
    key: day.date,
    className: `p-1 border border-slate-300 text-center w-[30px] w-day ${day.isWeekend || day.holiday ? 'bg-red-50 text-red-600 font-bold' : 'bg-white font-normal'}`
  }, React.createElement("div", {
    className: "text-[11px]"
  }, day.date), React.createElement("div", {
    className: "text-[9px] uppercase"
  }, weekDayNames[day.dayOfWeek]), day.holiday && React.createElement("div", {
    className: "text-[8px] text-red-500 mt-0.5 leading-tight truncate"
  }, day.holiday))), React.createElement("th", {
    className: "p-3 border border-slate-300 text-center w-[80px] sticky right-0 bg-slate-50 z-20 text-indigo-600 font-bold w-total"
  }, "總計"), React.createElement("th", {
    className: "p-3 border border-slate-300 text-center w-[45px] print-hidden print-hidden-col bg-slate-50"
  }, "操作"))), React.createElement("tbody", null, staffList.map(staff => React.createElement("tr", {
    key: staff.id,
    className: "hover:bg-slate-50/50 transition-colors group"
  }, React.createElement("td", {
    className: "p-0 border border-slate-300 sticky left-0 bg-white z-10 w-name"
  }, React.createElement("input", {
    type: "text",
    value: showNames ? staff.name : '',
    onChange: e => setStaffList(staffList.map(s => s.id === staff.id ? {
      ...s,
      name: e.target.value
    } : s)),
    placeholder: showNames ? "請輸入姓名" : "",
    className: "w-full h-full bg-transparent px-3 py-3 focus:bg-indigo-50 focus:outline-none font-bold text-slate-800 text-base",
    readOnly: !showNames
  })), daysInMonth.map(day => {
    const key = `${year}-${month}-${day.date}`;
    return React.createElement("td", {
      key: day.date,
      className: `p-0 border border-slate-300 w-day ${day.isWeekend || day.holiday ? 'bg-red-50/30' : ''}`
    }, React.createElement("input", {
      type: "number",
      value: staff.hours[key] || '',
      onChange: e => handleHoursChange(staff.id, day.date, e.target.value),
      className: "w-full h-full bg-transparent text-center py-3 focus:bg-white focus:outline-none transition font-bold text-sm"
    }));
  }), React.createElement("td", {
    className: "p-0 border border-slate-300 text-center font-black text-indigo-700 bg-slate-50 sticky right-0 z-10 text-base w-total"
  }, calculateTotal(staff)), React.createElement("td", {
    className: "p-1 border border-slate-300 text-center print-hidden print-hidden-col"
  }, React.createElement("button", {
    onClick: () => {
      if (window.confirm("確定刪除此人員資料？")) setStaffList(staffList.filter(s => s.id !== staff.id));
    },
    className: "text-slate-300 hover:text-red-500 transition-colors p-2"
  }, React.createElement(Icon, {
    name: "trash-2",
    size: 16
  })))))))), showBackupModal && React.createElement("div", {
    className: "fixed inset-0 bg-black/60 backdrop-blur-sm z-[200] flex items-center justify-center p-4"
  }, React.createElement("div", {
    className: "bg-white rounded-2xl shadow-2xl w-full max-w-lg overflow-hidden border border-slate-200"
  }, React.createElement("div", {
    className: "p-4 bg-slate-900 text-white flex justify-between items-center"
  }, React.createElement("h3", {
    className: "font-bold flex items-center gap-2"
  }, React.createElement(Icon, {
    name: "database"
  }), " 資料備份與還原"), React.createElement("button", {
    onClick: () => setShowBackupModal(false)
  }, React.createElement(Icon, {
    name: "x"
  }))), React.createElement("div", {
    className: "p-6 space-y-4"
  }, React.createElement("p", {
    className: "text-sm text-slate-600"
  }, "請複製下方的文字碼進行備份，或將備份碼貼到此處進行還原："), React.createElement("textarea", {
    value: backupString,
    onChange: e => setBackupString(e.target.value),
    className: "w-full h-40 p-3 bg-slate-50 border border-slate-200 rounded-lg text-[10px] font-mono focus:outline-none focus:ring-2 focus:ring-indigo-500 overflow-y-auto break-all",
    placeholder: "在此貼上備份碼..."
  }), React.createElement("div", {
    className: "grid grid-cols-2 gap-3"
  }, React.createElement("button", {
    onClick: copyBackupToClipboard,
    className: "bg-indigo-600 text-white py-2 rounded-lg font-bold hover:bg-indigo-700 flex items-center justify-center gap-2"
  }, React.createElement(Icon, {
    name: "copy",
    size: 16
  }), " 複製備份碼"), React.createElement("button", {
    onClick: handleRestore,
    className: "bg-emerald-600 text-white py-2 rounded-lg font-bold hover:bg-emerald-700 flex items-center justify-center gap-2"
  }, React.createElement(Icon, {
    name: "upload",
    size: 16
  }), " 還原資料")), React.createElement("div", {
    className: "bg-amber-50 border border-amber-200 p-3 rounded-lg flex gap-3"
  }, React.createElement(Icon, {
    name: "alert-triangle",
    className: "text-amber-500 shrink-0"
  }), React.createElement("p", {
    className: "text-[11px] text-amber-700 leading-relaxed"
  }, "提示：還原操作將覆蓋目前的輸入資料。備份碼包含所有人員姓名及工時資訊，請妥善保存。"))))), React.createElement("div", {
    className: "p-6 bg-slate-50 text-[11px] text-slate-400 flex justify-between items-center border-t print-hidden"
  }, React.createElement("div", {
    className: "flex gap-4"
  }, React.createElement("span", null, "※ 本系統已更新 2026 年人事行政局連假邏輯"), React.createElement("span", null, "※ 資料會自動保存在您的瀏覽器中")), React.createElement("span", {
    className: "hidden md:inline font-mono uppercase text-indigo-400"
  }, "Precision Backup-Sync v2.6"))));
};
const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(React.createElement(TimesheetApp, null));
