/* 帶入總務處提供的勞保／勞退保費表：從 Word 複製整張表格或貼上截圖，依姓名填入當月保費欄。 */
(() => {
  const h = React.createElement;
  const toNumber = text => {
    const value = String(text ?? '').replace(/[,\s元]/g, '');
    return /^\d+(\.\d+)?$/.test(value) ? Number(value) : null;
  };
  // 圖片辨識的數字常黏著句點、Word 換行符號等雜字，只取前面的數字
  const toOcrNumber = text => { const match = String(text ?? '').match(/^\D?(\d+)\D*$/); return match ? Number(match[1]) : null; };

  // 圖片辨識出的文字：去掉表格框線、全形數字轉半形、中文字間的空白合併、中文與數字分開
  const cleanOcrLine = line => line
    .replace(/[０-９]/g, digit => String.fromCharCode(digit.charCodeAt(0) - 0xFEE0))
    .replace(/[|｜¦\[\]{}_—。，、“”"'‘’«»¢]/g, ' ')
    .replace(/([一-鿿])\s+(?=[一-鿿])/g, '$1')
    .replace(/([一-鿿])(?=\d)|(\d)(?=[一-鿿])/g, '$1$2 ');

  // 去掉最後一位數（辨識時把 Word 的換行符號「↵」讀成數字）
  const dropLastDigit = value => value >= 10 ? Math.floor(value / 10) : value;

  // 用保費的計算關係檢查辨識結果：勞退＝級距×6%×天數/30；勞保個人負擔約為機關補助的 2/7
  const checkRow = (row, fromImage) => {
    const notes = [];
    if (row.pensionGrade && row.days) {
      const expected = Math.round(row.pensionGrade * 0.06 * row.days / 30);
      if (row.pensionPublic !== expected) {
        if (fromImage && dropLastDigit(row.pensionPublic) === expected) {
          notes.push('勞退公付 ' + row.pensionPublic + ' 疑為辨識錯誤，已改為 ' + expected);
          row.pensionPublic = expected;
        } else notes.push('勞退公付與「級距×6%×天數÷30＝' + expected + '」不同，請核對');
      }
    }
    const ratio = (self, pub) => self / pub;
    const reasonable = (self, pub) => ratio(self, pub) > 0.2 && ratio(self, pub) < 0.4;
    if (row.laborPublic >= 20 && !reasonable(row.laborSelf, row.laborPublic)) {
      if (fromImage && reasonable(dropLastDigit(row.laborSelf), row.laborPublic)) {
        notes.push('勞保自付 ' + row.laborSelf + ' 疑為辨識錯誤，已改為 ' + dropLastDigit(row.laborSelf));
        row.laborSelf = dropLastDigit(row.laborSelf);
      } else if (fromImage && reasonable(row.laborSelf, dropLastDigit(row.laborPublic))) {
        notes.push('勞保公付 ' + row.laborPublic + ' 疑為辨識錯誤，已改為 ' + dropLastDigit(row.laborPublic));
        row.laborPublic = dropLastDigit(row.laborPublic);
      } else notes.push('勞保自付與公付的比例不太對，請核對');
    }
    return { ...row, notes };
  };

  // 總務處表格欄位順序（姓名之後）：投保金額、勞保級距、勞保機關補助、勞保個人負擔、勞退級距、機關提繳6%、機關補助墊償、二代健保
  window.parseInsurancePaste = (text, names = []) => {
    const lines = String(text).replace(/\r/g, '').split('\n').map(line => line.includes('\t') ? line : cleanOcrLine(line));
    const period = String(text).replace(/\s+/g, '').match(/投保期間[:：]?(\d{1,2})(?!\d)/); // 圖片辨識常把「月」讀錯，只看數字
    const rows = [];
    lines.forEach((line, index) => {
      const tabbed = line.includes('\t');
      const cells = tabbed ? line.split('\t').map(cell => cell.trim())
        : line.trim().split(/\s+/).filter(cell => /[\d一-鿿]/.test(cell)); // 去掉只有標點的雜字
      // 先找本工具已有的姓名（辨識時姓名前後可能黏著別的字），找不到再找任何 2~4 個中文字
      let nameIndex = cells.findIndex(cell => names.includes(cell));
      if (nameIndex < 0) nameIndex = cells.findIndex(cell => names.some(name => cell.includes(name)));
      if (nameIndex >= 0) cells[nameIndex] = names.find(name => cells[nameIndex].includes(name));
      else nameIndex = cells.findIndex(cell => /^[一-鿿]{2,4}$/.test(cell));
      if (nameIndex < 0) return;
      let values;
      if (tabbed) {
        values = [1, 2, 3, 4, 5, 6, 7].map(offset => toNumber(cells[nameIndex + offset]));
      } else {
        // 沒有 Tab（從 PDF 複製或圖片辨識）：取姓名後連續的數字，空白的墊償欄無法分辨，只取前 6 個
        values = [];
        for (const cell of cells.slice(nameIndex + 1)) { const n = toOcrNumber(cell); if (n === null) break; values.push(n); }
        values = values.slice(0, 6);
      }
      const [insured, , laborPublic, laborSelf, pensionGrade, pension, arrears] = values;
      if (laborPublic == null && laborSelf == null && pension == null) return; // 表頭等非資料列
      // 投保天數在下一列同一欄（合併儲存格的第二列）
      const next = (lines[index + 1] || '').split(tabbed ? '\t' : /\s+/).map(cell => cell.trim());
      const days = next.some(cell => /^[一-鿿]{2,4}$/.test(cell)) ? null
        : tabbed ? toNumber(next[nameIndex + 1]) : next.map(toOcrNumber).find(n => n !== null) ?? null;
      rows.push(checkRow({
        name: cells[nameIndex], insured: insured ?? null, days, pensionGrade: pensionGrade ?? null,
        laborSelf: laborSelf ?? 0,
        laborPublic: (laborPublic ?? 0) + (arrears ?? 0),
        pensionPublic: pension ?? 0
      }, !tabbed));
    });
    return { month: period ? Number(period[1]) : null, rows };
  };

  const workedDays = (staff, year, month) => {
    let days = 0;
    for (let day = 1; day <= new Date(year, month, 0).getDate(); day++) {
      if (Number(staff.hours?.[year + '-' + month + '-' + day]) > 0) days++;
    }
    return days;
  };

  // 圖片文字辨識（Tesseract.js），第一次使用時才下載
  let tesseractLoading = null;
  const loadTesseract = () => tesseractLoading || (tesseractLoading = new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = 'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js';
    script.onload = () => resolve(window.Tesseract);
    script.onerror = () => { tesseractLoading = null; reject(new Error('辨識元件下載失敗，請確認網路連線。')); };
    document.head.appendChild(script);
  }));
  // 放大兩倍、轉黑白（灰色、黃色底色變白），再擦掉表格框線，辨識才讀得到表格裡的數字
  const prepareImage = file => new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => {
      const canvas = document.createElement('canvas');
      const width = canvas.width = image.width * 2, height = canvas.height = image.height * 2;
      const context = canvas.getContext('2d');
      context.imageSmoothingQuality = 'high';
      context.drawImage(image, 0, 0, width, height);
      URL.revokeObjectURL(image.src);
      const pixels = context.getImageData(0, 0, width, height), data = pixels.data;
      const dark = new Uint8Array(width * height);
      for (let i = 0; i < width * height; i++) dark[i] = 0.299 * data[i * 4] + 0.587 * data[i * 4 + 1] + 0.114 * data[i * 4 + 2] < 150 ? 1 : 0;
      const line = new Uint8Array(width * height), minRun = 80; // 連續 80 點以上的黑線視為框線
      const markRuns = (count, length, at) => {
        for (let a = 0; a < count; a++) {
          let run = 0;
          for (let b = 0; b <= length; b++) {
            if (b < length && dark[at(a, b)]) run++;
            else { if (run >= minRun) for (let k = b - run; k < b; k++) line[at(a, k)] = 1; run = 0; }
          }
        }
      };
      markRuns(height, width, (y, x) => y * width + x);
      markRuns(width, height, (x, y) => y * width + x);
      for (let i = 0; i < width * height; i++) {
        const value = dark[i] && !line[i] ? 0 : 255;
        data[i * 4] = data[i * 4 + 1] = data[i * 4 + 2] = value; data[i * 4 + 3] = 255;
      }
      context.putImageData(pixels, 0, 0);
      resolve(canvas);
    };
    image.onerror = () => reject(new Error('無法讀取這張圖片。'));
    image.src = URL.createObjectURL(file);
  });
  const recognizeImage = async file => {
    const Tesseract = await loadTesseract();
    const worker = await Tesseract.createWorker('chi_tra+eng');
    try {
      await worker.setParameters({ tessedit_pageseg_mode: '6', preserve_interword_spaces: '1' });
      return (await worker.recognize(await prepareImage(file))).data.text;
    } finally { await worker.terminate(); }
  };

  window.InsuranceImport = function InsuranceImport({ staffList, setStaffList, year, month }) {
    const [open, setOpen] = React.useState(false);
    const [text, setText] = React.useState('');
    const [ocrStatus, setOcrStatus] = React.useState('');
    const [edits, setEdits] = React.useState({}); // 預覽表中手動修改的金額 { 姓名: { 欄位: 金額 } }
    const names = staffList.map(staff => staff.name.trim()).filter(Boolean);
    const parsed = React.useMemo(() => window.parseInsurancePaste(text, names), [text, names.join()]);
    React.useEffect(() => setEdits({}), [text]);
    const readImage = async file => {
      setOcrStatus('辨識圖片中…（第一次使用需下載辨識資料，約 30 秒）');
      try { setText(await recognizeImage(file)); setOcrStatus('已辨識圖片。辨識可能有誤，請對照圖片核對預覽表的金額，錯的直接在表中修改。'); }
      catch (error) { setOcrStatus('圖片辨識失敗：' + (error.message || error)); }
    };
    const onPaste = event => {
      const file = [...(event.clipboardData?.files || [])].find(item => item.type.startsWith('image/'));
      if (!file) return;
      event.preventDefault();
      readImage(file);
    };
    const period = year + '-' + month;
    const matched = parsed.rows.map(row => ({ ...row, ...edits[row.name], staff: staffList.find(staff => staff.name.trim() === row.name) }));
    const usable = matched.filter(row => row.staff);
    const skipped = matched.filter(row => !row.staff).map(row => row.name);
    const fields = [['laborSelf', '勞保自付'], ['laborPublic', '勞保公付'], ['pensionPublic', '勞退公付']];
    const edit = (name, key, value) => {
      if (value !== '' && (!Number.isFinite(Number(value)) || Number(value) < 0)) return;
      setEdits(previous => ({ ...previous, [name]: { ...previous[name], [key]: value === '' ? 0 : Number(value) } }));
    };

    const apply = () => {
      if (parsed.month && parsed.month !== Number(month) &&
        !window.confirm('貼上的是 ' + parsed.month + ' 月保費，目前畫面是 ' + month + ' 月，確定要帶入 ' + month + ' 月？')) return;
      setStaffList(previous => previous.map(staff => {
        const row = usable.find(item => item.staff.id === staff.id);
        if (!row) return staff;
        return { ...staff, payroll: { ...staff.payroll, [period]: { ...staff.payroll?.[period],
          laborSelf: row.laborSelf, laborPublic: row.laborPublic, pensionPublic: row.pensionPublic } } };
      }));
      setText(''); setOcrStatus(''); setOpen(false);
      window.alert('已帶入 ' + usable.length + ' 人的勞保、勞退保費。');
    };
    const close = () => { setText(''); setOcrStatus(''); setOpen(false); };

    if (!open) return h('button', { type: 'button', className: 'insurance-import-toggle', onClick: () => setOpen(true) }, '帶入總務處保費');
    return h('div', { className: 'insurance-import' },
      h('p', null, 'Word 檔：選取整張表格（含「投保期間」那行更好）→ 複製 → 貼到下面。圖片：截圖後直接貼到下面，或按「選擇圖片」。'),
      h('textarea', { value: text, rows: 5, placeholder: '在這裡貼上保費表格或截圖', onChange: event => setText(event.target.value), onPaste, 'aria-label': '貼上保費表格' }),
      h('label', { className: 'insurance-file' }, '選擇圖片 ',
        h('input', { type: 'file', accept: 'image/*', onChange: event => { const file = event.target.files[0]; event.target.value = ''; if (file) readImage(file); } })),
      ocrStatus && h('p', { className: ocrStatus.includes('失敗') ? 'payroll-error' : 'payroll-explanation' }, ocrStatus),
      text && !ocrStatus.startsWith('辨識圖片中') && !parsed.rows.length && h('p', { className: 'payroll-error' }, '讀不到保費資料，請確認有複製到含姓名與金額的表格列。'),
      parsed.rows.length > 0 && h('div', { className: 'payroll-scroll' },
        parsed.month && h('p', { className: parsed.month === Number(month) ? 'payroll-explanation' : 'payroll-error' },
          '投保期間：' + parsed.month + ' 月' + (parsed.month === Number(month) ? '' : '（目前畫面是 ' + month + ' 月）')),
        h('table', { className: 'insurance-preview' },
          h('thead', null, h('tr', null, ['姓名', '投保天數', '本月出勤天數', ...fields.map(([, title]) => title)].map(title => h('th', { key: title }, title)))),
          h('tbody', null, usable.map(row => {
            const worked = workedDays(row.staff, year, month);
            const mismatch = row.days !== null && row.days !== worked;
            return h('tr', { key: row.name },
              h('td', null, row.name),
              h('td', { className: mismatch ? 'insurance-mismatch' : undefined }, row.days ?? '—'),
              h('td', { className: mismatch ? 'insurance-mismatch' : undefined }, worked),
              ...fields.map(([key, title]) => h('td', { key }, h('input', {
                type: 'number', min: 0, value: row[key], 'aria-label': row.name + ' ' + title,
                onChange: event => edit(row.name, key, event.target.value)
              }))));
          }))),
        usable.some(row => row.days !== null && row.days !== workedDays(row.staff, year, month)) &&
          h('p', { className: 'payroll-error' }, '紅字：投保天數與上方工時表的出勤天數不同，請確認。'),
        usable.filter(row => row.notes.length).map(row => h('p', { key: row.name, className: 'payroll-error' }, row.name + '：' + row.notes.join('；'))),
        skipped.length > 0 && h('p', { className: 'payroll-explanation' }, '本工具沒有這些人員，已略過：' + skipped.join('、'))),
      h('div', { className: 'payroll-actions' },
        h('button', { type: 'button', onClick: apply, disabled: !usable.length }, '帶入 ' + month + ' 月保費'),
        h('button', { type: 'button', onClick: close }, '取消')),
      h('p', { className: 'payroll-explanation' }, '對應方式：勞保個人負擔 → 勞保自付；勞保機關補助（＋機關補助墊償）→ 勞保公付；機關提繳 6% 勞退 → 勞退公付。二代健保欄不帶入。'));
  };
})();
