/* 帶入總務處提供的勞保／勞退保費表：從 Word 複製整張表格貼上，依姓名填入當月保費欄。 */
(() => {
  const h = React.createElement;
  const toNumber = text => {
    const value = String(text ?? '').replace(/[,\s元]/g, '');
    return /^\d+(\.\d+)?$/.test(value) ? Number(value) : null;
  };

  // 總務處表格欄位順序（姓名之後）：投保金額、勞保級距、勞保機關補助、勞保個人負擔、勞退級距、機關提繳6%、機關補助墊償、二代健保
  window.parseInsurancePaste = text => {
    const lines = String(text).replace(/\r/g, '').split('\n');
    const period = String(text).match(/投保期間\s*[:：]\s*(\d{1,2})\s*月/);
    const rows = [];
    lines.forEach((line, index) => {
      const tabbed = line.includes('\t');
      const cells = tabbed ? line.split('\t').map(cell => cell.trim()) : line.trim().split(/\s+/);
      const nameIndex = cells.findIndex(cell => /^[一-鿿]{2,4}$/.test(cell));
      if (nameIndex < 0) return;
      let values;
      if (tabbed) {
        values = [1, 2, 3, 4, 5, 6, 7].map(offset => toNumber(cells[nameIndex + offset]));
      } else {
        // 沒有 Tab（例如從 PDF 複製）：取姓名後連續的數字，空白的墊償欄無法分辨，只取前 6 個
        values = [];
        for (const cell of cells.slice(nameIndex + 1)) { const n = toNumber(cell); if (n === null) break; values.push(n); }
        values = values.slice(0, 6);
      }
      const [insured, , laborPublic, laborSelf, , pension, arrears] = values;
      if (laborPublic == null && laborSelf == null && pension == null) return; // 表頭等非資料列
      // 投保天數在下一列同一欄（合併儲存格的第二列）
      const next = (lines[index + 1] || '').split(tabbed ? '\t' : /\s+/).map(cell => cell.trim());
      const days = next.some(cell => /^[一-鿿]{2,4}$/.test(cell)) ? null : toNumber(tabbed ? next[nameIndex + 1] : next.find(cell => toNumber(cell) !== null));
      rows.push({
        name: cells[nameIndex], insured: insured ?? null, days,
        laborSelf: laborSelf ?? 0,
        laborPublic: (laborPublic ?? 0) + (arrears ?? 0),
        pensionPublic: pension ?? 0
      });
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

  window.InsuranceImport = function InsuranceImport({ staffList, setStaffList, year, month }) {
    const [open, setOpen] = React.useState(false);
    const [text, setText] = React.useState('');
    const parsed = React.useMemo(() => window.parseInsurancePaste(text), [text]);
    const period = year + '-' + month;
    const matched = parsed.rows.map(row => ({ ...row, staff: staffList.find(staff => staff.name.trim() === row.name) }));
    const usable = matched.filter(row => row.staff);
    const skipped = matched.filter(row => !row.staff).map(row => row.name);

    const apply = () => {
      if (parsed.month && parsed.month !== Number(month) &&
        !window.confirm('貼上的是 ' + parsed.month + ' 月保費，目前畫面是 ' + month + ' 月，確定要帶入 ' + month + ' 月？')) return;
      setStaffList(previous => previous.map(staff => {
        const row = usable.find(item => item.staff.id === staff.id);
        if (!row) return staff;
        return { ...staff, payroll: { ...staff.payroll, [period]: { ...staff.payroll?.[period],
          laborSelf: row.laborSelf, laborPublic: row.laborPublic, pensionPublic: row.pensionPublic } } };
      }));
      setText(''); setOpen(false);
      window.alert('已帶入 ' + usable.length + ' 人的勞保、勞退保費。');
    };

    if (!open) return h('button', { type: 'button', className: 'insurance-import-toggle', onClick: () => setOpen(true) }, '帶入總務處保費');
    return h('div', { className: 'insurance-import' },
      h('p', null, '在總務處的保費 Word 檔中，選取整張表格（含「投保期間」那行更好）→ 複製 → 貼到下面。'),
      h('textarea', { value: text, rows: 5, placeholder: '在這裡貼上保費表格', onChange: event => setText(event.target.value), 'aria-label': '貼上保費表格' }),
      text && !parsed.rows.length && h('p', { className: 'payroll-error' }, '讀不到保費資料，請確認有複製到含姓名與金額的表格列。'),
      parsed.rows.length > 0 && h('div', { className: 'payroll-scroll' },
        parsed.month && h('p', { className: parsed.month === Number(month) ? 'payroll-explanation' : 'payroll-error' },
          '投保期間：' + parsed.month + ' 月' + (parsed.month === Number(month) ? '' : '（目前畫面是 ' + month + ' 月）')),
        h('table', { className: 'insurance-preview' },
          h('thead', null, h('tr', null, ['姓名', '投保天數', '本月出勤天數', '勞保自付', '勞保公付', '勞退公付'].map(title => h('th', { key: title }, title)))),
          h('tbody', null, usable.map(row => {
            const worked = workedDays(row.staff, year, month);
            const mismatch = row.days !== null && row.days !== worked;
            return h('tr', { key: row.name },
              h('td', null, row.name),
              h('td', { className: mismatch ? 'insurance-mismatch' : undefined }, row.days ?? '—'),
              h('td', { className: mismatch ? 'insurance-mismatch' : undefined }, worked),
              h('td', null, row.laborSelf), h('td', null, row.laborPublic), h('td', null, row.pensionPublic));
          }))),
        usable.some(row => row.days !== null && row.days !== workedDays(row.staff, year, month)) &&
          h('p', { className: 'payroll-error' }, '紅字：投保天數與上方工時表的出勤天數不同，請確認。'),
        skipped.length > 0 && h('p', { className: 'payroll-explanation' }, '本工具沒有這些人員，已略過：' + skipped.join('、'))),
      h('div', { className: 'payroll-actions' },
        h('button', { type: 'button', onClick: apply, disabled: !usable.length }, '帶入 ' + month + ' 月保費'),
        h('button', { type: 'button', onClick: () => { setText(''); setOpen(false); } }, '取消')),
      h('p', { className: 'payroll-explanation' }, '對應方式：勞保個人負擔 → 勞保自付；勞保機關補助（＋機關補助墊償）→ 勞保公付；機關提繳 6% 勞退 → 勞退公付。二代健保欄不帶入。'));
  };
})();
