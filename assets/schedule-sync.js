/* 從「送餐人員排班系統 v2」帶入當月班表：請假日改記在代班人員名下，帶入後仍可手動修改。 */
(() => {
  const h = React.createElement;
  const SCHEDULE_API = 'https://script.google.com/macros/s/AKfycbxDSYwI3Ho8E_7siAkiTKZ0NU_qHUL6YqRMZ3Ma_yTvq-uc88HtqiUI1gN9uO0uK1A/exec';
  const REGULARS = ['溫美萍', '郭恒妘'];
  const HOURS_PER_DAY = 2;

  // 依班表算出每人每天的工時；不修改傳入的 staffList，回傳新的陣列與摘要文字
  window.applyScheduleToStaff = (staffList, year, month, submissions) => {
    const submitted = REGULARS.filter(name => submissions[name] && Array.isArray(submissions[name].rows));
    if (!submitted.length) throw new Error(year + '年' + month + '月尚無人提交班表。');

    // 班表上的每一天：記錄每位正式人員當天是出勤 (null) 還是由誰代班
    const schedule = {};
    submitted.forEach(name => submissions[name].rows.forEach(row => {
      const day = Number(String(row.date).slice(8, 10));
      schedule[day] = schedule[day] || {};
      schedule[day][name] = row.status === '請假' ? String(row.substitute || '').trim() : null;
    }));

    const assigned = {};
    const add = (name, day) => {
      assigned[name] = assigned[name] || {};
      assigned[name][day] = (assigned[name][day] || 0) + HOURS_PER_DAY;
    };
    Object.keys(schedule).forEach(dayText => {
      const day = Number(dayText);
      REGULARS.forEach(name => {
        const substitute = schedule[day][name];
        if (substitute) add(substitute, day);
        else add(name, day); // 出勤，或此人尚未提交（以出勤計）
      });
    });

    const next = staffList.map(staff => ({ ...staff, hours: { ...staff.hours } }));
    const findOrAdd = name => {
      let staff = next.find(item => item.name.trim() === name);
      if (!staff) {
        staff = { id: 'schedule-' + name + '-' + Date.now(), name, hours: {} };
        next.push(staff);
      }
      return staff;
    };
    const daysInMonth = new Date(year, month, 0).getDate();
    const summary = [];

    // 正式人員：本月整月依班表重填（班表外的日期清空）
    REGULARS.forEach(name => {
      const staff = findOrAdd(name);
      for (let day = 1; day <= daysInMonth; day++) {
        staff.hours[year + '-' + month + '-' + day] = assigned[name]?.[day] || '';
      }
    });
    // 代班人員：只填代班當天，不動其他日期
    Object.keys(assigned).filter(name => !REGULARS.includes(name)).forEach(name => {
      const staff = findOrAdd(name);
      // 代班日以外的日期保留原本輸入
      Object.keys(assigned[name]).forEach(day => { staff.hours[year + '-' + month + '-' + day] = assigned[name][day]; });
    });

    const substitutes = Object.keys(assigned).filter(name => !REGULARS.includes(name));
    REGULARS.filter(name => assigned[name]).concat(substitutes).forEach(name => {
      const days = Object.keys(assigned[name]).length;
      const hours = Object.values(assigned[name]).reduce((total, value) => total + value, 0);
      summary.push('・' + name + '：' + days + ' 天，' + hours + ' 小時' + (REGULARS.includes(name) ? '' : '（代班）'));
    });
    REGULARS.filter(name => !submitted.includes(name)).forEach(name => summary.push('・' + name + ' 尚未提交班表，暫以全部出勤計算'));

    return { staffList: next, summary: '已帶入 ' + year + '年' + month + '月班表：\n' + summary.join('\n') + '\n\n臨時請假請直接修改格子。' };
  };

  window.ScheduleImport = function ScheduleImport({ staffList, setStaffList, year, month, Icon }) {
    const [busy, setBusy] = React.useState(false);
    const run = async () => {
      const message = '從送餐人員排班系統帶入 ' + year + '年' + month + '月班表：\n' +
        '・溫美萍、郭恒妘本月工時會依班表重新填入（會覆蓋本月已輸入的數字）\n' +
        '・代班人員在代班日填入 ' + HOURS_PER_DAY + ' 小時\n\n帶入後仍可手動修改，確定帶入？';
      if (!window.confirm(message)) return;
      setBusy(true);
      try {
        const response = await fetch(SCHEDULE_API + '?action=month&year=' + year + '&month=' + month);
        const result = await response.json();
        if (result.status !== 'success') throw new Error(result.message || '讀取失敗');
        const applied = window.applyScheduleToStaff(staffList, year, month, result.data.submissions || {});
        setStaffList(applied.staffList);
        window.alert(applied.summary);
      } catch (error) {
        window.alert('無法帶入班表：' + (error.message || error));
      } finally {
        setBusy(false);
      }
    };
    return h('button', {
      type: 'button',
      onClick: run,
      disabled: busy,
      className: 'flex items-center gap-1.5 bg-amber-500 text-white px-4 py-2 rounded-lg hover:bg-amber-600 shadow-sm transition active:scale-95 disabled:opacity-60'
    }, h(Icon, { name: 'calendar-check' }), busy ? ' 讀取班表中…' : ' 帶入排班表');
  };
})();
