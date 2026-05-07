// ==================== MENU ====================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📅 บันทึกวันลา')
    .addItem('📋 เปิดแผงบันทึกวันลา', 'showSidebar')
    .addSeparator()
    .addItem('⚙️ ตั้งค่าสิทธิ์วันลา', 'showSettingsDialog')
    .addItem('🎯 คำนวณโบนัสในชีต', 'calculateAllBonuses')
    .addItem('🔧 สร้างหัวตารางใหม่', 'setupHeaders')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📊 ประวัติโบนัส')
      .addItem('ดูประวัติการคำนวณ', 'showCalculationHistory')
      .addItem('ล้างประวัติ', 'clearCalculationHistory'))
    .addToUi();
}

// ==================== WEB APP ENTRY POINT ====================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('📅 บันทึกวันลา & โบนัส')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==================== SIDEBAR ====================
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('📅 บันทึกวันลา & โบนัส')
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ==================== SETTINGS DIALOG ====================
function showSettingsDialog() {
  const html = HtmlService.createHtmlOutputFromFile('SettingsDialog')
    .setWidth(400)
    .setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, '⚙️ ตั้งค่าสิทธิ์วันลา');
}

// ==================== SETUP HEADERS ====================
function setupHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('data');
  if (!sheet) sheet = ss.insertSheet('data');

  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert('ยืนยัน', 'จะสร้างหัวตารางใหม่ในชีต "data" ใช่ไหม?\n(แถว 1-3 จะถูกเขียนทับ)', ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  const headers = [
    'วันที่ลา','ลาป่วย','ลากิจจ่าย','ลากิจหัก','ลาพักร้อน','หมายเหตุ',
    'ประเภทการลา','สิทธิ์ลาป่วย','สิทธิ์ลากิจจ่าย','สิทธิ์ลากิจหัก','สิทธิ์ลาพักร้อน',
    'ใช้ไปลาป่วย','ใช้ไปลากิจจ่าย','ใช้ไปลากิจหัก','ใช้ไปลาพักร้อน'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1a73e8').setFontColor('white')
    .setFontWeight('bold').setHorizontalAlignment('center');

  sheet.getRange('G2').setValue('📊 จำนวนสิทธิ์');
  sheet.getRange('H2').setValue(30);
  sheet.getRange('I2').setValue(6);
  sheet.getRange('J2').setValue(999);
  sheet.getRange('K2').setValue(6);
  sheet.getRange('G2:K2').setBackground('#e8f0fe').setFontWeight('bold');

  sheet.getRange('G3').setValue('📈 ใช้ไปแล้ว');
  sheet.getRange('L3').setFormula('=IFERROR(SUM(B4:B),0)');
  sheet.getRange('M3').setFormula('=IFERROR(SUM(C4:C),0)');
  sheet.getRange('N3').setFormula('=IFERROR(SUM(D4:D),0)');
  sheet.getRange('O3').setFormula('=IFERROR(SUM(E4:E),0)');
  sheet.getRange('H3').setFormula('=L3');
  sheet.getRange('I3').setFormula('=M3');
  sheet.getRange('J3').setFormula('=N3');
  sheet.getRange('K3').setFormula('=O3');
  sheet.getRange('G3:K3').setBackground('#fce8e6').setFontWeight('bold');

  sheet.getRange('A4:A').setNumberFormat('dd/MM/yyyy');
  sheet.setColumnWidths(1, 15, 100);
  sheet.setColumnWidth(6, 160);
  sheet.setFrozenRows(3);

  ui.alert('✅ สร้างหัวตารางเรียบร้อยแล้ว!');
}

// ==================== GET QUOTA ====================
function getQuota() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data');
  if (!sheet) return { sick: 30, paidLeave: 6, unpaidLeave: 999, vacation: 6 };
  return {
    sick:       Number(sheet.getRange('H2').getValue()) || 30,
    paidLeave:  Number(sheet.getRange('I2').getValue()) || 6,
    unpaidLeave:Number(sheet.getRange('J2').getValue()) || 999,
    vacation:   Number(sheet.getRange('K2').getValue()) || 6,
  };
}

// ==================== SAVE QUOTA ====================
function saveQuota(quota) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data');
  if (!sheet) throw new Error('ไม่พบชีต "data"');
  sheet.getRange('H2').setValue(Number(quota.sick));
  sheet.getRange('I2').setValue(Number(quota.paidLeave));
  sheet.getRange('J2').setValue(Number(quota.unpaidLeave));
  sheet.getRange('K2').setValue(Number(quota.vacation));
  return '✅ บันทึกสิทธิ์เรียบร้อยแล้ว';
}

// ==================== GET SUMMARY ====================
function getSummary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data');
  if (!sheet) return null;

  const quota = getQuota();
  const lastRow = sheet.getLastRow();
  let sick = 0, paid = 0, unpaid = 0, vacation = 0;

  if (lastRow >= 4) {
    const data = sheet.getRange(4, 2, lastRow - 3, 4).getValues();
    data.forEach(row => {
      sick     += Number(row[0]) || 0;
      paid     += Number(row[1]) || 0;
      unpaid   += Number(row[2]) || 0;
      vacation += Number(row[3]) || 0;
    });
  }

  return {
    quota,
    used:      { sick, paid, unpaid, vacation },
    remaining: {
      sick:     quota.sick - sick,
      paid:     quota.paidLeave - paid,
      unpaid:   quota.unpaidLeave - unpaid,
      vacation: quota.vacation - vacation,
    }
  };
}

// ==================== GET RECENT LEAVES ====================
function getRecentLeaves(limit) {
  limit = limit || 5;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data');
  if (!sheet || sheet.getLastRow() < 4) return [];

  const lastRow = sheet.getLastRow();
  const startRow = Math.max(4, lastRow - limit + 1);
  const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 7).getValues();

  return data.reverse().map(row => {
    const date = row[0] ? Utilities.formatDate(new Date(row[0]), 'Asia/Bangkok', 'dd/MM/yyyy') : '';
    const sick = Number(row[1]), paid = Number(row[2]), unpaid = Number(row[3]), vacation = Number(row[4]);
    let type = '', days = 0;
    if (sick)     { type = 'ลาป่วย';     days = sick; }
    if (paid)     { type = 'ลากิจจ่าย'; days = paid; }
    if (unpaid)   { type = 'ลากิจหัก';   days = unpaid; }
    if (vacation) { type = 'ลาพักร้อน';  days = vacation; }
    return { date, type, days, note: row[5] || '', typeLabel: row[6] || '' };
  }).filter(r => r.date);
}

// ==================== SAVE LEAVE ====================
function saveLeave(formData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data');
  if (!sheet) throw new Error('ไม่พบชีต "data" กรุณารัน "สร้างหัวตารางใหม่" ก่อน');

  const parts = formData.date.split('-');
  const leaveDate = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  const days = Number(formData.days) || 1;
  const typeMap = { sick: 2, paid: 3, unpaid: 4, vacation: 5 };
  const typeLabel = { sick: 'ลาป่วย', paid: 'ลากิจจ่าย', unpaid: 'ลากิจหัก', vacation: 'ลาพักร้อน' };

  let durLabel = ' (เต็มวัน)';
  if (days === 0.5)  durLabel = ' (ครึ่งวัน)';
  else if (days < 1) durLabel = ` (${days * 8} ชั่วโมง)`;

  const nextRow = Math.max(sheet.getLastRow() + 1, 4);
  sheet.getRange(nextRow, 1).setValue(leaveDate).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(nextRow, typeMap[formData.type]).setValue(days);
  sheet.getRange(nextRow, 6).setValue(formData.note || '');
  sheet.getRange(nextRow, 7).setValue(typeLabel[formData.type] + durLabel);

  const rowColors = { sick:'#fce8e6', paid:'#e6f4ea', unpaid:'#fef7e0', vacation:'#e8f0fe' };
  sheet.getRange(nextRow, 1, 1, 15).setBackground(rowColors[formData.type]);

  SpreadsheetApp.flush();
  return '✅ บันทึกการลาเรียบร้อยแล้ว';
}

// ==================== BONUS: CALCULATE ====================
function calculateBonus(usedLeaveDays) {
  let bonusPercentage, bonusLevel;
  if      (usedLeaveDays <= 6)  { bonusPercentage = 100; bonusLevel = 'A'; }
  else if (usedLeaveDays <= 8)  { bonusPercentage = 75;  bonusLevel = 'B'; }
  else if (usedLeaveDays <= 10) { bonusPercentage = 50;  bonusLevel = 'C'; }
  else                          { bonusPercentage = 0;   bonusLevel = 'D'; }

  const descriptions = {
    A: 'ยอดเยี่ยม! รวมวันลาไม่เกิน 6 วัน ได้โบนัสเต็มจำนวน',
    B: 'ดี! รวมวันลาไม่เกิน 8 วัน ได้โบนัส 75%',
    C: 'พอใช้! รวมวันลาไม่เกิน 10 วัน ได้โบนัสครึ่งหนึ่ง',
    D: 'ต้องการปรับปรุง! รวมวันลาเกิน 10 วัน ไม่ได้โบนัส'
  };
  return { totalUsedDays: usedLeaveDays, bonusPercentage, bonusLevel, description: descriptions[bonusLevel] };
}

// ==================== BONUS: GET BONUS DATA FROM SHEET ====================
function getBonusFromSheet() {
  const summary = getSummary();
  if (!summary) return null;
  const { used, quota, remaining } = summary;
  const totalForBonus = used.sick + used.paid + used.unpaid;
  const bonusResult = calculateBonus(totalForBonus);
  return {
    used, quota, remaining, totalForBonus,
    bonusPercentage: bonusResult.bonusPercentage,
    bonusLevel:      bonusResult.bonusLevel,
    description:     bonusResult.description,
  };
}

// ==================== BONUS: SAVE CALCULATION HISTORY ====================
function saveBonusCalculation(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('บันทึกโบนัส') || ss.insertSheet('บันทึกโบนัส');

    if (sheet.getLastRow() === 0) {
      const headers = ['วันที่คำนวณ','ลาป่วย','ลากิจจ่าย','ลากิจหัก','ลาพักร้อน','รวมวันลา','%โบนัส','ระดับ'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers])
        .setFontWeight('bold').setBackground('#1a73e8').setFontColor('white');
    }

    const nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1, 1, 8).setValues([[
      new Date(),
      data.sickLeave, data.paidLeave, data.deductLeave, data.vacationLeave,
      data.totalLeaveForBonus,
      data.bonusPercentage + '%',
      data.bonusLevel
    ]]);
    sheet.getRange(nextRow, 1).setNumberFormat('dd/MM/yyyy HH:mm');
    sheet.autoResizeColumns(1, 8);

    return { success: true, message: '✅ บันทึกประวัติโบนัสเรียบร้อย', row: nextRow };
  } catch(e) {
    return { success: false, message: '❌ ' + e.toString() };
  }
}

// ==================== BONUS: GET HISTORY ====================
function getCalculationHistory() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('บันทึกโบนัส');
  if (!sheet || sheet.getLastRow() <= 1) return [];

  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues()
    .map(row => ({
      timestamp:       row[0] ? Utilities.formatDate(new Date(row[0]), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm') : '',
      sickLeave:       row[1], paidLeave:  row[2],
      deductLeave:     row[3], vacationLeave: row[4],
      totalLeave:      row[5], bonusPercentage: row[6],
      bonusLevel:      row[7]
    }))
    .reverse();
}

// ==================== BONUS: CLEAR HISTORY ====================
function clearCalculationHistory() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert('ล้างประวัติ', 'ต้องการล้างประวัติการคำนวณโบนัสทั้งหมดใช่ไหม?', ui.ButtonSet.YES_NO);
  if (res === ui.Button.YES) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('บันทึกโบนัส');
    if (sheet) { sheet.clear(); ui.alert('✅ ล้างประวัติเรียบร้อยแล้ว'); }
  }
}

// ==================== BONUS: CALCULATE ALL IN SHEET ====================
function calculateAllBonuses() {
  const data = getBonusFromSheet();
  if (!data) { SpreadsheetApp.getUi().alert('❌ ไม่พบข้อมูลในชีต "data"'); return; }
  const ui = SpreadsheetApp.getUi();
  ui.alert('🎯 ผลการคำนวณโบนัส',
    `รวมวันลา (ลาป่วย+ลากิจจ่าย+ลากิจหัก): ${data.totalForBonus} วัน\nระดับโบนัส: ${data.bonusLevel}\nเปอร์เซ็นต์: ${data.bonusPercentage}%\n\n${data.description}`,
    ui.ButtonSet.OK);
}

// ==================== SHOW CALCULATION HISTORY DIALOG ====================
function showCalculationHistory() {
  const history = getCalculationHistory();
  const ui = SpreadsheetApp.getUi();
  if (!history.length) { ui.alert('ยังไม่มีประวัติการคำนวณ'); return; }
  let msg = 'ประวัติ 5 รายการล่าสุด:\n\n';
  history.slice(0, 5).forEach((h, i) => {
    msg += `${i+1}. ${h.timestamp} — ระดับ ${h.bonusLevel} (${h.bonusPercentage})\n`;
  });
  ui.alert('📊 ประวัติการคำนวณโบนัส', msg, ui.ButtonSet.OK);
}

// ==================== GET ALL LEAVES (for edit/delete) ====================
function getAllLeaves() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data');
  if (!sheet || sheet.getLastRow() < 4) return [];

  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(4, 1, lastRow - 3, 7).getValues();

  return data.map((row, i) => ({
    rowIndex: i + 4, // แถวจริงใน Sheet
    date:     row[0] ? Utilities.formatDate(new Date(row[0]), 'Asia/Bangkok', 'dd/MM/yyyy') : '',
    sick:     Number(row[1]) || 0,
    paid:     Number(row[2]) || 0,
    unpaid:   Number(row[3]) || 0,
    vacation: Number(row[4]) || 0,
    note:     row[5] || '',
    typeLabel:row[6] || '',
  })).filter(r => r.date).reverse();
}

// ==================== UPDATE LEAVE ====================
function updateLeave(rowIndex, formData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data');
  if (!sheet) throw new Error('ไม่พบชีต "data"');

  const days = Number(formData.days) || 1;
  const typeMap = { sick: 2, paid: 3, unpaid: 4, vacation: 5 };
  const typeLabel = { sick: 'ลาป่วย', paid: 'ลากิจจ่าย', unpaid: 'ลากิจหัก', vacation: 'ลาพักร้อน' };

  let durLabel = ' (เต็มวัน)';
  if (days === 0.5)  durLabel = ' (ครึ่งวัน)';
  else if (days < 1) durLabel = ` (${days * 8} ชั่วโมง)`;

  // ล้างคอลัมน์ B-E ก่อน แล้วเขียนใหม่
  sheet.getRange(rowIndex, 2, 1, 4).clearContent();
  sheet.getRange(rowIndex, typeMap[formData.type]).setValue(days);
  sheet.getRange(rowIndex, 6).setValue(formData.note || '');
  sheet.getRange(rowIndex, 7).setValue(typeLabel[formData.type] + durLabel);

  const rowColors = { sick:'#fce8e6', paid:'#e6f4ea', unpaid:'#fef7e0', vacation:'#e8f0fe' };
  sheet.getRange(rowIndex, 1, 1, 15).setBackground(rowColors[formData.type]);

  SpreadsheetApp.flush();
  return '✅ แก้ไขรายการเรียบร้อยแล้ว';
}

// ==================== DELETE LEAVE ====================
function deleteLeave(rowIndex) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data');
  if (!sheet) throw new Error('ไม่พบชีต "data"');
  sheet.deleteRow(rowIndex);
  SpreadsheetApp.flush();
  return '✅ ลบรายการเรียบร้อยแล้ว';
}