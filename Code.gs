// ════════════════════════════════════════════════════════════
// Code.gs — Google Apps Script Backend
// Database Sheet: "Data base"
// Columns: A=User, B=Password, C=Plant Code, D=Plant Name,
//          E=Area, F=EN-ME, G=Plant E-mail
// ════════════════════════════════════════════════════════════

const SHEET_DB   = 'Data base'; // Sheet ข้อมูล User
const SHEET_DATA = 'Data';      // Sheet บันทึกผลตรวจ

function doGet(e) {
  try {
    const action = e.parameter.action;

    // ── Login ──────────────────────────────────────────────
    if (action === 'login') {
      const username = (e.parameter.username || '').trim();
      const password = (e.parameter.password || '').trim();
      const user = checkLogin(username, password);

      if (user) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'ok', user }))
          .setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'error', message: 'Invalid credentials' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ── Submit ─────────────────────────────────────────────
    if (action === 'submit') {
      saveToSheet(e.parameter);
      if (e.parameter.status === 'ABNORMAL') {
        sendAbnormalEmail(e.parameter, e.parameter.notifyEmail || '');
      }
      return ok();
    }

    // ── History ────────────────────────────────────────────
    if (action === 'history') {
      const rows = getHistory(
        e.parameter.machineId || '',
        parseInt(e.parameter.months) || 3
      );
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', rows }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ok('Machine Inspection API Ready');
  } catch(err) {
    Logger.log('doGet error: ' + err);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── POST: ใช้สำหรับแนบรูปภาพ (action=photos) ────────────────
function doPost(e) {
  try {
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    if (body.action === 'photos') {
      savePhotos(body);
    }
    return ok();
  } catch(err) {
    Logger.log('doPost error: ' + err);
    return ok();
  }
}

// ── Login จาก Sheet "Data base" ──────────────────────────────
// Columns: A=User, B=Password, C=Plant Code, D=Plant Name,
//          E=Area, F=EN-ME, G=Plant E-mail
function checkLogin(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DB);
  if (!sheet) {
    Logger.log('Sheet "Data base" not found');
    return null;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const rows = sheet.getRange(2, 1, lastRow - 1, 7).getValues();

  for (const row of rows) {
    const u = String(row[0]).trim(); // A = User
    const p = String(row[1]).trim(); // B = Password

    if (u === username && p === password) {
      return {
        username: u,
        password: p,
        plantCode: String(row[2]).trim(), // C = Plant Code
        plantName: String(row[3]).trim(), // D = Plant Name
        area: String(row[4]).trim(),      // E = Area
        enMe: String(row[5]).trim(),      // F = EN-ME
        email: String(row[6]).trim(),     // G = Plant E-mail
        // สำหรับแสดงผลในหน้าตรวจ
        name: u,
        // admin = username "admin" เท่านั้น จะเห็นเมนู "หรือเลือกจุดตรวจสอบ"
        role: u.toLowerCase() === 'admin' ? 'admin' : 'technician'
      };
    }
  }
  return null;
}

// ── บันทึกลง Sheet "Data" ────────────────────────────────────
function saveToSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_DATA);
  if (!sheet) sheet = ss.insertSheet(SHEET_DATA);

  // หัวตารางตามที่กำหนด — คอลัมน์รูป "รูปที่ N" จะถูกเพิ่มอัตโนมัติทีหลัง
  // ตามจำนวนรายการตรวจสอบจริงของแต่ละเครื่อง (ดู savePhotos())
  const HEADERS = [
    'วันเวลา','Machine ID','ชื่อเครื่องจักร','โรงงาน (FC)','หน่วยผลิต',
    'ภูมิภาค','ME ผู้ดูแล','ตำแหน่ง','ผู้ตรวจสอบ','ผลการตรวจ',
    'รายการปกติ','รายการผิดปกติ','รายการที่ผิดปกติ','หมายเหตุ'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1,1,1,HEADERS.length)
      .setBackground('#1B4F8A').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  const row = [
    data.timestamp || new Date().toLocaleString('th-TH'),   // วันเวลา
    data.machineId || '',                                    // Machine ID
    data.machineName || '',                                   // ชื่อเครื่องจักร
    data.plantCode || data.fc || '',                          // โรงงาน (FC)
    data.plantName || '',                                     // หน่วยผลิต
    data.area || data.region || '',                           // ภูมิภาค
    data.enMe || data.me || '',                                // ME ผู้ดูแล
    data.location || '',                                       // ตำแหน่ง
    data.inspector || '',                                      // ผู้ตรวจสอบ
    data.status || '',                                         // ผลการตรวจ
    data.normalCount || 0,                                     // รายการปกติ
    data.abnormalCount || 0,                                   // รายการผิดปกติ
    data.abnormalItems || '',                                  // รายการที่ผิดปกติ
    data.note || '',                                           // หมายเหตุ
  ];

  sheet.appendRow(row);

  // ระบายสีแถว
  const lastRow = sheet.getLastRow();
  const rowRange = sheet.getRange(lastRow, 1, 1, HEADERS.length);

  if (data.status === 'ABNORMAL') {
    rowRange.setBackground('#FEF2F2');
    sheet.getRange(lastRow, HEADERS.indexOf('รายการที่ผิดปกติ') + 1)
      .setBackground('#FCA5A5').setFontColor('#B91C1C').setFontWeight('bold');
  } else if (data.status === 'NORMAL') {
    rowRange.setBackground('#F0FDF4');
  }
}

// ── บันทึกรูปภาพ: อัปโหลดขึ้น Google Drive แล้วเขียนลิงก์ลงชีท ──
// body = { machineId, inspector, timestamp, totalChecks, photos: {checkId: dataURL, ...} }
function savePhotos(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DATA);
  if (!sheet || sheet.getLastRow() < 1) return;

  let lastCol = sheet.getLastColumn();
  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // ── ขยายคอลัมน์ "รูปที่ N" ให้พอกับจำนวนรายการตรวจสอบของเครื่องนี้ ──
  const photoKeys = Object.keys(body.photos || {}).map(Number).filter(n => !isNaN(n));
  const neededPhotoCols = Math.max(
    parseInt(body.totalChecks, 10) || 0,
    photoKeys.length ? Math.max(...photoKeys) + 1 : 0
  );
  const existingPhotoCols = headers.filter(h => /^รูปที่ \d+$/.test(h)).length;

  if (neededPhotoCols > existingPhotoCols) {
    for (let i = existingPhotoCols + 1; i <= neededPhotoCols; i++) {
      const colIdx = headers.length + 1;
      sheet.getRange(1, colIdx).setValue(`รูปที่ ${i}`)
        .setBackground('#1B4F8A').setFontColor('#fff').setFontWeight('bold');
      headers.push(`รูปที่ ${i}`);
    }
    lastCol = headers.length;
  }

  // ── หาแถวที่ตรงกับการส่งผลนี้ (Machine ID + วันเวลา + ผู้ตรวจสอบ ตรงกันล่าสุด) ──
  const iTime = headers.indexOf('วันเวลา');
  const iMachineId = headers.indexOf('Machine ID');
  const iInspector = headers.indexOf('ผู้ตรวจสอบ');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  let targetRow = -1;
  for (let r = values.length - 1; r >= 0; r--) {
    if (String(values[r][iMachineId]) === String(body.machineId || '') &&
        String(values[r][iInspector]) === String(body.inspector || '') &&
        String(values[r][iTime]) === String(body.timestamp || '')) {
      targetRow = r + 2; // แปลงเป็นเลขแถวจริงในชีท
      break;
    }
  }
  if (targetRow === -1) return; // ไม่พบแถวที่ตรง — ข้าม

  // ── อัปโหลดรูปแต่ละใบขึ้น Drive แล้วเขียนลิงก์ลงคอลัมน์ตามลำดับข้อ ──
  const folder = getOrCreatePhotoFolder();
  Object.keys(body.photos || {}).forEach(checkIdStr => {
    const checkId = parseInt(checkIdStr, 10);
    const dataUrl = body.photos[checkIdStr];
    if (!dataUrl || isNaN(checkId)) return;

    const url = uploadPhotoToDrive(folder, dataUrl, body.machineId || '', body.timestamp || '', checkId);
    if (!url) return;

    const colName = `รูปที่ ${checkId + 1}`;
    const colIdx = headers.indexOf(colName);
    if (colIdx >= 0) {
      sheet.getRange(targetRow, colIdx + 1).setValue(url);
    }
  });
}

// ── โฟลเดอร์เก็บรูปใน Google Drive (สร้างครั้งแรกอัตโนมัติ) ──
function getOrCreatePhotoFolder() {
  const FOLDER_NAME = 'Machine Inspection Photos';
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(FOLDER_NAME);
}

// ── แปลง base64 dataURL เป็นไฟล์ + อัปโหลด + เปิดสิทธิ์ดูลิงก์ ──
function uploadPhotoToDrive(folder, dataUrl, machineId, timestamp, checkId) {
  try {
    const m = String(dataUrl).match(/^data:(image\/\w+);base64,(.+)$/);
    if (!m) return '';
    const mimeType = m[1];
    const base64 = m[2];
    const bytes = Utilities.base64Decode(base64);
    const safeTs = String(timestamp).replace(/[^\d]/g, '') || String(Date.now());
    const fileName = `${machineId}_${safeTs}_check${checkId + 1}.jpg`;
    const blob = Utilities.newBlob(bytes, mimeType, fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch(err) {
    Logger.log('uploadPhotoToDrive error: ' + err);
    return '';
  }
}

// ── History ──────────────────────────────────────────────────
function getHistory(machineId, months) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DATA);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - months);

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1,1,1,lastCol).getValues()[0];
  const idx = name => headers.indexOf(name);

  const iTime      = idx('วันเวลา');
  const iMachineId = idx('Machine ID');
  const iMachName  = idx('ชื่อเครื่องจักร');
  const iInspector = idx('ผู้ตรวจสอบ');
  const iStatus    = idx('ผลการตรวจ');
  const iAbnItems  = idx('รายการที่ผิดปกติ');
  const iNote      = idx('หมายเหตุ');

  return sheet.getRange(2,1,sheet.getLastRow()-1,lastCol).getValues()
    .filter(r => {
      const d = r[iTime] instanceof Date ? r[iTime] : new Date(r[iTime]);
      const inRange = !isNaN(d) && d >= cutoff;
      const matchMachine = !machineId || String(r[iMachineId]).trim() === String(machineId).trim();
      return inRange && matchMachine;
    })
    .map(r => ({
      time: r[iTime] instanceof Date ? r[iTime].toLocaleString('th-TH') : String(r[iTime]),
      machineId: iMachineId >= 0 ? r[iMachineId] : '',
      machineName: iMachName >= 0 ? r[iMachName] : '',
      inspector: iInspector >= 0 ? r[iInspector] : '',
      status: iStatus >= 0 ? r[iStatus] : '',
      abnormalItems: iAbnItems >= 0 ? r[iAbnItems] : '',
      note: iNote >= 0 ? r[iNote] : '',
    }))
    .reverse();
}

// ── Email แจ้งเตือน ──────────────────────────────────────────
function sendAbnormalEmail(data, toEmail) {
  if (!toEmail) return;

  const subject = `⚠️ พบความผิดปกติ: ${data.plantName || data.plantCode}`;
  const abnList = (data.abnormalItems || '').split(',')
    .filter(s => s.trim())
    .map(s => `<li style="margin:4px 0">${s.trim()}</li>`).join('');

  const html = `
  <div style="font-family:sans-serif;max-width:580px;margin:0 auto">
    <div style="background:linear-gradient(135deg,#0F3260,#1B4F8A);color:#fff;padding:22px 24px;border-radius:10px 10px 0 0">
      <div style="font-size:11px;opacity:.75;margin-bottom:6px">ระบบตรวจสอบเครื่องจักร — แจ้งเตือนอัตโนมัติ</div>
      <div style="font-size:22px;font-weight:700">⚠️ พบความผิดปกติ</div>
      <div style="font-size:15px;margin-top:4px;opacity:.9">${data.plantName||''} (${data.plantCode||''})</div>
    </div>
    <div style="background:#fff;padding:22px 24px;border:1px solid #E5E7EB;border-top:none">
      <table style="width:100%;border-collapse:collapse;margin-bottom:16px">
        ${[
          ['Plant Code', data.plantCode||'—'],
          ['Plant Name', data.plantName||'—'],
          ['Area', data.area||'—'],
          ['EN-ME', data.enMe||'—'],
          ['ผู้ตรวจสอบ', data.inspector||'—'],
          ['วันเวลา', data.timestamp||'—'],
        ].map(([l,v])=>`<tr style="border-bottom:1px solid #F3F4F6">
          <td style="padding:9px 12px;background:#F9FAFB;font-weight:600;font-size:13px;color:#4B5563;width:40%">${l}</td>
          <td style="padding:9px 12px;font-size:13px">${v}</td></tr>`).join('')}
        <tr><td style="padding:9px 12px;background:#FEF2F2;font-weight:600;font-size:13px;color:#B91C1C">ผลการตรวจ</td>
        <td style="padding:9px 12px;background:#FEF2F2;font-weight:700;color:#B91C1C">❌ ABNORMAL</td></tr>
      </table>
      ${abnList?`<div style="background:#FEF2F2;border-left:4px solid #B91C1C;border-radius:4px;padding:14px 16px;margin-bottom:16px">
        <b style="color:#B91C1C">รายการที่ผิดปกติ:</b>
        <ul style="margin:8px 0 0;padding-left:20px;color:#7F1D1D">${abnList}</ul></div>`:''}
      ${data.note?`<div style="background:#EFF6FF;border-left:4px solid #1B4F8A;border-radius:4px;padding:14px 16px">
        <b style="color:#1B4F8A">หมายเหตุ:</b>
        <div style="margin-top:6px;font-size:13px;color:#1E3A5F">${data.note}</div></div>`:''}
      <div style="margin-top:16px;background:#FFFBEB;border-radius:6px;padding:12px 16px;font-size:13px;color:#92400E">
        ⚡ กรุณาดำเนินการตรวจสอบและแก้ไขโดยด่วน</div>
    </div>
    <div style="background:#F9FAFB;padding:10px;text-align:center;font-size:11px;color:#9CA3AF;border-radius:0 0 10px 10px;border:1px solid #E5E7EB;border-top:none">
      แจ้งเตือนอัตโนมัติโดยระบบตรวจสอบเครื่องจักร</div>
  </div>`;

  try {
    MailApp.sendEmail({ to: toEmail, subject, htmlBody: html });
    Logger.log('Email sent to: ' + toEmail);
  } catch(e) { Logger.log('Email error: ' + e); }
}

function ok(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: msg||'' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function testRun() {
  const user = checkLogin('bfc43000', '1111111!');
  Logger.log(user ? 'Login OK: ' + JSON.stringify(user) : 'Login FAILED');
}
