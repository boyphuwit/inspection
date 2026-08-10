// ════════════════════════════════════════════════════════════
//  Code.gs — Google Apps Script Backend
//  Database Sheet: "Data base"
//  Columns: A=User, B=Password, C=Plant Code, D=Plant Name,
//           E=Area, F=EN-ME, G=Plant E-mail
// ════════════════════════════════════════════════════════════

const SHEET_DB   = 'Data base';  // Sheet ข้อมูล User
const SHEET_DATA = 'Data';       // Sheet บันทึกผลตรวจ

function doGet(e) {
  try {
    const action = e.parameter.action;

    // ── Login ──────────────────────────────────────────────
    if (action === 'login') {
      const username = (e.parameter.username || '').trim();
      const password = (e.parameter.password || '').trim();
      const user     = checkLogin(username, password);

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

// ── Login จาก Sheet "Data base" ──────────────────────────────
// Columns: A=User, B=Password, C=Plant Code, D=Plant Name,
//          E=Area, F=EN-ME, G=Plant E-mail
function checkLogin(username, password) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
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
        username:  u,
        password:  p,
        plantCode: String(row[2]).trim(), // C = Plant Code
        plantName: String(row[3]).trim(), // D = Plant Name
        area:      String(row[4]).trim(), // E = Area
        enMe:      String(row[5]).trim(), // F = EN-ME
        email:     String(row[6]).trim(), // G = Plant E-mail
        // สำหรับแสดงผลในหน้าตรวจ
        name:      u,
        role:      'technician'
      };
    }
  }
  return null;
}

// ── บันทึกลง Sheet "Data" ────────────────────────────────────
function saveToSheet(data) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_DATA);
  if (!sheet) sheet = ss.insertSheet(SHEET_DATA);

  const checks = JSON.parse(data.checksJson || '[]');

  if (sheet.getLastRow() === 0) {
    const base = [
      'วันเวลา','Plant Code','Plant Name','Area','EN-ME',
      'ผู้ตรวจสอบ','ผลการตรวจ'
    ];
    const chk  = checks.map(c =>
      c.label.replace(/^\d+[a-z]*\.\s*\[.*?\]\s*/, '').split('—')[0].trim()
    );
    const tail = ['รายการผิดปกติ','หมายเหตุ','รูปที่ 1','รูปที่ 2','รูปที่ 3'];
    const all  = [...base, ...chk, ...tail];
    sheet.appendRow(all);
    sheet.getRange(1,1,1,all.length)
      .setBackground('#1B4F8A').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  const base = [
    data.timestamp   || new Date().toLocaleString('th-TH'),
    data.plantCode   || data.fc || '',
    data.plantName   || data.machineName || '',
    data.area        || '',
    data.enMe        || data.me || '',
    data.inspector   || '',
    data.status      || '',
  ];
  const chkVals = checks.map(c =>
    c.state === 'normal' ? 'ปกติ' : c.state === 'abnormal' ? 'ผิดปกติ' : '—'
  );
  const tail = [data.abnormalItems||'', data.note||'', '','',''];

  sheet.appendRow([...base, ...chkVals, ...tail]);

  // ระบายสีแถว
  const lastRow  = sheet.getLastRow();
  const totalCol = base.length + chkVals.length + tail.length;
  const rowRange = sheet.getRange(lastRow, 1, 1, totalCol);

  if (data.status === 'ABNORMAL') {
    rowRange.setBackground('#FEF2F2');
    checks.forEach((c, i) => {
      if (c.state === 'abnormal') {
        sheet.getRange(lastRow, base.length + 1 + i)
          .setBackground('#FCA5A5').setFontColor('#B91C1C').setFontWeight('bold');
      }
    });
  } else if (data.status === 'NORMAL') {
    rowRange.setBackground('#F0FDF4');
  }
}

// ── History ──────────────────────────────────────────────────
function getHistory(machineId, months) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DATA);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const cutoff  = new Date();
  cutoff.setMonth(cutoff.getMonth() - months);
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1,1,1,lastCol).getValues()[0];
  const statusIdx = headers.indexOf('ผลการตรวจ');
  const abnIdx    = headers.indexOf('รายการผิดปกติ');
  const noteIdx   = headers.indexOf('หมายเหตุ');

  return sheet.getRange(2,1,sheet.getLastRow()-1,lastCol).getValues()
    .filter(r => {
      const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
      return !isNaN(d) && d >= cutoff;
    })
    .map(r => ({
      time:          r[0] instanceof Date ? r[0].toLocaleString('th-TH') : String(r[0]),
      plantCode:     r[1], plantName: r[2], inspector: r[5],
      status:        statusIdx >= 0 ? r[statusIdx] : '',
      abnormalItems: abnIdx    >= 0 ? r[abnIdx]    : '',
      note:          noteIdx   >= 0 ? r[noteIdx]   : '',
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
        ['Area',       data.area||'—'],
        ['EN-ME',      data.enMe||'—'],
        ['ผู้ตรวจสอบ', data.inspector||'—'],
        ['วันเวลา',    data.timestamp||'—'],
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
