// ════════════════════════════════════════════════════════════
//  Code.gs — Google Apps Script Backend
//  รับข้อมูลจาก inspect.html โดยตรง (ไม่ผ่าน Worker)
// ════════════════════════════════════════════════════════════

const DEFAULT_EMAIL = 'your-manager@company.com';

// ── รับข้อมูลจาก GET (ข้อมูลหลัก) ─────────────────────────
function doGet(e) {
  try {
    const action = e.parameter.action;

    // รับผลการตรวจสอบ
    if (action === 'submit') {
      const data = e.parameter;
      saveToSheet(data, []);
      if (data.status === 'ABNORMAL') {
        sendAbnormalEmail(data, [], data.notifyEmail || DEFAULT_EMAIL);
      }
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ดึงประวัติ
    if (action === 'history') {
      const rows = getHistory(e.parameter.machineId || '', parseInt(e.parameter.months) || 3);
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', rows }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', message: 'Machine Inspection API Ready' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    Logger.log('doGet error: ' + err);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── รับรูปภาพจาก POST (no-cors) ────────────────────────────
function doPost(e) {
  try {
    const raw  = e.postData ? e.postData.contents : '{}';
    const data = JSON.parse(raw);

    if (data.action === 'photos' && data.photos && data.photos.length > 0) {
      const photoUrls = uploadPhotos(data.photos, data.machineId || 'photo', data.timestamp || '');

      // หาแถวล่าสุดของเครื่องนี้แล้วใส่ link รูป
      const sheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      const lastRow = sheet.getLastRow();
      for (let r = lastRow; r >= 2; r--) {
        const rowId        = sheet.getRange(r, 2).getValue();
        const rowInspector = String(sheet.getRange(r, 9).getValue());
        if (rowId === data.machineId && rowInspector.includes(String(data.inspector || '').substring(0, 5))) {
          photoUrls.forEach((url, i) => {
            if (url && i < 3) sheet.getRange(r, 15 + i).setValue(url);
          });
          break;
        }
      }

      // ส่งรูปผ่าน Email เพิ่มเติม (ถ้ามี)
      if (photoUrls.length > 0) {
        Logger.log('Photos uploaded: ' + photoUrls.length);
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    Logger.log('doPost error: ' + err);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── บันทึกลง Sheet ──────────────────────────────────────────
function saveToSheet(data, photoUrls) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (sheet.getLastRow() === 0) {
    const headers = [
      'วันเวลา','Machine ID','ชื่อเครื่องจักร',
      'โรงงาน (FC)','หน่วยผลิต','ภูมิภาค','ME ผู้ดูแล','ตำแหน่ง',
      'ผู้ตรวจสอบ','ผลการตรวจ',
      'รายการปกติ','รายการผิดปกติ','รายการที่ผิดปกติ','หมายเหตุ',
      'รูปที่ 1','รูปที่ 2','รูปที่ 3'
    ];
    sheet.appendRow(headers);
    sheet.getRange(1,1,1,headers.length)
      .setBackground('#1B4F8A').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    data.timestamp    || new Date().toLocaleString('th-TH'),
    data.machineId    || '',
    data.machineName  || '',
    data.fc           || '',
    data.unit         || '',
    data.region       || '',
    data.me           || '',
    data.location     || '',
    data.inspector    || '',
    data.status       || '',
    data.normalCount  || 0,
    data.abnormalCount || 0,
    data.abnormalItems || '',
    data.note         || '',
    photoUrls[0]      || '',
    photoUrls[1]      || '',
    photoUrls[2]      || '',
  ]);

  const lastRow = sheet.getLastRow();
  const rowRange = sheet.getRange(lastRow, 1, 1, 17);
  if (data.status === 'ABNORMAL') {
    rowRange.setBackground('#FEF2F2');
    sheet.getRange(lastRow, 10).setFontColor('#B91C1C').setFontWeight('bold');
  } else if (data.status === 'NORMAL') {
    rowRange.setBackground('#F0FDF4');
  }
}

// ── อัปโหลดรูปขึ้น Drive ────────────────────────────────────
function uploadPhotos(photos, machineId, timestamp) {
  const urls = [];
  const folder = getOrCreateFolder('Machine Inspection/' + machineId);
  photos.forEach((b64, i) => {
    if (!b64) return;
    try {
      const clean = b64.replace(/^data:image\/\w+;base64,/, '');
      const ts    = timestamp || Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmmss');
      const blob  = Utilities.newBlob(
        Utilities.base64Decode(clean), 'image/jpeg',
        `${machineId}_${ts}_${i+1}.jpg`
      );
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      urls.push('https://drive.google.com/uc?export=view&id=' + file.getId());
    } catch(imgErr) {
      Logger.log('Photo error: ' + imgErr);
    }
  });
  return urls;
}

// ── ดึงประวัติ ───────────────────────────────────────────────
function getHistory(machineId, months) {
  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - months);

  return sheet.getRange(2, 1, lastRow - 1, 17).getValues()
    .filter(r => {
      if (machineId && r[1] !== machineId) return false;
      const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
      return !isNaN(d) && d >= cutoff;
    })
    .map(r => ({
      time:          r[0] instanceof Date ? r[0].toLocaleString('th-TH') : String(r[0]),
      machineId:     r[1], machineName: r[2],
      inspector:     r[8], status:      r[9],
      normalCount:   r[10], abnormalCount: r[11],
      abnormalItems: r[12], note:        r[13],
      photo1: r[14], photo2: r[15], photo3: r[16],
    }))
    .reverse();
}

// ── ส่ง Email ────────────────────────────────────────────────
function sendAbnormalEmail(data, photoUrls, toEmail) {
  const subject = `⚠️ พบความผิดปกติ: ${data.machineName || data.machineId}`;
  const abnList = (data.abnormalItems || '').split(',')
    .filter(s => s.trim())
    .map(s => `<li style="margin:4px 0">${s.trim()}</li>`).join('');

  const html = `
<div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;max-width:580px;margin:0 auto">
  <div style="background:linear-gradient(135deg,#0F3260,#1B4F8A);color:#fff;padding:22px 24px;border-radius:10px 10px 0 0">
    <div style="font-size:12px;opacity:.75;margin-bottom:6px">ระบบตรวจสอบเครื่องจักร — แจ้งเตือนอัตโนมัติ</div>
    <div style="font-size:22px;font-weight:700">⚠️ พบความผิดปกติ</div>
    <div style="font-size:15px;margin-top:4px;opacity:.9">${data.machineName || data.machineId}</div>
  </div>
  <div style="background:#fff;padding:22px 24px;border:1px solid #E5E7EB;border-top:none">
    <table style="width:100%;border-collapse:collapse;margin-bottom:16px">
      ${[
        ['Machine ID', data.machineId],
        ['โรงงาน / หน่วยผลิต', `${data.fc||'—'} / ${data.unit||'—'}`],
        ['ME ผู้ดูแล', data.me || '—'],
        ['ผู้ตรวจสอบ', data.inspector],
        ['วันเวลา', data.timestamp],
      ].map(([label, val]) => `
        <tr style="border-bottom:1px solid #F3F4F6">
          <td style="padding:9px 12px;background:#F9FAFB;font-weight:600;font-size:13px;color:#4B5563;width:40%">${label}</td>
          <td style="padding:9px 12px;font-size:13px">${val}</td>
        </tr>`).join('')}
      <tr>
        <td style="padding:9px 12px;background:#FEF2F2;font-weight:600;font-size:13px;color:#B91C1C">ผลการตรวจ</td>
        <td style="padding:9px 12px;background:#FEF2F2;font-weight:700;color:#B91C1C">❌ ABNORMAL</td>
      </tr>
    </table>

    ${abnList ? `<div style="background:#FEF2F2;border-left:4px solid #B91C1C;border-radius:4px;padding:14px 16px;margin-bottom:16px">
      <div style="font-weight:700;color:#B91C1C;margin-bottom:8px">รายการที่ผิดปกติ:</div>
      <ul style="margin:0;padding-left:20px;color:#7F1D1D">${abnList}</ul>
    </div>` : ''}

    ${data.note ? `<div style="background:#EFF6FF;border-left:4px solid #1B4F8A;border-radius:4px;padding:14px 16px;margin-bottom:16px">
      <div style="font-weight:700;color:#1B4F8A;margin-bottom:6px">หมายเหตุ:</div>
      <div style="color:#1E3A5F;font-size:13px">${data.note}</div>
    </div>` : ''}

    <div style="background:#FFFBEB;border-radius:6px;padding:12px 16px;font-size:13px;color:#92400E">
      ⚡ กรุณาดำเนินการตรวจสอบและแก้ไขโดยด่วน
    </div>
  </div>
  <div style="background:#F9FAFB;padding:10px;text-align:center;font-size:11px;color:#9CA3AF;border-radius:0 0 10px 10px;border:1px solid #E5E7EB;border-top:none">
    แจ้งเตือนอัตโนมัติโดยระบบตรวจสอบเครื่องจักร
  </div>
</div>`;

  try {
    MailApp.sendEmail({ to: toEmail, subject, htmlBody: html });
    Logger.log('Email sent to: ' + toEmail);
  } catch(e) { Logger.log('Email error: ' + e); }
}

// ── Helper ───────────────────────────────────────────────────
function getOrCreateFolder(path) {
  let folder = DriveApp.getRootFolder();
  path.split('/').forEach(name => {
    const f = folder.getFoldersByName(name);
    folder = f.hasNext() ? f.next() : folder.createFolder(name);
  });
  return folder;
}

function testRun() { Logger.log('✅ Apps Script พร้อมใช้งาน'); }
