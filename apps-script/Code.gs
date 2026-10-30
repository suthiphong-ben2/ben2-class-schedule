/**
 * ระบบตารางเรียน/ตารางสอบ โรงเรียนเบญจมราชรังสฤษฎิ์ ๒ — ส่วนหลังบ้าน (Google Apps Script)
 *
 * ชีตที่ใช้ (สร้างให้อัตโนมัติเมื่อรัน setup):
 *   Timetable  ข้อมูลตารางสอนจากไฟล์ T_ONE เก็บเป็น JSON แบ่งเป็นท่อน (อย่าแก้มือ)
 *   Exams      ตารางสอบ  — แก้ในชีตตรงๆ ได้ (วันที่ใช้รูปแบบ 2027-03-01)
 *   Periods    เวลาคาบ 1–10
 *   Subjects   รหัสวิชา → ชื่อวิชา
 *
 * ติดตั้ง: 1) แก้ ADMIN_PASSWORD ด้านล่าง  2) รันฟังก์ชัน setup หนึ่งครั้ง
 *          3) Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone)
 */

const ADMIN_PASSWORD = 'เปลี่ยนรหัสผ่านตรงนี้';   // ใช้ตอนรัน setup เท่านั้น ระบบเก็บไว้ใน Script Properties
const CHUNK = 45000;                              // ช่องหนึ่งช่องใน Sheets เก็บได้ไม่เกิน 50,000 ตัวอักษร
const TZ = 'Asia/Bangkok';

const SHEETS = {
  Timetable: ['chunk'],
  Exams: ['วันที่ (ค.ศ. ปปปป-ดด-วว)', 'เวลา', 'ระดับชั้น (1–6)', 'รหัสวิชา', 'ชื่อวิชา', 'ห้องสอบ', 'หมายเหตุ'],
  Periods: ['คาบ', 'เริ่ม', 'จบ'],
  Subjects: ['รหัสวิชา', 'ชื่อวิชา'],
};
const EXAM_KEYS = ['date', 'time', 'level', 'code', 'name', 'rooms', 'note'];

/* ---------- ติดตั้ง ---------- */
function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(SHEETS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    sh.getRange('A:G').setNumberFormat('@'); // เก็บเป็นข้อความ กันวันที่/เวลาถูกแปลงเอง
  });
  const per = ss.getSheetByName('Periods');
  if (per.getLastRow() < 2) {
    const rows = [];
    for (let i = 1; i <= 10; i++) rows.push([String(i), '', '']);
    per.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  if (ADMIN_PASSWORD && ADMIN_PASSWORD !== 'เปลี่ยนรหัสผ่านตรงนี้') {
    PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD', ADMIN_PASSWORD);
  } else {
    throw new Error('กรุณาแก้ ADMIN_PASSWORD ด้านบนก่อนรัน setup');
  }
  bumpVersion_();
}

/* ---------- อ่านข้อมูล (หน้าดูตาราง) ---------- */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'data';
  if (action === 'version') return json_({ ok: true, version: getVersion_() });
  if (action === 'data') return json_(Object.assign({ ok: true }, readAll_()));
  return json_({ ok: false, error: 'unknown action' });
}

function readAll_() {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActive();
  let tt = { meta: { term: '', year: '', school: '' }, teachers: [] };
  const tsh = ss.getSheetByName('Timetable');
  if (tsh && tsh.getLastRow() > 1) {
    const txt = tsh.getRange(2, 1, tsh.getLastRow() - 1, 1).getValues().map(r => r[0]).join('');
    if (txt) tt = JSON.parse(txt);
  }
  return {
    version: getVersion_(),
    updatedAt: props.getProperty('UPDATED_AT') || '',
    meta: tt.meta, teachers: tt.teachers,
    exams: rows_('Exams').map(r => obj_(EXAM_KEYS, r)).filter(x => x.date || x.code),
    periods: rows_('Periods').map(r => ({ p: +r[0], start: str_(r[1]), end: str_(r[2]) })).filter(x => x.p >= 1 && x.p <= 10),
    subjects: Object.fromEntries(rows_('Subjects').filter(r => r[0]).map(r => [str_(r[0]).trim(), str_(r[1]).trim()])),
  };
}

/* ---------- เขียนข้อมูล (หน้าแอดมิน) ---------- */
function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'รูปแบบคำขอไม่ถูกต้อง' }); }
  const pw = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!pw || req.password !== pw) { Utilities.sleep(1500); return json_({ ok: false, error: 'รหัสผ่านไม่ถูกต้อง' }); }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return json_({ ok: false, error: 'มีคนกำลังบันทึกอยู่ ลองใหม่อีกครั้ง' });
  try {
    switch (req.action) {
      case 'login':
        return json_({ ok: true });
      case 'saveTimetable': {
        const d = req.data || {};
        if (!d.teachers || !d.teachers.length) return json_({ ok: false, error: 'ไม่มีข้อมูลครูในไฟล์' });
        const txt = JSON.stringify({ meta: d.meta, teachers: d.teachers });
        const parts = [];
        for (let i = 0; i < txt.length; i += CHUNK) parts.push([txt.slice(i, i + CHUNK)]);
        replaceRows_('Timetable', parts);
        break;
      }
      case 'saveExams':
        replaceRows_('Exams', (req.rows || []).map(x => EXAM_KEYS.map(k => str_(x[k]))));
        break;
      case 'savePeriods':
        replaceRows_('Periods', (req.rows || []).map(x => [String(x.p), str_(x.start), str_(x.end)]));
        break;
      case 'saveSubjects':
        replaceRows_('Subjects', Object.keys(req.subjects || {}).filter(k => req.subjects[k]).sort()
          .map(k => [k, str_(req.subjects[k])]));
        break;
      default:
        return json_({ ok: false, error: 'unknown action' });
    }
    return json_({ ok: true, version: bumpVersion_() });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- ตัวช่วย ---------- */
function rows_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
}
function replaceRows_(name, rows) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getMaxColumns()).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, rows[0].length).setNumberFormat('@').setValues(rows);
}
function obj_(keys, r) { const o = {}; keys.forEach((k, i) => o[k] = str_(r[i])); return o; }
function str_(v) {
  if (v instanceof Date) {
    // ช่องที่ Sheets แปลงเป็นวันที่/เวลาเอง: 1899-12-30 คือค่าเวลาล้วน
    return v.getFullYear() < 1900 ? Utilities.formatDate(v, TZ, 'HH:mm') : Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  }
  return v == null ? '' : String(v);
}
function getVersion_() { return Number(PropertiesService.getScriptProperties().getProperty('VERSION') || 0); }
function bumpVersion_() {
  const props = PropertiesService.getScriptProperties();
  const v = getVersion_() + 1;
  props.setProperties({ VERSION: String(v), UPDATED_AT: Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss") });
  return v;
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* แก้ข้อมูลในชีตด้วยมือ (เช่นตารางสอบ) แล้วหน้าเว็บจะโหลดข้อมูลใหม่เอง */
function onEdit(e) { try { bumpVersion_(); } catch (err) {} }
function refreshVersion() { bumpVersion_(); }
