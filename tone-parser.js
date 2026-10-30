/* อ่านไฟล์ตารางสอน T_ONE (.XLS แบบเก่าที่ส่งออกจากโปรแกรมจัดตารางสอน)
 * 1 แถว = ครู 1 คน, คอลัมน์ a101..a510 = วันจันทร์–ศุกร์ คาบ 1–10
 * ใช้คู่กับ SheetJS: rows = XLSX.utils.sheet_to_json(sheet, {header:1, defval:''})
 * คืนค่า {meta:{school,term,year}, teachers:[{id,n,g,t,x:[{d,p,s,c,r}|{d,p,a}]}], warnings:[]}
 */
(function (root) {
  const DAYCOL = [4, 15, 26, 37, 48]; // คอลัมน์คาบ 1 ของแต่ละวัน (ตามหัวคอลัมน์ a101, a201, ...)

  function parseCell(v, d, p) {
    const parts = v.split(/\s{2,}/);
    const e = { d, p };
    const flat = v.replace(/\s+/g, ' ');
    const m1 = /^(เสริม\S+?|CEFR)\s*(\d\/\d+)\s+(\S+)$/.exec(flat);
    if (/^[ก-ฮA-Za-z]\d{5}/.test(parts[0])) {
      e.s = parts[0];
      if (parts.length > 1) e.c = parts[1];
      if (parts.length > 2) e.r = parts[2];
    } else if (m1) {
      e.s = m1[1]; e.c = m1[2]; e.r = m1[3];
    } else if (/^[ก-ฮ]\s+เสรี/.test(v)) {
      e.s = parts[0] + ' (เสรี)'; e.c = 'เสรี';
      if (parts.length > 2) e.r = parts[2];
    } else {
      e.a = flat.replace('การ เงิน', 'การเงิน');
    }
    return e;
  }

  function parseTOne(rows) {
    const warnings = [];
    let meta = null;
    const teachers = [];
    const header = (rows[0] || []).map(String);
    if (header[4] && header[4] !== 'a101') warnings.push('หัวคอลัมน์ไม่ตรงรูปแบบ T_ONE (คาดว่าคอลัมน์ E คือ a101) ผลลัพธ์อาจคลาดเคลื่อน');
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r].map(x => String(x == null ? '' : x));
      const nm = /ชื่อ-สกุล\s+(\S+)\s+(.+?)\s+จำนวนคาบ\s*(\d+)/.exec(row[2] || '');
      if (!nm) { if (row.join('').trim()) warnings.push(`แถว ${r + 1}: อ่านชื่อครูไม่ได้ ข้ามแถวนี้`); continue; }
      const hdr = row[1] || '';
      const m = /ภาคเรียนที่\s*(\d)\s+ปีการศึกษา\s*(\d{4})/.exec(hdr);
      if (m && !meta) meta = { term: m[1], year: m[2], school: (row[0] || '').trim() };
      const dept = hdr.split('งาน').pop().trim();
      const slots = [];
      DAYCOL.forEach((c0, d) => {
        for (let p = 0; p < 10; p++) {
          const v = (row[c0 + p] || '').trim();
          if (v) slots.push(parseCell(v, d, p));
        }
      });
      teachers.push({ id: nm[1], n: nm[2].trim(), g: dept, t: +nm[3], x: slots });
    }
    if (!meta) warnings.push('ไม่พบข้อความ "ภาคเรียนที่ … ปีการศึกษา …" ในไฟล์');
    if (!teachers.length) warnings.push('ไม่พบข้อมูลครูในไฟล์ ตรวจสอบว่าเป็นไฟล์ T_ONE');
    return { meta: meta || { term: '', year: '', school: '' }, teachers, warnings };
  }

  root.parseTOne = parseTOne;
  if (typeof module !== 'undefined') module.exports = { parseTOne };
})(typeof window !== 'undefined' ? window : globalThis);
