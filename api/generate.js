const fs = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');

const MASTER_FILE = path.join(process.cwd(), 'SiamTin_Master.xlsx');
const MASTER_SHEET = 'ONE (2)';

/*
  Mapping จากข้อมูล Plan Load -> Cell ใน Excel Master
  ยึดตามไฟล์ SiamTin_Master.xlsx ที่ผู้ใช้ส่งมา
*/
const CELL_MAP = {
  'FA NO.': 'D6',
  'INVOICE NO.': 'F6',

  'FROM': 'D10',
  'TO': 'H10',
  'DRIVER NAME': 'D11',

  'CONT NO.': 'D12',
  'BOOKING NO': 'D13',
  'Liner': 'D14',
  'SEAL NO.': 'D15',

  'TRUCK NO.': 'J13',
  'TRAILER NO.': 'J14',
  'DRIVER PHONE': 'I15',

  'PICKUP CONTACT': 'C16',
  'PICKUP PLACE': 'C17',
  'PICKUP DATE': 'H17',
  'FACTORY DATE': 'K17',
  'FACTORY TIME': 'M17',

  'RETURN PLACE': 'C18',
  'Return date': 'H18',
  'RETURN TIME': 'M18',

  'APPROVER': 'K19',
  'APPROVE DATE': 'K20'
};

function escapeXml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function decodeXml(value) {
  return String(value ?? '')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function normalizeDate(value) {
  if (!value) return '';

  const text = String(value).trim();

  let m = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) {
    return [
      String(m[3]).padStart(2, '0'),
      String(m[2]).padStart(2, '0'),
      m[1]
    ].join('/');
  }

  m = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) {
    return [
      String(m[1]).padStart(2, '0'),
      String(m[2]).padStart(2, '0'),
      m[3]
    ].join('/');
  }

  return text;
}

function normalizeTime(value) {
  if (!value) return '';
  const text = String(value).trim();
  const m = text.match(/(\d{1,2}):(\d{2})/);
  if (!m) return text;

  return `${String(m[1]).padStart(2, '0')}:${m[2]}`;
}

function setCellInlineString(xml, cellRef, value) {
  const safe = escapeXml(value);
  const cellContent = `<is><t xml:space="preserve">${safe}</t></is>`;

  // Existing normal cell.
  const normal = new RegExp(
    `<c\\b([^>]*\\br="${cellRef}"[^>]*)>([\\s\\S]*?)<\\/c>`,
    'i'
  );

  if (normal.test(xml)) {
    return xml.replace(normal, (full, attrs) => {
      let cleaned = attrs
        .replace(/\s+t="[^"]*"/gi, '')
        .replace(/\s+cm="[^"]*"/gi, '');

      return `<c${cleaned} t="inlineStr">${cellContent}</c>`;
    });
  }

  // Existing self-closing cell.
  const selfClosing = new RegExp(
    `<c\\b([^>]*\\br="${cellRef}"[^>]*)\\/>`,
    'i'
  );

  if (selfClosing.test(xml)) {
    return xml.replace(selfClosing, (full, attrs) => {
      let cleaned = attrs
        .replace(/\s+t="[^"]*"/gi, '')
        .replace(/\s+cm="[^"]*"/gi, '');

      return `<c${cleaned} t="inlineStr">${cellContent}</c>`;
    });
  }

  // Cell does not exist: insert it inside the correct row.
  const rowNumber = cellRef.match(/\d+$/)?.[0];
  if (!rowNumber) {
    throw new Error(`Invalid cell reference: ${cellRef}`);
  }

  const rowRegex = new RegExp(
    `(<row\\b[^>]*\\br="${rowNumber}"[^>]*>)([\\s\\S]*?)(<\\/row>)`,
    'i'
  );

  if (!rowRegex.test(xml)) {
    throw new Error(`Row ${rowNumber} was not found in the Excel master.`);
  }

  return xml.replace(rowRegex, (full, open, inside, close) => {
    return `${open}${inside}<c r="${cellRef}" t="inlineStr">${cellContent}</c>${close}`;
  });
}

function getSheetXmlPath(zip, sheetName) {
  const workbookEntry = zip.getEntry('xl/workbook.xml');
  const relsEntry = zip.getEntry('xl/_rels/workbook.xml.rels');

  if (!workbookEntry || !relsEntry) {
    throw new Error('Excel workbook structure is incomplete.');
  }

  const workbookXml = workbookEntry.getData().toString('utf8');
  const relsXml = relsEntry.getData().toString('utf8');

  const escapedName = sheetName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

  const sheetMatch = workbookXml.match(
    new RegExp(
      `<sheet\\b[^>]*name="${escapedName}"[^>]*r:id="([^"]+)"[^>]*/?>`,
      'i'
    )
  );

  if (!sheetMatch) {
    throw new Error(`ไม่พบชีต "${sheetName}" ใน Excel Master`);
  }

  const relId = sheetMatch[1];

  const relMatch = relsXml.match(
    new RegExp(
      `<Relationship\\b[^>]*Id="${relId}"[^>]*Target="([^"]+)"[^>]*/?>`,
      'i'
    )
  );

  if (!relMatch) {
    throw new Error(`ไม่พบ worksheet relationship ของ "${sheetName}"`);
  }

  let target = decodeXml(relMatch[1]).replace(/^\/+/, '');

  if (!target.startsWith('xl/')) {
    target = 'xl/' + target.replace(/^\.\//, '');
  }

  return target;
}

function cleanFilePart(value) {
  return String(value || '')
    .trim()
    .replace(/[\\/:*?"<>|]+/g, '-')
    .replace(/\s+/g, '_')
    .slice(0, 60);
}

function applyTransportCompany(sheetXml, data) {
  const type = String(data['TRUCK TYPE'] || '').trim().toUpperCase();
  const company = data['Truck Com.'] || '';
  const phone = data['DRIVER PHONE'] || '';

  const trailer =
    type.includes('TRAILER') ||
    type.includes('หัวลาก');

  if (trailer) {
    sheetXml = setCellInlineString(sheetXml, 'D8', '');
    sheetXml = setCellInlineString(sheetXml, 'H8', '');
    sheetXml = setCellInlineString(sheetXml, 'D9', company);
    sheetXml = setCellInlineString(sheetXml, 'H9', phone);
  } else {
    sheetXml = setCellInlineString(sheetXml, 'D8', company);
    sheetXml = setCellInlineString(sheetXml, 'H8', phone);
    sheetXml = setCellInlineString(sheetXml, 'D9', '');
    sheetXml = setCellInlineString(sheetXml, 'H9', '');
  }

  return sheetXml;
}

function applyOptionalMarks(sheetXml, data) {
  /*
    Excel Master จริงเก็บตำแหน่งกล่องเดิมไว้
    ตัว API ใส่ "X" เฉพาะเมื่อหน้าเว็บส่ง markBoxes=true เท่านั้น
    ค่าเริ่มต้น = false เพื่อไม่ให้เกิด X เกิน/ลอย
  */
  if (data.markBoxes !== true) {
    return sheetXml;
  }

  const imex = String(data['IMPORT/EXPORT'] || '').toUpperCase();
  const product = String(data['PRODUCT TYPE'] || '').toUpperCase();
  const generator = String(data['GENERATOR SET'] || '').toUpperCase();
  const truckType = String(data['TRUCK TYPE'] || '').toUpperCase();

  sheetXml = setCellInlineString(sheetXml, 'H5', imex === 'IMPORT' ? 'X' : '');
  sheetXml = setCellInlineString(sheetXml, 'K5', imex === 'EXPORT' ? 'X' : '');

  sheetXml = setCellInlineString(sheetXml, 'H6', product === 'CANNING' ? 'X' : '');
  sheetXml = setCellInlineString(sheetXml, 'K6', product === 'FROZEN' ? 'X' : '');

  sheetXml = setCellInlineString(
    sheetXml,
    'B8',
    (truckType.includes('TRUCK') || truckType.includes('บรรทุก')) ? 'X' : ''
  );

  sheetXml = setCellInlineString(
    sheetXml,
    'B9',
    (truckType.includes('TRAILER') || truckType.includes('หัวลาก')) ? 'X' : ''
  );

  sheetXml = setCellInlineString(sheetXml, 'J12', generator === 'YES' ? 'X' : '');
  sheetXml = setCellInlineString(sheetXml, 'L12', generator === 'NO' ? 'X' : '');

  return sheetXml;
}

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(204).end();
  }

  if (req.method === 'GET') {
    return res.status(200).json({
      ok: true,
      service: 'Siam Tin Excel Export',
      masterSheet: MASTER_SHEET
    });
  }

  if (req.method !== 'POST') {
    return res.status(405).json({
      ok: false,
      error: 'Method not allowed'
    });
  }

  try {
    if (!fs.existsSync(MASTER_FILE)) {
      throw new Error('ไม่พบ SiamTin_Master.xlsx ในโปรเจกต์ Vercel');
    }

    const data =
      typeof req.body === 'string'
        ? JSON.parse(req.body || '{}')
        : (req.body || {});

    const masterBuffer = fs.readFileSync(MASTER_FILE);
    const zip = new AdmZip(masterBuffer);

    const sheetPath = getSheetXmlPath(zip, MASTER_SHEET);
    const sheetEntry = zip.getEntry(sheetPath);

    if (!sheetEntry) {
      throw new Error(`ไม่พบ ${sheetPath} ใน Excel Master`);
    }

    let sheetXml = sheetEntry.getData().toString('utf8');

    // Main fields.
    for (const [field, cell] of Object.entries(CELL_MAP)) {
      let value = data[field] ?? '';

      if ([
        'PICKUP DATE',
        'FACTORY DATE',
        'Return date',
        'APPROVE DATE'
      ].includes(field)) {
        value = normalizeDate(value);
      }

      if ([
        'FACTORY TIME',
        'RETURN TIME'
      ].includes(field)) {
        value = normalizeTime(value);
      }

      sheetXml = setCellInlineString(
        sheetXml,
        cell,
        value
      );
    }

    // Company / phone row depends on truck type.
    sheetXml = applyTransportCompany(
      sheetXml,
      data
    );

    // Leave checkboxes blank by default.
    sheetXml = applyOptionalMarks(
      sheetXml,
      data
    );

    zip.updateFile(
      sheetPath,
      Buffer.from(sheetXml, 'utf8')
    );

    const output = zip.toBuffer();

    const fileNameParts = [
      cleanFilePart(data['FA NO.'] || 'SiamTin'),
      cleanFilePart(data['BOOKING NO'] || '')
    ].filter(Boolean);

    const fileName =
      fileNameParts.join('_') +
      '.xlsx';

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );

    res.setHeader(
      'Content-Disposition',
      `attachment; filename="${fileName}"`
    );

    return res.status(200).send(output);

  } catch (error) {
    console.error(error);

    return res.status(500).json({
      ok: false,
      error: error.message || String(error)
    });
  }
};
