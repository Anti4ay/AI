const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
        Table, TableRow, TableCell, WidthType, BorderStyle, PageBreak } = require('docx');
const items = require('./content.js');

const CONTENT_W = 9026, COL = CONTENT_W / 2;
const nb = (s) => s
  .replace(/(\d)\s(?=\d{3}(\D|$))/g, '$1 ')
  .replace(/(\d)\s(?=\()/g, '$1 ')
  .replace(/№\s/g, '№ ')
  .replace(/\bд\.\s(?=\d)/g, 'д. ')
  .replace(/\bкв\.\s(?=\d)/g, 'кв. ');

const COMPANY = [
  ['КОМПАНИЯ', true], ['ООО «СтроимВместе»', true],
  ['ИНН 3702124940, КПП 370201001'], ['ОГРН 1153702027038'],
  ['Юридический адрес: 153021, г. Иваново,'], ['Проезд Красных Зорь, д. 2'],
  ['Банк: Тульское отделение № 8604 ПАО Сбербанк'], ['Р/с 40702810117000013610'],
  ['К/с 30101810300000000608'], ['БИК 047003608'], ['E-mail: Nvk0906@mail.ru'],
  [''], ['Директор', true], [''], ['_______________ / М.Б. Кузьмин'], ['М.П.'],
];
const CONTRACTOR = [
  ['ИСПОЛНИТЕЛЬ', true], ['ИП Орешкин Антон Вадимович', true],
  ['ИНН 100123323420'], ['ОГРНИП 324100000000411'],
  ['Адрес: 185014, Республика Карелия,'], ['г. Петрозаводск, пер. Попова, д. 6, кв. 42'],
  ['Банк: АО «ТБанк»'], ['Р/с 40802810400005859561'],
  ['К/с 30101810145250000974'], ['БИК 044525974'], ['E-mail: Antonoreshkin.Work@yandex.ru'],
  [''], ['Индивидуальный предприниматель', true], [''], ['_______________ / А.В. Орешкин'], ['М.П. (при наличии)'],
];
const SIG2 = [['Директор ООО «СтроимВместе»', true], [''], ['_______________ / М.Б. Кузьмин']];
const SIG2R = [['Индивидуальный предприниматель', true], [''], ['_______________ / А.В. Орешкин']];

/* ---------- markdown ---------- */
const md = [];
for (const [type, text] of items) {
  switch (type) {
    case 'title': md.push(`# ${text}`, ''); break;
    case 'center': md.push(`*${text}*`, ''); break;
    case 'plain': md.push(text, ''); break;
    case 'h': md.push(`## ${text}`, ''); break;
    case 'h2': md.push(`### ${text}`, ''); break;
    case 'p': md.push(text, ''); break;
    case 'sub': md.push(text, ''); break;
    case 'hr': md.push('---', ''); break;
    case 'pagebreak': md.push('---', ''); break;
    case 'requisites':
      md.push(...COMPANY.map(([t]) => t).filter(Boolean), '');
      md.push(...CONTRACTOR.map(([t]) => t).filter(Boolean), '');
      break;
    case 'sigline2':
      md.push(...SIG2.map(([t]) => t).filter(Boolean), '');
      md.push(...SIG2R.map(([t]) => t).filter(Boolean), '');
      break;
  }
}
fs.writeFileSync(process.argv[3], md.join('\n').replace(/\n{3,}/g, '\n\n') + '\n');

/* ---------- docx ---------- */
function runs(text, opts = {}) {
  return nb(text).split(/(\*\*[^*]+\*\*)/).filter(Boolean).map((chunk) => {
    const bold = chunk.startsWith('**') && chunk.endsWith('**');
    return new TextRun({ text: bold ? chunk.slice(2, -2) : chunk, bold: bold || opts.bold, size: opts.size || 22 });
  });
}
const para = (text, opts = {}) => new Paragraph({
  children: runs(text, opts),
  alignment: opts.alignment || AlignmentType.JUSTIFIED,
  spacing: { after: 120, line: 276 },
  indent: opts.indent,
});
const cellP = (text, bold, size) => new Paragraph({
  children: [new TextRun({ text: nb(text), bold: !!bold, size: size || 20 })],
  spacing: { after: 60 },
});
const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: none, bottom: none, left: none, right: none };
const twoCol = (l, r) => new Table({
  columnWidths: [COL, COL],
  width: { size: CONTENT_W, type: WidthType.DXA },
  borders: noBorders,
  rows: [new TableRow({ children: [l, r].map((col) => new TableCell({
    width: { size: COL, type: WidthType.DXA }, borders: noBorders,
    margins: { top: 60, bottom: 60, left: 0, right: 200 },
    children: col.map(([t, b]) => cellP(t, b)),
  })) })],
});

const body = [];
for (const [type, text] of items) {
  switch (type) {
    case 'title':
      body.push(new Paragraph({ children: [new TextRun({ text, bold: true, size: 30 })],
        alignment: AlignmentType.CENTER, spacing: { after: 120 } })); break;
    case 'center':
      body.push(new Paragraph({ children: [new TextRun({ text: nb(text), size: 22 })],
        alignment: AlignmentType.CENTER, spacing: { after: 240 } })); break;
    case 'plain':
      body.push(new Paragraph({ children: [new TextRun({ text, size: 22 })], spacing: { after: 300 } })); break;
    case 'h':
      body.push(new Paragraph({ children: [new TextRun({ text, bold: true, size: 24 })],
        heading: HeadingLevel.HEADING_1, spacing: { before: 300, after: 140 } })); break;
    case 'h2':
      body.push(new Paragraph({ children: [new TextRun({ text, bold: true, size: 22 })],
        heading: HeadingLevel.HEADING_2, spacing: { before: 220, after: 120 } })); break;
    case 'p': body.push(para(text)); break;
    case 'sub': body.push(para(text, { indent: { left: 340 } })); break;
    case 'hr':
      body.push(new Paragraph({ children: [],
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '999999', space: 6 } },
        spacing: { before: 120, after: 200 } })); break;
    case 'pagebreak':
      body.push(new Paragraph({ children: [new PageBreak()] })); break;
    case 'requisites': body.push(twoCol(COMPANY, CONTRACTOR)); break;
    case 'sigline2': body.push(twoCol(SIG2, SIG2R)); break;
  }
}

const doc = new Document({
  styles: { default: {
    document: { run: { font: 'Times New Roman', size: 22 } },
    heading1: { run: { font: 'Times New Roman', size: 24, bold: true, color: '000000' } },
    heading2: { run: { font: 'Times New Roman', size: 22, bold: true, color: '000000' } },
  } },
  sections: [{ properties: { page: { margin: { top: 1134, bottom: 1134, left: 1440, right: 1134 } } }, children: body }],
});
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(process.argv[2], b); console.log('docx + md собраны'); });
