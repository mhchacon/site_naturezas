const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx-js-style');
const fs = require('fs');
const path = require('path');

const app = express();
const port = process.env.PORT || 10000;
const natureDictionary = loadNatureDictionary();
const natureColumnName = 'COD_NATUREZA_TITULO_PAGAR';
const costCenterColumnName = 'COD_CENTRO_CUSTO_TITULO_PAGAR';
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (_request, file, callback) => {
    const isSpreadsheet = /\.xlsx?$/i.test(file.originalname);
    callback(isSpreadsheet ? null : new Error('Envie um arquivo .xlsx ou .xls.'), isSpreadsheet);
  }
});

app.use(express.static(path.join(__dirname, 'public')));

app.get('/health', (_request, response) => {
  response.json({ status: 'ok' });
});

app.post('/api/validate', upload.single('spreadsheet'), (request, response) => {
  try {
    if (!request.file) {
      return response.status(400).json({ error: 'Selecione uma planilha para continuar.' });
    }

    const workbook = XLSX.read(request.file.buffer, { type: 'buffer', cellStyles: true });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    const headers = rows[0] || [];
    const firstColumn = findColumnIndex(headers, natureColumnName);
    const secondColumn = findColumnIndex(headers, costCenterColumnName);

    if (firstColumn < 0 || secondColumn < 0) {
      return response.status(422).json({
        error: `A planilha precisa conter as colunas ${natureColumnName} e ${costCenterColumnName} na primeira linha.`,
        availableColumns: headers.filter(Boolean)
      });
    }

    const yellowFill = { patternType: 'solid', fgColor: { rgb: 'FFF1B8' } };
    let checkedRows = 0;
    let invalidRows = 0;

    for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
      const row = rows[rowIndex];
      if (row.every((value) => String(value).trim() === '')) continue;
      checkedRows += 1;

      const natureCode = normalizeCode(row[firstColumn]);
      const costCenterCode = normalizeCode(row[secondColumn]);
      const allowedCostCenters = natureDictionary.get(natureCode);
      const matches = allowedCostCenters?.has(costCenterCode) || false;
      if (!matches) {
        invalidRows += 1;
        highlightRow(sheet, rowIndex + 1, headers.length, yellowFill);
      }
    }

    const output = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx', cellStyles: true });
    const baseName = path.basename(request.file.originalname, path.extname(request.file.originalname));
    response.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="${baseName}-validada.xlsx"`,
      'X-Checked-Rows': String(checkedRows),
      'X-Invalid-Rows': String(invalidRows),
      'Access-Control-Expose-Headers': 'X-Checked-Rows, X-Invalid-Rows'
    });
    return response.send(output);
  } catch (error) {
    console.error(error);
    return response.status(500).json({ error: 'Não foi possível processar a planilha.' });
  }
});

app.use((error, _request, response, _next) => {
  response.status(400).json({ error: error.message || 'Arquivo inválido.' });
});

function findColumnIndex(headers, requestedName) {
  return headers.findIndex((header) => normalize(header) === normalize(requestedName));
}

function normalize(value) {
  return String(value ?? '').trim();
}

function normalizeCode(value) {
  return normalize(value).replace(/^0+(?=\d)/, '');
}

function loadNatureDictionary() {
  const dictionaryPath = path.join(__dirname, 'dicionario.csv');
  const lines = fs.readFileSync(dictionaryPath, 'utf8').replace(/^\uFEFF/, '').split(/\r?\n/);
  const dictionary = new Map();

  for (const line of lines.slice(1)) {
    if (!line.trim()) continue;
    const columns = line.split(';');
    const natureCode = normalizeCode(columns[0]);
    const costCenters = columns.slice(1).map(normalizeCode).filter(Boolean);
    if (natureCode) dictionary.set(natureCode, new Set(costCenters));
  }

  return dictionary;
}

function highlightRow(sheet, rowNumber, columnCount, fill) {
  for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
    const address = XLSX.utils.encode_cell({ r: rowNumber - 1, c: columnIndex });
    if (!sheet[address]) sheet[address] = { t: 's', v: '' };
    sheet[address].s = { ...(sheet[address].s || {}), fill };
  }
}

app.listen(port, () => {
  console.log(`Naturezas Planilha running on port ${port}`);
});
