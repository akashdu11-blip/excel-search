const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const DATA_DIR = path.join(__dirname, '..', 'data');
const PRODUCTS_JSON_PATH = path.join(DATA_DIR, 'products.json');

function normalizeHeader(name) {
  if (!name) return '';
  return String(name)
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function findRequiredColumns(headerRow) {
  const indices = {
    productNo: -1,
    description: -1,
    rate: -1
  };

  headerRow.forEach((h, idx) => {
    const norm = normalizeHeader(h);
    const productCandidates = [
      'productno',
      'productnumber',
      'product',
      'productcode',
      'productid',
      'itemcode',
      'code'
    ];
    const descriptionCandidates = [
      'description',
      'desc',
      'productdescription',
      'item',
      'itemname',
      'productname',
      'productdesc'
    ];
    const rateCandidates = [
      'rate',
      'price',
      'unitrate',
      'mrp',
      'sellingprice',
      'unitprice'
    ];

    if (indices.productNo === -1 && productCandidates.includes(norm)) {
      indices.productNo = idx;
    } else if (indices.description === -1 && descriptionCandidates.includes(norm)) {
      indices.description = idx;
    } else if (indices.rate === -1 && rateCandidates.includes(norm)) {
      indices.rate = idx;
    }
  });

  if (indices.productNo === -1 || indices.description === -1 || indices.rate === -1) {
    return null;
  }
  return indices;
}

function extractFromSheet(sheet) {
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
  let headerRowIndex = range.s.r;
  let colIdx = null;

  // Try to detect header row within the first 10 rows
  const maxHeaderRow = Math.min(range.s.r + 9, range.e.r);
  for (let R = range.s.r; R <= maxHeaderRow; R++) {
    const headerRow = [];
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = sheet[cellAddress];
      headerRow.push(cell ? cell.v : '');
    }
    colIdx = findRequiredColumns(headerRow);
    if (colIdx) {
      headerRowIndex = R;
      break;
    }
  }

  if (!colIdx) return [];

  const products = [];
  for (let R = headerRowIndex + 1; R <= range.e.r; R++) {
    const row = [];
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = sheet[cellAddress];
      row.push(cell ? cell.v : '');
    }

    const productNo = row[colIdx.productNo];
    const description = row[colIdx.description];
    let rate = row[colIdx.rate];

    if (productNo == null || productNo === '' || description == null || description === '') {
      continue;
    }

    if (typeof rate === 'string') {
      const parsed = parseFloat(rate.replace(/[^0-9.-]/g, ''));
      rate = isNaN(parsed) ? null : parsed;
    }

    products.push({
      product_no: String(productNo).trim(),
      description: String(description).trim(),
      rate: typeof rate === 'number' ? rate : null
    });
  }

  return products;
}

function convertExcelsToProducts(filePaths) {
  const existingProducts = readProductsJson();
  const newProducts = [];

  filePaths.forEach((filePath) => {
    const workbook = XLSX.readFile(filePath);
    workbook.SheetNames.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      const products = extractFromSheet(sheet);
      newProducts.push(...products);
    });
  });

  const allProducts = [...existingProducts, ...newProducts];

  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
  }

  fs.writeFileSync(PRODUCTS_JSON_PATH, JSON.stringify(allProducts, null, 2), 'utf8');

  return {
    count: allProducts.length,
    added: newProducts.length,
    path: PRODUCTS_JSON_PATH
  };
}

function readProductsJson() {
  if (!fs.existsSync(PRODUCTS_JSON_PATH)) {
    return [];
  }
  try {
    const raw = fs.readFileSync(PRODUCTS_JSON_PATH, 'utf8');
    if (!raw.trim()) return [];
    return JSON.parse(raw);
  } catch (err) {
    console.error('Error reading products.json:', err);
    return [];
  }
}

module.exports = {
  convertExcelsToProducts,
  readProductsJson,
  PRODUCTS_JSON_PATH
};

