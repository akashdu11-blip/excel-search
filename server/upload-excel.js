const path = require('path');
const fs = require('fs');
const express = require('express');
const cors = require('cors');
const multer = require('multer');

const {
  convertExcelsToProducts,
  readProductsJson,
  PRODUCTS_JSON_PATH
} = require('./convert-excel');

const app = express();
const PORT = process.env.PORT || 3000;

const UPLOAD_DIR = path.join(__dirname, '..', 'uploads');
const DATA_DIR = path.join(__dirname, '..', 'data');
const UPLOADED_FILENAMES_PATH = path.join(DATA_DIR, 'uploaded-filenames.json');

if (!fs.existsSync(UPLOAD_DIR)) {
  fs.mkdirSync(UPLOAD_DIR, { recursive: true });
}
if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}

function getUploadedFilenames() {
  if (!fs.existsSync(UPLOADED_FILENAMES_PATH)) return [];
  try {
    const raw = fs.readFileSync(UPLOADED_FILENAMES_PATH, 'utf8');
    const list = JSON.parse(raw || '[]');
    return Array.isArray(list) ? list : [];
  } catch {
    return [];
  }
}

function addUploadedFilenames(names) {
  const list = getUploadedFilenames();
  const set = new Set(list);
  names.forEach((n) => set.add(n));
  fs.writeFileSync(UPLOADED_FILENAMES_PATH, JSON.stringify([...set], null, 2), 'utf8');
}

const storage = multer.diskStorage({
  destination: function (_req, _file, cb) {
    cb(null, UPLOAD_DIR);
  },
  filename: function (_req, file, cb) {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1e9);
    cb(null, uniqueSuffix + '-' + file.originalname);
  }
});

const upload = multer({ storage });

app.use(cors());
app.use(express.json());

app.use(express.static(path.join(__dirname, '..', 'public')));

app.get('/health', (_req, res) => {
  res.json({ status: 'ok' });
});

app.post('/upload', upload.array('files', 20), (req, res) => {
  try {
    const files = req.files || [];
    if (!files.length) {
      return res.status(400).json({ error: 'No files uploaded' });
    }

    const uploadedNames = getUploadedFilenames();
    const incomingNames = files.map((f) => f.originalname);
    const duplicates = incomingNames.filter((name) => uploadedNames.includes(name));
    if (duplicates.length) {
      const message =
        duplicates.length === 1
          ? `File already uploaded: "${duplicates[0]}". Use a different file or rename it.`
          : `These files were already uploaded: ${duplicates.map((d) => `"${d}"`).join(', ')}. Use different files or rename them.`;
      return res.status(400).json({ error: message, duplicate_files: duplicates });
    }

    const filePaths = files.map((f) => f.path);
    const result = convertExcelsToProducts(filePaths);

    addUploadedFilenames(incomingNames);

    res.json({
      message: 'Files processed successfully',
      total_records: result.count,
      added: result.added,
      json_path: PRODUCTS_JSON_PATH
    });
  } catch (err) {
    console.error('Upload error:', err);
    res.status(500).json({ error: 'Failed to process files' });
  }
});

app.get('/products', (_req, res) => {
  const products = readProductsJson();
  res.json({
    total: products.length,
    products
  });
});

app.listen(PORT, () => {
  console.log(`Server listening on http://localhost:${PORT}`);
});

