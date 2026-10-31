const express = require('express');
const multer = require('multer');
const cors = require('cors');
const { execFile } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');

const app = express();
app.use(cors({ origin: process.env.ALLOWED_ORIGIN || true }));
const upload = multer({ dest: os.tmpdir(), limits: { fileSize: 100 * 1024 * 1024 } });

app.get('/', (req, res) => res.send('Converter is running.'));

// LibreOffice handles one job at a time, so jobs wait in line
let queue = Promise.resolve();
const convert = (input, outDir, profile) => new Promise((resolve, reject) => {
  execFile('soffice', ['--headless', '--norestore', '-env:UserInstallation=file://' + profile,
    '--convert-to', 'pdf', '--outdir', outDir, input], { timeout: 5 * 60 * 1000 }, err => err ? reject(err) : resolve());
});

app.post('/convert', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file received.' });
  if (!/\.docx?$/i.test(req.file.originalname)) { fs.rm(req.file.path, { force: true }, () => {}); return res.status(400).json({ error: 'Only .docx files are supported.' }); }
  const id = crypto.randomBytes(6).toString('hex');
  const dir = path.join(os.tmpdir(), 'job-' + id);
  fs.mkdirSync(dir);
  const input = path.join(dir, 'input.docx');
  fs.renameSync(req.file.path, input);
  const cleanup = () => fs.rm(dir, { recursive: true, force: true }, () => {});

  queue = queue.then(() => convert(input, dir, path.join(dir, 'profile'))).then(() => {
    const pdf = path.join(dir, 'input.pdf');
    if (!fs.existsSync(pdf)) throw new Error('no output');
    res.setHeader('Content-Type', 'application/pdf');
    res.sendFile(pdf, () => cleanup());
  }).catch(() => {
    cleanup();
    if (!res.headersSent) res.status(500).json({ error: 'The server could not convert this file.' });
  });
});

app.listen(process.env.PORT || 10000, () => console.log('Converter ready'));
