const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const os = require('os');
// Use local copy inside backend folder for Render deploys
const { generatePDF } = require('./pdf-generator');

const app = express();
const PORT = process.env.PORT || 8080;

app.use(cors({ origin: '*'}));
app.use(express.json({ limit: '50mb' }));

function writeBase64Image(base64, prefix = 'img') {
  return new Promise((resolve, reject) => {
    try {
      const matches = base64.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
      const buffer = Buffer.from(matches ? matches[2] : base64, 'base64');
      const ext = matches ? (matches[1].split('/')[1] || 'png') : 'png';
      const filePath = path.join(os.tmpdir(), `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2)}.${ext}`);
      fs.writeFile(filePath, buffer, err => {
        if (err) return reject(err);
        resolve(filePath);
      });
    } catch (e) {
      reject(e);
    }
  });
}

app.get('/health', (_req, res) => res.json({ ok: true }));

app.post('/generate', async (req, res) => {
  try {
    const { data = {}, images = {}, logo_base64 } = req.body || {};

    // Prepare images as file paths for existing generator
    const sections = ['cover', 'property', 'floor_plans', 'directions', 'city'];
    const imagePaths = {};
    const tempFiles = [];
    for (const section of sections) {
      const arr = Array.isArray(images[section]) ? images[section] : [];
      imagePaths[section] = [];
      for (const b64 of arr) {
        if (typeof b64 === 'string' && b64.startsWith('http')) {
          // Remote URLs are not supported directly by PDFKit.addImage; skip for now
          continue;
        } else if (typeof b64 === 'string') {
          const fp = await writeBase64Image(b64, section);
          tempFiles.push(fp);
          imagePaths[section].push(fp);
        }
      }
    }

    // Logo
    let logoPath = null;
    if (logo_base64) {
      logoPath = await writeBase64Image(logo_base64, 'logo');
      tempFiles.push(logoPath);
      console.log('Logo received from frontend, saved to:', logoPath);
    } else {
      const defaultLogo = path.join(__dirname, 'logo.png');
      if (fs.existsSync(defaultLogo)) {
        logoPath = defaultLogo;
        console.log('Using default logo from:', defaultLogo);
      } else {
        console.warn('No logo_base64 provided and default logo not found at:', defaultLogo);
      }
    }

    // Output PDF path
    const outPath = path.join(os.tmpdir(), `report-${Date.now()}.pdf`);

    // Generate PDF using existing generator
    await generatePDF(data, imagePaths, outPath, logoPath);

    // Stream file and then cleanup
    res.setHeader('Content-Type', 'application/pdf');
    const safeName = (data.address || 'Property Report').replace(/[^a-z0-9 \-_]/gi, '');
    res.setHeader('Content-Disposition', `attachment; filename="${safeName}.pdf"`);
    const stream = fs.createReadStream(outPath);
    stream.pipe(res);
    stream.on('close', () => {
      try { fs.unlinkSync(outPath); } catch {}
      for (const f of tempFiles) { try { fs.unlinkSync(f); } catch {} }
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || 'Failed to generate PDF' });
  }
});

app.listen(PORT, () => {
  console.log(`PDF backend listening on port ${PORT}`);
});


