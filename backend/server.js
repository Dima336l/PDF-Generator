const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const os = require('os');
const https = require('https');
const http = require('http');
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

// Search for city images using Pexels API
app.get('/search-city-images', async (req, res) => {
  try {
    const city = req.query.city || 'liverpool';
    const cityLower = city.toLowerCase();
    const searchQuery = `${cityLower} uk`;
    
    // Pexels API key - get a free one at https://www.pexels.com/api/
    // Using provided API key (for production, use environment variable instead)
    const PEXELS_API_KEY = process.env.PEXELS_API_KEY || '5azhaVVM7wQ71cypZeDeoFxGffvuwLH75qfzY44fYvACS2D0Oah5xPIR';
    
    if (PEXELS_API_KEY) {
      // Use Pexels API to get city-specific images
      const apiUrl = `https://api.pexels.com/v1/search?query=${encodeURIComponent(searchQuery)}&per_page=3&orientation=landscape`;
      
      https.get(apiUrl, {
        headers: {
          'Authorization': PEXELS_API_KEY  // Pexels API uses the key directly, not "Bearer {key}"
        }
      }, (apiRes) => {
        let data = '';
        apiRes.on('data', chunk => data += chunk);
        apiRes.on('end', () => {
          try {
            const jsonData = JSON.parse(data);
            if (jsonData.photos && jsonData.photos.length > 0) {
              const images = jsonData.photos.map(photo => {
                const baseUrl = photo.src.large;
                // Check if URL already has query parameters
                const separator = baseUrl.includes('?') ? '&' : '?';
                return `${baseUrl}${separator}auto=compress&cs=tinysrgb&w=800&h=600&fit=crop`;
              });
              console.log(`Found ${images.length} Liverpool-specific images from Pexels API`);
              return res.json({ images });
            }
          } catch (parseErr) {
            console.error('Error parsing Pexels response:', parseErr);
          }
          // Fallback if API fails
          returnFallback();
        });
      }).on('error', (err) => {
        console.error('Pexels API error:', err.message);
        returnFallback();
      });
    } else {
      // No API key - use fallback
      console.log('No Pexels API key found. Using fallback images.');
      console.log('To get Liverpool-specific images, get a free API key at https://www.pexels.com/api/');
      console.log('Then set it as: PEXELS_API_KEY=your_key_here');
      returnFallback();
    }
    
    function returnFallback() {
      // Fallback: These may not be Liverpool-specific
      // For best results, get a Pexels API key
      const fallbackImages = [
        'https://images.pexels.com/photos/417074/pexels-photo-417074.jpeg?auto=compress&cs=tinysrgb&w=800&h=600&fit=crop',
        'https://images.pexels.com/photos/417049/pexels-photo-417049.jpeg?auto=compress&cs=tinysrgb&w=800&h=600&fit=crop',
        'https://images.pexels.com/photos/417078/pexels-photo-417078.jpeg?auto=compress&cs=tinysrgb&w=800&h=600&fit=crop'
      ];
      res.json({ images: fallbackImages });
    }
    
  } catch (err) {
    console.error('Error searching city images:', err);
    res.status(500).json({ error: 'Failed to search images' });
  }
});

// Proxy endpoint to fetch images from external sources (bypasses CORS)
app.get('/proxy-image', async (req, res) => {
  try {
    const imageUrl = req.query.url;
    if (!imageUrl) {
      return res.status(400).json({ error: 'Missing url parameter' });
    }

    console.log('Proxy request for URL:', imageUrl);

    // Validate URL
    let urlObj;
    try {
      urlObj = new URL(imageUrl);
    } catch (e) {
      console.error('Invalid URL:', imageUrl, e);
      return res.status(400).json({ error: 'Invalid URL' });
    }

    // Only allow certain domains for security
    const allowedDomains = [
      'source.unsplash.com',
      'images.unsplash.com',
      'unsplash.com',
      'images.pexels.com',
      'pexels.com',
      'api.pexels.com'
    ];
    
    if (!allowedDomains.some(domain => urlObj.hostname.includes(domain))) {
      console.error('Domain not allowed:', urlObj.hostname);
      return res.status(403).json({ error: 'Domain not allowed' });
    }

    // Fetch the image with redirect handling
    const protocol = urlObj.protocol === 'https:' ? https : http;
    
    const fetchImage = (url, redirectCount = 0) => {
      if (redirectCount > 5) {
        return res.status(503).json({ error: 'Too many redirects' });
      }
      
      const options = {
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
      };
      
      protocol.get(url, options, (imageRes) => {
        // Handle redirects (301, 302, 307, 308)
        if (imageRes.statusCode >= 300 && imageRes.statusCode < 400 && imageRes.headers.location) {
          const redirectUrl = imageRes.headers.location;
          console.log(`Following redirect ${redirectCount + 1} to:`, redirectUrl);
          return fetchImage(redirectUrl, redirectCount + 1);
        }
        
        if (imageRes.statusCode !== 200) {
          console.error('Image fetch failed with status:', imageRes.statusCode, 'for URL:', url);
          // Don't send response if already sent
          if (!res.headersSent) {
            return res.status(503).json({ error: `Failed to fetch image: ${imageRes.statusCode}` });
          }
          return;
        }

        // Don't send response if already sent (from redirect)
        if (res.headersSent) {
          return;
        }

        // Set appropriate headers
        const contentType = imageRes.headers['content-type'] || 'image/jpeg';
        res.setHeader('Content-Type', contentType);
        res.setHeader('Cache-Control', 'public, max-age=86400'); // Cache for 1 day

        // Pipe the image data to the response
        imageRes.pipe(res);
      }).on('error', (err) => {
        console.error('Error fetching image:', err.message, 'for URL:', url);
        // Don't send response if already sent
        if (!res.headersSent) {
          res.status(503).json({ error: 'Failed to fetch image: ' + err.message });
        }
      });
    };
    
    fetchImage(imageUrl);
  } catch (err) {
    console.error('Proxy error:', err);
    res.status(500).json({ error: err.message || 'Proxy error' });
  }
});

app.post('/generate', async (req, res) => {
  try {
    const { data = {}, images = {}, logo_base64 } = req.body || {};
    
    // Debug logging
    console.log('Backend - received selected_calculators:', data.selected_calculators);
    console.log('Backend - received calculator_type:', data.calculator_type);
    console.log('Backend - calculator data keys:', Object.keys(data).filter(k => k.startsWith('calculator_')));

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


