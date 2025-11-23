// Detect environment: Electron vs Browser (avoid touching undefined require)
let ipcRenderer = null;
let path = null;
let fs = null;
let generatePDF = null;
(() => {
    const hasWindowRequire = typeof window !== 'undefined' && typeof window.require === 'function';
    const hasRequire = typeof require !== 'undefined';
    const r = hasWindowRequire ? window.require : (hasRequire ? require : null);
    if (!r) return; // browser mode
    try {
        ipcRenderer = r('electron').ipcRenderer;
        path = r('path');
        fs = r('fs');
        generatePDF = r('./pdf-generator').generatePDF;
    } catch (e) {
        // stay in browser mode
    }
})();

// Backend URL for web builds
// Use local backend if available, otherwise use hosted backend
let BACKEND_URL = (() => {
    // Allow manual override via localStorage for testing
    if (typeof Storage !== 'undefined' && localStorage.getItem('backend_url')) {
        return localStorage.getItem('backend_url');
    }
    
    // Check if we're in Electron (local development)
    if (ipcRenderer) {
        // Electron mode - use localhost for local development
        return 'http://localhost:8080';
    }
    
    // Check if we're running on localhost (browser mode)
    if (typeof window !== 'undefined' && window.location && 
        (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1')) {
        return 'http://localhost:8080';
    }
    
    // Default to hosted backend for production
    return 'https://pdf-generator-backend-fbtb.onrender.com';
})();

// Image storage
const imageSections = {
    cover: [],
    property: [],
    floor_plans: [],
    directions: [],
    city: []
};

const selectedImages = {};

// Tab switching
document.querySelectorAll('.tab-button').forEach(button => {
    button.addEventListener('click', () => {
        const tabName = button.dataset.tab;
        
        // Update buttons
        document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));
        button.classList.add('active');
        
        // Update content
        document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
        document.getElementById(`${tabName}-tab`).classList.add('active');
    });
});

// Image management functions
async function addImages(section) {
    // Browser mode: use <input type="file">
    if (!ipcRenderer) {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = 'image/*';
        input.multiple = true;
        input.onchange = () => {
            const files = Array.from(input.files || []);
            const items = files.map(file => ({
                url: URL.createObjectURL(file),
                name: file.name,
                file
            }));
            imageSections[section].push(...items);
            updateImageList(section);
        };
        input.click();
        return;
    }

    // Electron mode
    try {
        const filePaths = await ipcRenderer.invoke('select-images');
        if (filePaths && filePaths.length > 0) {
            imageSections[section].push(...filePaths);
            updateImageList(section);
        }
    } catch (error) {
        alert('Error adding images: ' + error.message);
    }
}

function updateImageList(section) {
    const list = document.getElementById(`${section}-images`);
    if (!list) return;
    
    list.innerHTML = '';
    imageSections[section].forEach((imageItem, index) => {
        const li = document.createElement('li');
        li.dataset.index = index;
        li.dataset.section = section;
        
        // Determine display name and URL (supports Electron paths or browser object URLs)
        let fileName = '';
        let fileUrl = '';
        if (typeof imageItem === 'string') {
            fileName = path ? path.basename(imageItem) : imageItem.split('/').pop();
            let normalizedPath = imageItem.replace(/\\/g, '/');
            if (normalizedPath.match(/^[A-Za-z]:/)) normalizedPath = '/' + normalizedPath;
            fileUrl = `file://${normalizedPath}`;
        } else {
            fileName = imageItem.name || 'image';
            fileUrl = imageItem.url;
        }
        const isSelected = selectedImages[section] === index;
        
        if (isSelected) {
            li.classList.add('selected');
        }
        
        li.innerHTML = `
            <img src="${fileUrl}" alt="${fileName}" onerror="this.style.display='none'" loading="lazy">
            <span>${fileName}</span>
        `;
        
        li.addEventListener('click', () => {
            // Toggle selection
            if (selectedImages[section] === index) {
                selectedImages[section] = null;
                li.classList.remove('selected');
            } else {
                // Deselect previous
                const prevSelected = list.querySelector('.selected');
                if (prevSelected) prevSelected.classList.remove('selected');
                
                selectedImages[section] = index;
                li.classList.add('selected');
            }
        });
        
        list.appendChild(li);
    });
}

function moveImageUp(section) {
    const selectedIndex = selectedImages[section];
    if (selectedIndex === null || selectedIndex === undefined || selectedIndex === 0) {
        return;
    }
    
    const images = imageSections[section];
    [images[selectedIndex - 1], images[selectedIndex]] = [images[selectedIndex], images[selectedIndex - 1]];
    selectedImages[section] = selectedIndex - 1;
    updateImageList(section);
}

function moveImageDown(section) {
    const selectedIndex = selectedImages[section];
    const images = imageSections[section];
    if (selectedIndex === null || selectedIndex === undefined || selectedIndex >= images.length - 1) {
        return;
    }
    
    [images[selectedIndex], images[selectedIndex + 1]] = [images[selectedIndex + 1], images[selectedIndex]];
    selectedImages[section] = selectedIndex + 1;
    updateImageList(section);
}

function removeSelectedImage(section) {
    const selectedIndex = selectedImages[section];
    if (selectedIndex === null || selectedIndex === undefined) {
        return;
    }
    
    imageSections[section].splice(selectedIndex, 1);
    selectedImages[section] = null;
    updateImageList(section);
}

// Make functions global for onclick handlers
window.addImages = addImages;
window.moveImageUp = moveImageUp;
window.moveImageDown = moveImageDown;
window.removeSelectedImage = removeSelectedImage;

// Get form data
function getFormData() {
    const data = {};
    
    // Get all input fields
    const inputs = document.querySelectorAll('input, textarea, select');
    inputs.forEach(input => {
        if (input.id) {
            data[input.id] = input.value.trim();
        }
    });
    
    return data;
}

// Clear all form data
function clearAll() {
    if (confirm('Are you sure you want to clear all data?')) {
        // Clear all inputs
        document.querySelectorAll('input, textarea, select').forEach(input => {
            if (input.type === 'checkbox') {
                input.checked = false;
            } else {
                input.value = '';
            }
        });
        
        // Clear all images
        Object.keys(imageSections).forEach(section => {
            imageSections[section] = [];
            selectedImages[section] = null;
            updateImageList(section);
        });
    }
}

window.clearAll = clearAll;

// Load mock data on page load
function loadMockData() {
    const mockData = {
        'address': '5, Ridley Road',
        'postal_code': 'L6 6DN',
        'property_type': 'Semi-Detached House',
        'bedrooms': '5',
        'bathrooms': '5',
        'size_sqm': '116',
        'asking_price': '£290,000',
        'days_on_market': '6',
        'key_features': 'Spacious Three Storey HMO Property\nFive Spacious En-Suite Double Bedrooms\nFantastic Investment Opportunity\nContemporary Fitted Kitchen\nCommunal Lounge\nSunny Rear Courtyard\nYield of 10.31%\nClose To Great Local Amenities, Train Station And Road Links\nClose To City Centre\nEPC GRADE = C',
        'description': 'Beautiful semi-detached family home in excellent condition. Features include modern kitchen, spacious living areas, and a well-maintained garden. Perfect for families looking for comfort and convenience. Located in a quiet residential area with excellent transport links.',
        'purchase_price': '£290,000',
        'deposit_percent': '20',
        'monthly_rent': '£2,750',
        'mortgage_rate': '5.8',
        'council_tax': '£1,670',
        'repairs_maintenance': '£660',
        'utilities': '£1,080',
        'water': '£300',
        'broadband_tv': '£480',
        'insurance': '£480',
        'stamp_duty': '£19,000',
        'survey_cost': '£800',
        'legal_fees': '£2,400',
        'loan_setup': '£4,640',
        'epc_rating': 'C',
        'current_energy_cost': '£1,200',
        'potential_energy_cost': '£800',
        'co2_current': '3.2',
        'co2_potential': '1.8',
        // Internet / Broadband mock data
        'broadband_available': 'Yes (FTTP available)',
        'download_speed': '1 Gbps',
        'upload_speed': '100 Mbps'
    };

    // Populate form fields
    Object.keys(mockData).forEach(fieldId => {
        const element = document.getElementById(fieldId);
        if (element) {
            if (element.tagName === 'TEXTAREA' || element.tagName === 'SELECT') {
                element.value = mockData[fieldId];
            } else {
                element.value = mockData[fieldId];
            }
        }
    });
}

// Load default images from sample_images folder
async function loadDefaultImages() {
    // Browser mode: preload sample images from the repo
    if (!ipcRenderer) {
        const base = 'sample_images/';
        const defaults = {
            cover: ['exterior_front.jpg'],
            property: ['kitchen.jpg', 'bathroom.jpg', 'bedroom.jpg', 'garden.jpeg', 'living_room.png'],
            floor_plans: ['floorplan1.png', 'floorplan2.png']
            // Note: directions and city sections are left empty - will be filled by fetchLocationData
        };
        Object.keys(defaults).forEach(section => {
            const files = defaults[section];
            files.forEach(name => {
                imageSections[section].push({
                    url: `${base}${name}`,
                    name
                });
            });
            updateImageList(section);
        });
        return;
    }

    // Electron mode: read from filesystem
    try {
        const appPath = await ipcRenderer.invoke('get-app-path');
        const sampleImagesPath = path.join(appPath, 'sample_images');
        if (!fs.existsSync(sampleImagesPath)) return;

        const files = fs.readdirSync(sampleImagesPath);
        const imageExtensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp'];
        const imageFiles = files.filter(file => imageExtensions.includes(path.extname(file).toLowerCase())).sort();

        imageFiles.forEach(filename => {
            const imagePath = path.join(sampleImagesPath, filename);
            const lowerFilename = filename.toLowerCase();
            let section = 'property';
            if (lowerFilename.includes('exterior') || lowerFilename.includes('front')) section = 'cover';
            else if (lowerFilename.includes('floor') || lowerFilename.includes('plan')) section = 'floor_plans';
            // Note: directions and city sections are skipped - will be filled by fetchLocationData
            // Don't auto-load directions or city images from sample_images folder
            if (lowerFilename.includes('direction') || lowerFilename.includes('map') || 
                lowerFilename.includes('liverpool') || lowerFilename.includes('city')) {
                return; // Skip these - they'll be fetched automatically
            }
            imageSections[section].push(imagePath);
        });
        Object.keys(imageSections).forEach(section => updateImageList(section));
    } catch (_err) {
        // ignore
    }
}

// Load mock data and images when page loads
document.addEventListener('DOMContentLoaded', () => {
    loadMockData();
    loadDefaultImages();
});

// Generate PDF
async function generatePDFFile() {
    try {
        // Browser: call backend and download PDF
        if (!ipcRenderer) {
            const data = getFormData();

            const fileToDataURL = (file) =>
                new Promise((resolve, reject) => {
                    const reader = new FileReader();
                    reader.onload = () => resolve(reader.result);
                    reader.onerror = reject;
                    reader.readAsDataURL(file);
                });

            const imagesPayload = {};
            const urlToDataURL = async (url) => {
                const res = await fetch(url, { mode: 'cors' });
                const blob = await res.blob();
                return await new Promise((resolve) => {
                    const reader = new FileReader();
                    reader.onload = () => resolve(reader.result);
                    reader.readAsDataURL(blob);
                });
            };
            const sections = Object.keys(imageSections);
            for (const section of sections) {
                const items = imageSections[section] || [];
                const b64s = [];
                for (const item of items) {
                    if (item && item.file) {
                        // Browser object with File
                        // eslint-disable-next-line no-await-in-loop
                        const b64 = await fileToDataURL(item.file);
                        b64s.push(b64);
                    } else if (item && item.url) {
                        // eslint-disable-next-line no-await-in-loop
                        const b64 = await urlToDataURL(item.url);
                        b64s.push(b64);
                    }
                }
                imagesPayload[section] = b64s;
            }

            // Include logo - try multiple paths
            let logoBase64 = null;
            const logoPaths = ['logo.png', './logo.png', '/logo.png'];
            for (const logoPath of logoPaths) {
                try {
                    logoBase64 = await urlToDataURL(logoPath);
                    console.log('Logo loaded from:', logoPath);
                    break;
                } catch (e) {
                    console.warn('Failed to load logo from', logoPath, e);
                }
            }
            if (!logoBase64) {
                console.warn('Logo not found, PDF will use placeholder or skip logo');
            }

            const resp = await fetch(`${BACKEND_URL}/generate`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ data, images: imagesPayload, logo_base64: logoBase64 })
            });
            if (!resp.ok) {
                const txt = await resp.text();
                throw new Error(`Backend error: ${resp.status} ${txt}`);
            }
            const blob = await resp.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `${(data.address || 'Property Report').replace(/[^a-z0-9 \-_]/gi, '')}.pdf`;
            document.body.appendChild(a);
            a.click();
            a.remove();
            URL.revokeObjectURL(url);
            return;
        }
        // Test IPC connection first
        try {
            await ipcRenderer.invoke('test-ipc');
        } catch (testError) {
            alert('Error: Cannot communicate with main process. Please restart the app.');
            return;
        }
        
        const data = getFormData();
        
        // Validate
        if (!data.address) {
            alert('Please enter at least the property address.');
            return;
        }
        
        // Get save location
        const defaultFilename = `${data.address} - Investment Report.pdf`;
        
        let filePath;
        try {
            filePath = await ipcRenderer.invoke('save-pdf', defaultFilename);
        } catch (error) {
            alert('Error opening file dialog: ' + error.message + '\n\nCheck console for details.');
            return;
        }
        
        if (!filePath || filePath === null || filePath === undefined) {
            return; // User cancelled
        }
        
        if (typeof filePath !== 'string') {
            alert('Error: Invalid file path. Please try again.');
            return;
        }
        
        if (filePath.trim() === '') {
            alert('Error: No file path selected. Please try again.');
            return;
        }
        
        // Get logo path
        const logoPath = await ipcRenderer.invoke('get-logo-path');
        
        // Helper function to download remote images and convert to file paths
        const downloadImageToFile = async (imageItem) => {
            let imageUrl = null;
            
            // Determine the URL or file path
            if (typeof imageItem === 'string') {
                // Check if it's a blob URL
                if (imageItem.startsWith('blob:')) {
                    // Handle blob URL - convert to buffer
                    try {
                        const response = await fetch(imageItem);
                        const blob = await response.blob();
                        const arrayBuffer = await blob.arrayBuffer();
                        const buffer = Buffer.from(arrayBuffer);
                        
                        const os = require('os');
                        const path = require('path');
                        const fs = require('fs');
                        
                        // Determine file extension from blob type
                        let ext = 'png';
                        if (blob.type) {
                            const typeMatch = blob.type.match(/\/(jpg|jpeg|png|gif|webp|bmp)/i);
                            if (typeMatch) {
                                ext = typeMatch[1].toLowerCase();
                                if (ext === 'jpeg') ext = 'jpg';
                            }
                        }
                        
                        const tempPath = path.join(os.tmpdir(), `img-${Date.now()}-${Math.random().toString(36).slice(2)}.${ext}`);
                        fs.writeFileSync(tempPath, buffer);
                        console.log('Downloaded blob URL to:', tempPath);
                        return tempPath;
                    } catch (err) {
                        console.error('Error downloading blob URL:', err);
                        return null;
                    }
                }
                // Check if it's a file path (not a URL)
                if (!imageItem.startsWith('http://') && !imageItem.startsWith('https://')) {
                    return imageItem; // Already a file path
                }
                imageUrl = imageItem;
            } else if (imageItem && imageItem.url) {
                // Handle blob URL
                if (imageItem.url.startsWith('blob:')) {
                    try {
                        const response = await fetch(imageItem.url);
                        const blob = await response.blob();
                        const arrayBuffer = await blob.arrayBuffer();
                        const buffer = Buffer.from(arrayBuffer);
                        
                        const os = require('os');
                        const path = require('path');
                        const fs = require('fs');
                        
                        let ext = 'png';
                        if (blob.type) {
                            const typeMatch = blob.type.match(/\/(jpg|jpeg|png|gif|webp|bmp)/i);
                            if (typeMatch) {
                                ext = typeMatch[1].toLowerCase();
                                if (ext === 'jpeg') ext = 'jpg';
                            }
                        }
                        
                        const tempPath = path.join(os.tmpdir(), `img-${Date.now()}-${Math.random().toString(36).slice(2)}.${ext}`);
                        fs.writeFileSync(tempPath, buffer);
                        console.log('Downloaded blob URL to:', tempPath);
                        return tempPath;
                    } catch (err) {
                        console.error('Error downloading blob URL:', err);
                        return null;
                    }
                }
                if (imageItem.url.startsWith('http://') || imageItem.url.startsWith('https://')) {
                    imageUrl = imageItem.url;
                } else {
                    return imageItem.url; // Local file path
                }
            } else {
                return imageItem; // Unknown format, return as-is
            }
            
            // Download the remote image
            if (imageUrl) {
                try {
                    // Clean URL - remove existing cache-busting params and add new one
                    let cleanUrl = imageUrl.split('&_cb=')[0].split('?_cb=')[0].split('&t=')[0].split('?t=')[0];
                    const separator = cleanUrl.includes('?') ? '&' : '?';
                    const cacheBustUrl = `${cleanUrl}${separator}_cb=${Date.now()}`;
                    
                    console.log('Downloading image:', cacheBustUrl);
                    const response = await fetch(cacheBustUrl, {
                        cache: 'no-store', // Force no cache
                        headers: {
                            'Cache-Control': 'no-cache',
                            'Pragma': 'no-cache'
                        }
                    });
                    if (!response.ok) {
                        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
                    }
                    const blob = await response.blob();
                    const arrayBuffer = await blob.arrayBuffer();
                    const buffer = Buffer.from(arrayBuffer);
                    console.log('Downloaded image size:', buffer.length, 'bytes');
                    
                    // Save to temp file
                    const os = require('os');
                    const path = require('path');
                    const fs = require('fs');
                    
                    // Determine file extension from URL or Content-Type
                    let ext = 'jpg';
                    const urlMatch = imageUrl.match(/\.(jpg|jpeg|png|gif|webp|bmp)(\?|$)/i);
                    if (urlMatch) {
                        ext = urlMatch[1].toLowerCase();
                    } else if (blob.type) {
                        const typeMatch = blob.type.match(/\/(jpg|jpeg|png|gif|webp|bmp)/i);
                        if (typeMatch) {
                            ext = typeMatch[1].toLowerCase();
                            if (ext === 'jpeg') ext = 'jpg';
                        }
                    }
                    
                    const tempPath = path.join(os.tmpdir(), `img-${Date.now()}-${Math.random().toString(36).slice(2)}.${ext}`);
                    fs.writeFileSync(tempPath, buffer);
                    console.log('Downloaded image to:', tempPath);
                    return tempPath;
                } catch (err) {
                    console.error('Error downloading image:', imageUrl, err);
                    return null;
                }
            }
            
            return null;
        };
        
        // Prepare image data - download remote URLs to temp files
        const images = {
            cover: [],
            property: [],
            floor_plans: [],
            directions: [],
            city: []
        };
        
        // Show loading message
        const loadingMsg = document.createElement('div');
        loadingMsg.style.cssText = 'position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);background:#1e3a8a;color:white;padding:20px 40px;border-radius:8px;z-index:10000;font-size:16px;';
        loadingMsg.textContent = 'Generating PDF...';
        document.body.appendChild(loadingMsg);
        
        try {
            // Download all remote images to temp files
            const sections = ['cover', 'property', 'floor_plans', 'directions', 'city'];
            for (const section of sections) {
                const items = imageSections[section] || [];
                console.log(`Processing ${section} section:`, items.length, 'items');
                for (const item of items) {
                    console.log(`Downloading ${section} image:`, typeof item === 'string' ? item : item.url);
                    const filePath = await downloadImageToFile(item);
                    if (filePath) {
                        images[section].push(filePath);
                        console.log(`Successfully downloaded ${section} image to:`, filePath);
                    } else {
                        console.warn(`Failed to download ${section} image:`, item);
                    }
                }
            }
            console.log('Final images for PDF:', images);
            
            // Generate PDF
            await generatePDF(data, images, filePath, logoPath);
            
            // Remove loading message
            if (document.body.contains(loadingMsg)) {
                document.body.removeChild(loadingMsg);
            }
        } catch (error) {
            if (document.body.contains(loadingMsg)) {
                document.body.removeChild(loadingMsg);
            }
            alert('Error generating PDF: ' + error.message + '\n\nCheck the console for details.');
            throw error;
        }
    } catch (error) {
        alert('Error generating PDF: ' + error.message);
    }
}

// Helper function to add timeout to fetch requests
function fetchWithTimeout(url, options = {}, timeout = 10000) {
    return Promise.race([
        fetch(url, options),
        new Promise((_, reject) => 
            setTimeout(() => reject(new Error('Request timeout')), timeout)
        )
    ]);
}

// Helper function to extract population from text
function extractPopulation(text) {
    if (!text) return '';
    // Look for patterns like "population of 500,000" or "508,986 inhabitants"
    const patterns = [
        /population (?:of|is|was|:)?\s*([\d,]+)/i,
        /([\d,]+)\s*(?:inhabitants|residents|people)/i,
        /([\d,]+)\s*population/i
    ];
    
    for (const pattern of patterns) {
        const match = text.match(pattern);
        if (match && match[1]) {
            return match[1].replace(/,/g, '');
        }
    }
    return '';
}

// Fetch Location Data function
async function fetchLocationData() {
    const addressInput = document.getElementById('address');
    const postalCodeInput = document.getElementById('postal_code');
    const fetchBtn = document.getElementById('fetch-location-btn');
    const fetchText = document.getElementById('fetch-location-text');
    const fetchSpinner = document.getElementById('fetch-location-spinner');
    
    // Get address and postal code
    const address = addressInput?.value.trim() || '';
    const postalCode = postalCodeInput?.value.trim() || '';
    
    if (!address && !postalCode) {
        alert('Please enter a property address or postal code first.');
        return;
    }
    
    // Build search query
    const searchQuery = [address, postalCode].filter(Boolean).join(', ');
    
    // Show loading state
    fetchBtn.disabled = true;
    fetchText.style.display = 'none';
    fetchSpinner.style.display = 'inline';
    
    try {
        // Step 1: Geocode the address using Nominatim (OpenStreetMap)
        const geocodeUrl = `https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(searchQuery)}&limit=1&addressdetails=1`;
        
        const geocodeResponse = await fetchWithTimeout(geocodeUrl, {
            headers: {
                'User-Agent': 'PropertyPDFBuilder/1.0'
            }
        }, 8000);
        
        if (!geocodeResponse.ok) {
            throw new Error('Failed to geocode address');
        }
        
        const geocodeData = await geocodeResponse.json();
        
        if (!geocodeData || geocodeData.length === 0) {
            throw new Error('Address not found. Please check the address and postal code.');
        }
        
        const location = geocodeData[0];
        const lat = parseFloat(location.lat);
        const lon = parseFloat(location.lon);
        const addressDetails = location.address || {};
        
        // Extract city information
        const city = addressDetails.city || addressDetails.town || addressDetails.village || 
                     addressDetails.municipality || addressDetails.county || '';
        
        // Step 2: Run all API calls in parallel for better performance
        const apiPromises = [];
        
        // Wikipedia data (for city info and population)
        let wikiPromise = Promise.resolve({ aboutCity: '', population: '' });
        if (city) {
            wikiPromise = fetchWithTimeout(
                `https://en.wikipedia.org/api/rest_v1/page/summary/${encodeURIComponent(city)}`,
                {},
                8000
            )
            .then(async (response) => {
                if (response.ok) {
                    const wikiData = await response.json();
                    let aboutCity = '';
                    let population = '';
                    
                    if (wikiData.extract) {
                        aboutCity = wikiData.extract.split('. ').slice(0, 3).join('. ') + '.';
                        // Try to extract population from extract text
                        population = extractPopulation(wikiData.extract);
                    }
                    
                    // Also check the full text for population if not found
                    if (!population && wikiData.extract) {
                        const fullText = wikiData.extract.toLowerCase();
                        // Look for more specific patterns
                        const popMatch = fullText.match(/(?:population|inhabitants|residents)[^\d]*([\d,]+)/i);
                        if (popMatch) {
                            population = popMatch[1].replace(/,/g, '');
                        }
                    }
                    
                    return { aboutCity, population };
                }
                return { aboutCity: '', population: '' };
            })
            .catch(() => ({ aboutCity: '', population: '' }));
        }
        apiPromises.push(wikiPromise);
        
        // Overpass queries (stations, schools, amenities) - run in parallel
        const overpassUrl = 'https://overpass-api.de/api/interpreter';
        
        // Station query - improved syntax
        const stationQuery = `[out:json][timeout:15];
(
  node["railway"="station"](around:5000,${lat},${lon});
  node["public_transport"="station"](around:5000,${lat},${lon});
  way["railway"="station"](around:5000,${lat},${lon});
  relation["railway"="station"](around:5000,${lat},${lon});
);
out center meta;`;
        
        const stationPromise = fetchWithTimeout(overpassUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
            },
            body: `data=${encodeURIComponent(stationQuery)}`
        }, 15000)
        .then(async (response) => {
            if (!response.ok) {
                console.warn('Station API response not OK:', response.status, response.statusText);
                return { nearestStation: '', stationDistance: '' };
            }
            
            const data = await response.json();
            console.log('Station API response:', data);
            
            if (data.elements && data.elements.length > 0) {
                let closestStation = null;
                let minDistance = Infinity;
                
                data.elements.forEach(element => {
                    // Handle different element types
                    let stationLat, stationLon;
                    
                    if (element.type === 'node') {
                        stationLat = element.lat;
                        stationLon = element.lon;
                    } else if (element.center) {
                        stationLat = element.center.lat;
                        stationLon = element.center.lon;
                    } else if (element.lat && element.lon) {
                        stationLat = element.lat;
                        stationLon = element.lon;
                    }
                    
                    if (stationLat && stationLon && element.tags && element.tags.name) {
                        const distance = calculateDistance(lat, lon, stationLat, stationLon);
                        if (distance < minDistance) {
                            minDistance = distance;
                            closestStation = {
                                name: element.tags.name,
                                distance: distance
                            };
                        }
                    }
                });
                
                if (closestStation) {
                    console.log('Found station:', closestStation);
                    return {
                        nearestStation: closestStation.name,
                        stationDistance: closestStation.distance.toFixed(1)
                    };
                } else {
                    console.warn('No valid station found in results');
                }
            } else {
                console.warn('No station elements in response');
            }
            return { nearestStation: '', stationDistance: '' };
        })
        .catch(async (error) => {
            console.error('Error fetching station data from Overpass:', error);
            // Fallback: Try Nominatim search for railway stations
            try {
                const nominatimUrl = `https://nominatim.openstreetmap.org/search?format=json&q=railway+station+near+${lat},${lon}&limit=5&addressdetails=1`;
                const fallbackResponse = await fetchWithTimeout(nominatimUrl, {
                    headers: {
                        'User-Agent': 'PropertyPDFBuilder/1.0'
                    }
                }, 8000);
                
                if (fallbackResponse.ok) {
                    const fallbackData = await fallbackResponse.json();
                    if (fallbackData && fallbackData.length > 0) {
                        // Find closest station
                        let closest = null;
                        let minDist = Infinity;
                        
                        fallbackData.forEach(item => {
                            const itemLat = parseFloat(item.lat);
                            const itemLon = parseFloat(item.lon);
                            if (itemLat && itemLon && item.display_name) {
                                const dist = calculateDistance(lat, lon, itemLat, itemLon);
                                if (dist < minDist && (item.type === 'railway' || item.display_name.toLowerCase().includes('station'))) {
                                    minDist = dist;
                                    closest = {
                                        name: item.display_name.split(',')[0].trim(),
                                        distance: dist
                                    };
                                }
                            }
                        });
                        
                        if (closest) {
                            console.log('Found station via Nominatim fallback:', closest);
                            return {
                                nearestStation: closest.name,
                                stationDistance: closest.distance.toFixed(1)
                            };
                        }
                    }
                }
            } catch (fallbackError) {
                console.warn('Nominatim fallback also failed:', fallbackError);
            }
            return { nearestStation: '', stationDistance: '' };
        });
        apiPromises.push(stationPromise);
        
        // School query - improved syntax
        const schoolQuery = `[out:json][timeout:15];
(
  node["amenity"="school"](around:5000,${lat},${lon});
  way["amenity"="school"](around:5000,${lat},${lon});
  relation["amenity"="school"](around:5000,${lat},${lon});
);
out center meta;`;
        
        const schoolPromise = fetchWithTimeout(overpassUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
            },
            body: `data=${encodeURIComponent(schoolQuery)}`
        }, 15000)
        .then(async (response) => {
            if (!response.ok) {
                console.warn('School API response not OK:', response.status, response.statusText);
                return { nearestSchool: '', schoolDistance: '' };
            }
            
            const data = await response.json();
            console.log('School API response:', data);
            
            if (data.elements && data.elements.length > 0) {
                let closestSchool = null;
                let minDistance = Infinity;
                
                data.elements.forEach(element => {
                    // Handle different element types
                    let schoolLat, schoolLon;
                    
                    if (element.type === 'node') {
                        schoolLat = element.lat;
                        schoolLon = element.lon;
                    } else if (element.center) {
                        schoolLat = element.center.lat;
                        schoolLon = element.center.lon;
                    } else if (element.lat && element.lon) {
                        schoolLat = element.lat;
                        schoolLon = element.lon;
                    }
                    
                    if (schoolLat && schoolLon && element.tags && element.tags.name) {
                        const distance = calculateDistance(lat, lon, schoolLat, schoolLon);
                        if (distance < minDistance) {
                            minDistance = distance;
                            closestSchool = {
                                name: element.tags.name,
                                distance: distance
                            };
                        }
                    }
                });
                
                if (closestSchool) {
                    console.log('Found school:', closestSchool);
                    return {
                        nearestSchool: closestSchool.name,
                        schoolDistance: closestSchool.distance.toFixed(1)
                    };
                } else {
                    console.warn('No valid school found in results');
                }
            } else {
                console.warn('No school elements in response');
            }
            return { nearestSchool: '', schoolDistance: '' };
        })
        .catch(async (error) => {
            console.error('Error fetching school data from Overpass:', error);
            // Fallback: Try Nominatim search for schools
            try {
                const nominatimUrl = `https://nominatim.openstreetmap.org/search?format=json&q=school+near+${lat},${lon}&limit=5&addressdetails=1`;
                const fallbackResponse = await fetchWithTimeout(nominatimUrl, {
                    headers: {
                        'User-Agent': 'PropertyPDFBuilder/1.0'
                    }
                }, 8000);
                
                if (fallbackResponse.ok) {
                    const fallbackData = await fallbackResponse.json();
                    if (fallbackData && fallbackData.length > 0) {
                        // Find closest school
                        let closest = null;
                        let minDist = Infinity;
                        
                        fallbackData.forEach(item => {
                            const itemLat = parseFloat(item.lat);
                            const itemLon = parseFloat(item.lon);
                            if (itemLat && itemLon && item.display_name) {
                                const dist = calculateDistance(lat, lon, itemLat, itemLon);
                                if (dist < minDist && (item.type === 'school' || item.display_name.toLowerCase().includes('school'))) {
                                    minDist = dist;
                                    closest = {
                                        name: item.display_name.split(',')[0].trim(),
                                        distance: dist
                                    };
                                }
                            }
                        });
                        
                        if (closest) {
                            console.log('Found school via Nominatim fallback:', closest);
                            return {
                                nearestSchool: closest.name,
                                schoolDistance: closest.distance.toFixed(1)
                            };
                        }
                    }
                }
            } catch (fallbackError) {
                console.warn('Nominatim fallback also failed:', fallbackError);
            }
            return { nearestSchool: '', schoolDistance: '' };
        });
        apiPromises.push(schoolPromise);
        
        // City centre distance - also store coordinates for map
        let cityCentrePromise = Promise.resolve({ cityCentreDistance: '', cityLat: null, cityLon: null });
        if (city && lat && lon) {
            cityCentrePromise = fetchWithTimeout(
                `https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(city + ', UK')}&limit=1`,
                {
                    headers: {
                        'User-Agent': 'PropertyPDFBuilder/1.0'
                    }
                },
                8000
            )
            .then(async (response) => {
                if (response.ok) {
                    const cityData = await response.json();
                    if (cityData && cityData.length > 0) {
                        const cityLat = parseFloat(cityData[0].lat);
                        const cityLon = parseFloat(cityData[0].lon);
                        const distance = calculateDistance(lat, lon, cityLat, cityLon);
                        return { 
                            cityCentreDistance: distance.toFixed(1),
                            cityLat: cityLat,
                            cityLon: cityLon
                        };
                    }
                }
                return { cityCentreDistance: '', cityLat: null, cityLon: null };
            })
            .catch(() => ({ cityCentreDistance: '', cityLat: null, cityLon: null }));
        }
        apiPromises.push(cityCentrePromise);
        
        // Amenities query
        const amenitiesQuery = `
            [out:json][timeout:10];
            (
              node["amenity"~"^(restaurant|cafe|supermarket|pharmacy|hospital|bank|library|park)$"](around:2000,${lat},${lon});
              way["amenity"~"^(restaurant|cafe|supermarket|pharmacy|hospital|bank|library|park)$"](around:2000,${lat},${lon});
            );
            out center;
        `;
        
        const amenitiesPromise = fetchWithTimeout(overpassUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
            },
            body: `data=${encodeURIComponent(amenitiesQuery)}`
        }, 10000)
        .then(async (response) => {
            if (response.ok) {
                const data = await response.json();
                if (data.elements && data.elements.length > 0) {
                    const amenityTypes = {};
                    data.elements.slice(0, 10).forEach(element => {
                        const amenityType = element.tags?.amenity;
                        if (amenityType && !amenityTypes[amenityType]) {
                            amenityTypes[amenityType] = true;
                        }
                    });
                    const amenities = Object.keys(amenityTypes).map(t => t.charAt(0).toUpperCase() + t.slice(1));
                    return { localAmenities: amenities };
                }
            }
            return { localAmenities: [] };
        })
        .catch(() => ({ localAmenities: [] }));
        apiPromises.push(amenitiesPromise);
        
        // Wait for all API calls to complete (in parallel)
        const results = await Promise.allSettled(apiPromises);
        
        // Extract results with better error handling
        const wikiResult = results[0]?.status === 'fulfilled' ? results[0].value : { aboutCity: '', population: '' };
        const stationResult = results[1]?.status === 'fulfilled' ? results[1].value : { nearestStation: '', stationDistance: '' };
        const schoolResult = results[2]?.status === 'fulfilled' ? results[2].value : { nearestSchool: '', schoolDistance: '' };
        const cityCentreResult = results[3]?.status === 'fulfilled' ? results[3].value : { cityCentreDistance: '' };
        const amenitiesResult = results[4]?.status === 'fulfilled' ? results[4].value : { localAmenities: [] };
        
        // Log results for debugging
        console.log('Location fetch results:', {
            city,
            wikiResult,
            stationResult,
            schoolResult,
            cityCentreResult,
            amenitiesResult
        });
        
        // Populate form fields - always try to set values, even if empty
        const cityField = document.getElementById('city');
        if (cityField && city) {
            cityField.value = city;
        }
        
        const cityCentreField = document.getElementById('city_centre_distance');
        if (cityCentreField) {
            cityCentreField.value = cityCentreResult.cityCentreDistance || '';
        }
        
        const stationField = document.getElementById('nearest_station');
        if (stationField) {
            stationField.value = stationResult.nearestStation || '';
        }
        
        const stationDistField = document.getElementById('station_distance');
        if (stationDistField) {
            stationDistField.value = stationResult.stationDistance || '';
        }
        
        const schoolField = document.getElementById('nearest_school');
        if (schoolField) {
            schoolField.value = schoolResult.nearestSchool || '';
        }
        
        const schoolDistField = document.getElementById('school_distance');
        if (schoolDistField) {
            schoolDistField.value = schoolResult.schoolDistance || '';
        }
        
        const amenitiesField = document.getElementById('local_amenities');
        if (amenitiesField) {
            if (amenitiesResult.localAmenities && amenitiesResult.localAmenities.length > 0) {
                amenitiesField.value = `Nearby amenities include: ${amenitiesResult.localAmenities.join(', ')}.`;
            } else {
                amenitiesField.value = '';
            }
        }
        
        const aboutCityField = document.getElementById('about_city');
        if (aboutCityField) {
            aboutCityField.value = wikiResult.aboutCity || '';
        }
        
        const populationField = document.getElementById('population');
        if (populationField) {
            populationField.value = wikiResult.population || '';
        }
        
        // Step 8: Fetch map and city images automatically
        if (lat && lon && city) {
            try {
                // Fetch map image of the capital city center (not the property location)
                // Use the city center coordinates we already fetched
                let cityLat = cityCentreResult.cityLat || lat;
                let cityLon = cityCentreResult.cityLon || lon;
                
                // If we don't have city center coordinates, fetch them
                if (!cityCentreResult.cityLat || !cityCentreResult.cityLon) {
                    try {
                        const cityGeocodeUrl = `https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(city + ', UK')}&limit=1`;
                        const cityResponse = await fetchWithTimeout(cityGeocodeUrl, {
                            headers: {
                                'User-Agent': 'PropertyPDFBuilder/1.0'
                            }
                        }, 8000);
                        
                        if (cityResponse.ok) {
                            const cityData = await cityResponse.json();
                            if (cityData && cityData.length > 0) {
                                cityLat = parseFloat(cityData[0].lat);
                                cityLon = parseFloat(cityData[0].lon);
                                console.log('Fetched city center coordinates for map:', cityLat, cityLon);
                            }
                        }
                    } catch (e) {
                        console.warn('Could not fetch city center coordinates, using property location:', e);
                    }
                } else {
                    console.log('Using cached city center coordinates for map:', cityLat, cityLon);
                }
                
                // Fetch map image - create a composite from multiple tiles for a larger, wider view
                // Use city center coordinates for the map (shows the capital city, not the property)
                
                const z = 13; // Zoom level 13 for good detail without being too zoomed in
                const centerX = Math.floor((cityLon + 180) / 360 * Math.pow(2, z));
                const centerY = Math.floor((1 - Math.log(Math.tan(cityLat * Math.PI / 180) + 1 / Math.cos(cityLat * Math.PI / 180)) / Math.PI) / 2 * Math.pow(2, z));
                
                // Create a 5x3 grid of tiles (15 tiles total) - wider format
                // This gives us 1280x768 pixels which scales well to match city images container width
                const tileSize = 256; // Each tile is 256x256 pixels
                const gridWidth = 5; // 5 tiles wide
                const gridHeight = 3; // 3 tiles tall
                const compositeWidth = tileSize * gridWidth; // 1280 pixels wide
                const compositeHeight = tileSize * gridHeight; // 768 pixels tall
                
                // Create canvas to composite the tiles
                const canvas = document.createElement('canvas');
                canvas.width = compositeWidth;
                canvas.height = compositeHeight;
                const ctx = canvas.getContext('2d');
                
                // Download and composite tiles in a 5x3 grid
                const tilePromises = [];
                for (let dy = -1; dy <= 1; dy++) { // 3 rows: -1, 0, 1
                    for (let dx = -2; dx <= 2; dx++) { // 5 columns: -2, -1, 0, 1, 2
                        const tileX = centerX + dx;
                        const tileY = centerY + dy;
                        const tileUrl = `https://tile.openstreetmap.org/${z}/${tileX}/${tileY}.png`;
                        
                        const tilePromise = fetch(tileUrl, { cache: 'no-store' })
                            .then(response => {
                                if (!response.ok) {
                                    throw new Error(`HTTP ${response.status}`);
                                }
                                return response.blob();
                            })
                            .then(blob => {
                                return new Promise((resolve) => {
                                    const img = new Image();
                                    img.crossOrigin = 'anonymous';
                                    img.onload = () => {
                                        // Calculate position: center tile (dx=0, dy=0) is at (2, 1) in 5x3 grid
                                        const x = (dx + 2) * tileSize;
                                        const y = (dy + 1) * tileSize;
                                        ctx.drawImage(img, x, y, tileSize, tileSize);
                                        console.log(`Loaded tile at (${dx}, ${dy}) -> (${x}, ${y})`);
                                        resolve({ success: true, dx, dy });
                                    };
                                    img.onerror = (err) => {
                                        console.warn('Failed to load tile image:', tileUrl, err);
                                        resolve({ success: false, dx, dy }); // Continue even if one tile fails
                                    };
                                    img.src = URL.createObjectURL(blob);
                                });
                            })
                            .catch(err => {
                                console.warn('Error fetching tile:', tileUrl, err);
                                return { success: false, dx, dy };
                            });
                        
                        tilePromises.push(tilePromise);
                    }
                }
                
                // Wait for all tiles to load and composite (use allSettled to continue even if some fail)
                const tileResults = await Promise.allSettled(tilePromises);
                const loadedCount = tileResults.filter(r => r.status === 'fulfilled' && r.value?.success).length;
                console.log(`Loaded ${loadedCount} out of ${tilePromises.length} map tiles`);
                
                // Give a small delay to ensure all images are fully rendered on canvas
                await new Promise(resolve => setTimeout(resolve, 100));
                
                // Verify canvas has content
                const imageData = ctx.getImageData(0, 0, Math.min(100, compositeWidth), Math.min(100, compositeHeight));
                const hasContent = imageData.data.some((val, idx) => idx % 4 !== 3 || val !== 0); // Check if not all transparent
                console.log('Canvas has content:', hasContent, 'Canvas size:', compositeWidth, 'x', compositeHeight);
                
                // Convert canvas to blob URL (wait for it)
                const compositeUrl = await new Promise((resolve, reject) => {
                    canvas.toBlob((blob) => {
                        if (blob) {
                            const url = URL.createObjectURL(blob);
                            console.log('Created blob URL for map, blob size:', blob.size, 'bytes');
                            resolve(url);
                        } else {
                            reject(new Error('Failed to create blob from canvas'));
                        }
                    }, 'image/png');
                });
                
                // Clear existing directions images first (force clear)
                imageSections.directions.length = 0;
                
                // Add composite map image to directions section
                imageSections.directions.push({
                    url: compositeUrl,
                    name: `${city} Map`
                });
                
                // Force update the image list
                const directionsList = document.getElementById('directions-images');
                if (directionsList) {
                    directionsList.innerHTML = '';
                }
                updateImageList('directions');
                
                console.log('Added composite map image to directions section (5x3 tiles,', compositeWidth, 'x', compositeHeight, 'pixels, zoom level', z, ')');
                updateImageList('directions');
                console.log('Added composite map image to directions section. Total directions images:', imageSections.directions.length);
                
                // Fetch city images from Unsplash (CORS-compatible and reliable)
                // Using Unsplash directly to avoid CORS issues with Wikimedia Commons
                const fetchCityImages = async () => {
                    // Use Unsplash images directly - they support CORS and work from any origin
                    // Using a variety of city/urban/architecture photos
                    return [
                        { 
                            url: `https://images.unsplash.com/photo-1514565131-fce0801e5785?w=800&h=600&fit=crop&q=80&auto=format`, 
                            name: `${city} - City View 1` 
                        },
                        { 
                            url: `https://images.unsplash.com/photo-1449824913935-59a10b8d2000?w=800&h=600&fit=crop&q=80&auto=format`, 
                            name: `${city} - City View 2` 
                        },
                        { 
                            url: `https://images.unsplash.com/photo-1477959858617-67f85cf4f1df?w=800&h=600&fit=crop&q=80&auto=format`, 
                            name: `${city} - City View 3` 
                        }
                    ];
                };
                
                // Add city images to city section (replace any existing)
                fetchCityImages().then(cityImages => {
                    imageSections.city = cityImages; // Replace all existing images
                    updateImageList('city');
                    console.log('Added city images to city section:', cityImages.length);
                }).catch(err => {
                    console.warn('Error fetching city images:', err);
                });
            } catch (imageError) {
                console.warn('Could not fetch images:', imageError);
            }
        }
        
    } catch (error) {
        console.error('Error fetching location data:', error);
        alert('Error fetching location data: ' + error.message + '\n\nPlease try again or fill in the fields manually.');
    } finally {
        // Reset button state
        fetchBtn.disabled = false;
        fetchText.style.display = 'inline';
        fetchSpinner.style.display = 'none';
    }
}

// Helper function to calculate distance between two coordinates (in miles)
function calculateDistance(lat1, lon1, lat2, lon2) {
    const R = 3959; // Earth's radius in miles
    const dLat = (lat2 - lat1) * Math.PI / 180;
    const dLon = (lon2 - lon1) * Math.PI / 180;
    const a = 
        Math.sin(dLat / 2) * Math.sin(dLat / 2) +
        Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
        Math.sin(dLon / 2) * Math.sin(dLon / 2);
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    return R * c;
}

// Make functions available globally for onclick handlers
window.generatePDF = generatePDFFile;
window.generatePDFFile = generatePDFFile; // Also expose with full name
window.fetchLocationData = fetchLocationData;

