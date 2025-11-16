const { ipcRenderer } = require('electron');
const path = require('path');
const fs = require('fs');
const { generatePDF } = require('./pdf-generator');

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
    imageSections[section].forEach((imagePath, index) => {
        const li = document.createElement('li');
        li.dataset.index = index;
        li.dataset.section = section;
        
        const fileName = path.basename(imagePath);
        const isSelected = selectedImages[section] === index;
        
        if (isSelected) {
            li.classList.add('selected');
        }
        
        // Normalize file path for file:// URL (Windows uses backslashes)
        // Convert backslashes to forward slashes
        let normalizedPath = imagePath.replace(/\\/g, '/');
        
        // Handle Windows absolute paths (C:/path -> /C:/path for file://)
        // macOS/Linux paths already start with /
        if (normalizedPath.match(/^[A-Za-z]:/)) {
            // Windows absolute path (e.g., C:/Users/...)
            normalizedPath = '/' + normalizedPath;
        }
        
        const fileUrl = `file://${normalizedPath}`;
        
        li.innerHTML = `
            <img src="${fileUrl}" alt="${fileName}" onerror="this.style.display='none'">
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
        'upload_speed': '100 Mbps',
        // City info
        'city': 'Liverpool',
        'about_city': 'Liverpool is a port city and metropolitan borough in Merseyside, England. It is the administrative, cultural and economic centre of the Liverpool City Region.',
        'population': '508,986',
        'city_centre_distance': '1.8',
        'nearest_station': 'Liverpool Central',
        'station_distance': '0.5',
        'nearest_school': 'St. Mary\'s Primary',
        'school_distance': '0.3',
        'local_amenities': 'Close to shops, restaurants, and public transport. Excellent local amenities including supermarkets, cafes, and parks.'
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
    const fs = require('fs');
    const path = require('path');
    
    try {
        // Get the app directory from main process
        const appPath = await ipcRenderer.invoke('get-app-path');
        const sampleImagesPath = path.join(appPath, 'sample_images');
        
        // Check if sample_images folder exists
        if (!fs.existsSync(sampleImagesPath)) {
            return;
        }

        // Read all image files
        const files = fs.readdirSync(sampleImagesPath);
        const imageExtensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp'];
        
        const imageFiles = files.filter(file => {
            const ext = path.extname(file).toLowerCase();
            return imageExtensions.includes(ext);
        }).sort();

        // Categorize images based on filename
        imageFiles.forEach(filename => {
            const imagePath = path.join(sampleImagesPath, filename);
            const lowerFilename = filename.toLowerCase();
            
            let section = 'property'; // Default section
            
            if (lowerFilename.includes('exterior') || lowerFilename.includes('front')) {
                section = 'cover';
            } else if (lowerFilename.includes('floor') || lowerFilename.includes('plan')) {
                section = 'floor_plans';
            } else if (lowerFilename.includes('direction') || lowerFilename.includes('map')) {
                section = 'directions';
            } else if (lowerFilename.includes('liverpool') || lowerFilename.includes('city')) {
                section = 'city';
            }
            
            // Add to appropriate section
            if (!imageSections[section]) {
                imageSections[section] = [];
            }
            imageSections[section].push(imagePath);
        });

        // Update all image lists
        Object.keys(imageSections).forEach(section => {
            updateImageList(section);
        });
    } catch (error) {
        // Silent fail; UI already has manual image upload
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
        
        // Prepare image data
        const images = {
            cover: imageSections.cover,
            property: imageSections.property,
            floor_plans: imageSections.floor_plans,
            directions: imageSections.directions,
            city: imageSections.city
        };
        
        // Show loading message
        const loadingMsg = document.createElement('div');
        loadingMsg.style.cssText = 'position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);background:#1e3a8a;color:white;padding:20px 40px;border-radius:8px;z-index:10000;font-size:16px;';
        loadingMsg.textContent = 'Generating PDF...';
        document.body.appendChild(loadingMsg);
        
        try {
            // Generate PDF
            await generatePDF(data, images, filePath, logoPath);
            
            // Remove loading message
            if (document.body.contains(loadingMsg)) {
                document.body.removeChild(loadingMsg);
            }
            
            alert('PDF generated successfully!\n\nSaved to: ' + filePath);
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

// Make functions available globally for onclick handlers
window.generatePDF = generatePDFFile;
window.generatePDFFile = generatePDFFile; // Also expose with full name

