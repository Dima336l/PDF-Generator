const PDFDocument = require('pdfkit');
const fs = require('fs');
const path = require('path');

// Import calculator logic
let calculateInvestment, parseCurrencyCalc, formatCurrencyCalc;
try {
    const calcLogic = require('./calculator-logic');
    calculateInvestment = calcLogic.calculateInvestment;
    parseCurrencyCalc = calcLogic.parseCurrency;
    formatCurrencyCalc = calcLogic.formatCurrency;
} catch (e) {
    console.warn('Calculator logic not found, using fallback functions:', e.message);
    parseCurrencyCalc = parseCurrency;
    formatCurrencyCalc = formatCurrency;
    calculateInvestment = null;
}

// Constants
const A4_WIDTH = 595.28; // A4 width in points
const A4_HEIGHT = 841.89; // A4 height in points
const INCH = 72; // 1 inch = 72 points
const MARGIN = 0.75 * INCH;
const HEADER_TOP_OFFSET = 0.45 * INCH;

// Colors
const PRIMARY_BLUE = '#1e3a8a';
const ACCENT_GOLD = '#f59e0b';
const LIGHT_GREY = '#f8fafc';
const DARK_GREY = '#374151';
const SUCCESS_GREEN = '#10b981';

// Helper function to format date with ordinal
function formatDateWithOrdinal(date) {
    const day = date.getDate();
    let suffix = 'th';
    if (day % 10 === 1 && day % 100 !== 11) suffix = 'st';
    else if (day % 10 === 2 && day % 100 !== 12) suffix = 'nd';
    else if (day % 10 === 3 && day % 100 !== 13) suffix = 'rd';
    
    const months = ['January', 'February', 'March', 'April', 'May', 'June',
                   'July', 'August', 'September', 'October', 'November', 'December'];
    return `${day}${suffix} ${months[date.getMonth()]} ${date.getFullYear()}`;
}

// Helper function to parse currency
function parseCurrency(value) {
    if (!value) return 0;
    return parseFloat(value.toString().replace(/[£,\s]/g, '')) || 0;
}

// Helper function to format currency
function formatCurrency(value) {
    return `£${Math.round(value).toLocaleString()}`;
}

// Helper function to get actual image dimensions (synchronous using PDFKit's built-in method)
function getImageDimensions(imagePath) {
    try {
        if (!imagePath || !fs.existsSync(imagePath)) {
            return { width: 800, height: 600 };
        }
        // Use a simple approach - PDFKit will handle aspect ratio automatically
        // We'll calculate from a test image load, but for now return default
        // The actual sizing will be handled by PDFKit's image() method
        return { width: 800, height: 600 };
    } catch (error) {
        return { width: 800, height: 600 };
    }
}

// Helper function to add image to PDF with proper sizing
function addImageToPDF(doc, imagePath, x, y, maxWidth, maxHeight, fit = 'contain') {
    try {
        if (!fs.existsSync(imagePath) || !imagePath) {
            // Draw placeholder
            doc.rect(x, y, maxWidth, maxHeight)
               .fillColor('#CCCCCC')
               .fill();
            return;
        }

        // Respect aspect ratio using PDFKit's fit/cover options
        const options = { align: 'center', valign: 'center' };
        if (fit === 'cover') {
            options.cover = [maxWidth, maxHeight];
        } else if (fit === 'contain') {
            options.fit = [maxWidth, maxHeight];
        } else if (fit === 'width') {
            options.width = maxWidth;
        } else if (fit === 'height') {
            options.height = maxHeight;
        } else {
            options.fit = [maxWidth, maxHeight];
        }
        // Clip to the bounding box to prevent bleed into spacing
        doc.save();
        doc.rect(x, y, maxWidth, maxHeight).clip();
        doc.image(imagePath, x, y, options);
        doc.restore();
    } catch (error) {
        // Draw placeholder on error
        doc.rect(x, y, maxWidth, maxHeight)
           .fillColor('#CCCCCC')
           .fill();
    }
}

// Draw header on every page with proper logo size (1.4 inch width exactly)
function drawHeader(doc, logoPath) {
    const logoWidth = 1.4 * INCH; // Fixed width as per original - exactly 1.4 inches
    // PDFKit will maintain aspect ratio automatically, so we only specify width
    const logoX = MARGIN;
    const logoY = HEADER_TOP_OFFSET;

    // Draw logo if available (exact size: 1.4 inch width, height auto-calculated)
    if (logoPath && fs.existsSync(logoPath)) {
        try {
            // Only specify width - PDFKit maintains aspect ratio
            doc.image(logoPath, logoX, logoY, { 
                width: logoWidth
                // height is auto-calculated to maintain aspect ratio
            });
        } catch (error) {
        }
    }

    // Calculate tagline position - need to estimate logo height
    // Most logos have aspect ratio around 2:1 to 3:1, so estimate height
    // We'll use a conservative estimate and adjust if needed
    const estimatedLogoHeight = logoWidth * 0.4; // Estimate based on typical logo proportions
    const taglineY = logoY + estimatedLogoHeight + 12;
    
    // Draw tagline - "Elevating Your Property Experience" (as per original)
    doc.fontSize(9)
       .font('Helvetica')
       .fillColor('#334155') // Dark grey as per original
       .text('Elevating Your Property Experience', logoX, taglineY);
}

// Create cover page matching original exactly
function createCoverPage(doc, data, images, logoPath) {
    const pageWidth = A4_WIDTH;
    const pageHeight = A4_HEIGHT;
    const contentWidth = pageWidth - (2 * MARGIN);

    // Draw header with logo and tagline
    drawHeader(doc, logoPath);

    // Calculate content area (below header)
    // Header: top offset (0.45") + estimated logo height + tagline spacing (12pt) + tagline font (9pt) + section spacing (20pt)
    const logoWidth = 1.4 * INCH;
    const estimatedLogoHeight = logoWidth * 0.4; // Estimate based on typical logo proportions
    const STANDARD_SECTION_SPACING = 20; // Match Python: space after section title/header
    const headerHeight = HEADER_TOP_OFFSET + estimatedLogoHeight + 12 + 9 + STANDARD_SECTION_SPACING;
    const contentY = headerHeight;
    const availableHeight = pageHeight - contentY - MARGIN;

    // Property Address (large, bold, 24pt)
    const address = `${data.address || ''}${data.postal_code ? ', ' + data.postal_code : ''}`;
    doc.fontSize(24)
       .font('Helvetica-Bold')
       .fillColor('#000000')
       .text(address, MARGIN, contentY, {
           width: contentWidth,
           align: 'left'
       });

    const addressHeight = doc.heightOfString(address, { width: contentWidth });
    let currentY = contentY + addressHeight + (0.1 * INCH);

    // Main image (4.5 inch height)
    const mainImageHeight = 4.5 * INCH;
    const mainImageWidth = contentWidth;
    
    const coverImages = images.cover || [];
    const galleryImages = images.property || [];
    const allImages = [...coverImages, ...galleryImages];
    
    // Find main image (prefer exterior_front, otherwise first image)
    let mainImagePath = null;
    for (const imgPath of allImages) {
        if (imgPath) {
            const filename = path.basename(imgPath).toLowerCase();
            if (filename.includes('exterior') && filename.includes('front')) {
                mainImagePath = imgPath;
                break;
            }
        }
    }
    if (!mainImagePath && allImages.length > 0) {
        mainImagePath = allImages[0];
    }

    addImageToPDF(doc, mainImagePath, MARGIN, currentY, mainImageWidth, mainImageHeight, 'cover');
    currentY += mainImageHeight + (0.2 * INCH);

    // Three thumbnail images (1.5 inch height each)
    const thumbnailHeight = 1.5 * INCH;
    const thumbnailWidth = (contentWidth - (0.4 * INCH)) / 3;
    const thumbnailSpacing = 0.2 * INCH;

    // Get thumbnails (exclude main image)
    const thumbnails = allImages.filter(img => img !== mainImagePath).slice(0, 3);
    
    for (let i = 0; i < 3; i++) {
        const thumbX = MARGIN + i * (thumbnailWidth + thumbnailSpacing);
        const thumbPath = thumbnails[i] || null;
        // Use 'cover' to enforce identical thumbnail sizes
        addImageToPDF(doc, thumbPath, thumbX, currentY, thumbnailWidth, thumbnailHeight, 'cover');
    }
    currentY += thumbnailHeight + (0.2 * INCH);

    // Footer bar (gold background, 0.4 inch height)
    const footerHeight = 0.4 * INCH;
    const footerY = pageHeight - MARGIN - footerHeight;
    
    doc.rect(MARGIN, footerY, contentWidth, footerHeight)
       .fillColor(ACCENT_GOLD)
       .fill();

    const reportDate = formatDateWithOrdinal(new Date());
    const footerText = `Report created on ${reportDate}`;
    doc.fontSize(11)
       .font('Helvetica')
       .fillColor('#FFFFFF')
       .text(footerText, MARGIN, footerY + (footerHeight / 2) - 6, {
           width: contentWidth,
           align: 'center'
       });
}

// Helper function to calculate values for a single calculator
function calculateCalculatorValues(calculatorType, calcData, allData) {
    console.log(`calculateCalculatorValues - calculatorType: ${calculatorType}`);
    console.log(`calculateCalculatorValues - calcData keys:`, Object.keys(calcData));
    console.log(`calculateCalculatorValues - calculateInvestment available:`, !!calculateInvestment);
    
    let purchasePrice, depositAmount, mortgageAmount, totalInvestment, annualRent, monthlyRent, rentalYield, totalAnnualExpenses, annualProfit, monthlyProfit, roi;
    let depositPercent, mortgageRate, stampDuty, surveyCost, legalFees, loanSetup, annualMortgageInterest, councilTax, repairs, utilities, water, broadband, insurance;
    
    // Try to use calculator logic
    if (calculateInvestment) {
        const dataForCalc = { ...allData, ...calcData };
        dataForCalc.calculator_type = calculatorType;
        
        console.log(`calculateCalculatorValues - calling calculateInvestment with type: ${calculatorType}`);
        
        try {
            const calcResults = calculateInvestment(dataForCalc);
            console.log(`calculateCalculatorValues - calcResults:`, {
                purchasePrice: calcResults.purchasePrice,
                monthlyRent: calcResults.monthlyRent,
                annualProfit: calcResults.annualProfit,
                roi: calcResults.roi
            });
            
            // Extract results (works for all calculator types)
            purchasePrice = calcResults.purchasePrice || 0;
            depositAmount = calcResults.depositAmount || 0;
            mortgageAmount = calcResults.mortgageAmount || 0;
            totalInvestment = calcResults.totalInvestment || calcResults.netInvestment || 0;
            annualRent = calcResults.annualRent || calcResults.annualIncome || 0;
            monthlyRent = calcResults.monthlyRent || calcResults.monthlyIncome || 0;
            rentalYield = calcResults.rentalYield || 0;
            totalAnnualExpenses = calcResults.totalAnnualExpenses || 0;
            annualProfit = calcResults.annualProfit || calcResults.netProfit || 0;
            monthlyProfit = calcResults.monthlyProfit || 0;
            roi = calcResults.roi || 0;
            
            // Extract additional values for display
            depositPercent = parseFloat(calcData.deposit_percent || allData.deposit_percent) || 20;
            mortgageRate = parseFloat(calcData.mortgage_rate || allData.mortgage_rate) || 5.8;
            stampDuty = calcResults.stampDuty || parseCurrency(calcData.stamp_duty || allData.stamp_duty);
            surveyCost = calcResults.surveyCost || parseCurrency(calcData.survey_cost || allData.survey_cost);
            legalFees = calcResults.legalFees || parseCurrency(calcData.legal_fees || allData.legal_fees);
            loanSetup = calcResults.loanSetup || parseCurrency(calcData.loan_setup || allData.loan_setup);
            annualMortgageInterest = calcResults.annualMortgageInterest || 0;
            councilTax = calcResults.councilTax || parseCurrency(calcData.council_tax || allData.council_tax);
            repairs = calcResults.repairs || parseCurrency(calcData.repairs_maintenance || allData.repairs_maintenance);
            utilities = calcResults.utilities || parseCurrency(calcData.utilities || allData.utilities);
            water = calcResults.water || parseCurrency(calcData.water || allData.water);
            broadband = calcResults.broadband || parseCurrency(calcData.broadband_tv || allData.broadband_tv);
            insurance = calcResults.insurance || parseCurrency(calcData.insurance || allData.insurance);
            
            return { purchasePrice, depositAmount, mortgageAmount, totalInvestment, annualRent, monthlyRent, rentalYield, 
                     totalAnnualExpenses, annualProfit, monthlyProfit, roi, depositPercent, mortgageRate, 
                     stampDuty, surveyCost, legalFees, loanSetup, annualMortgageInterest, councilTax, repairs, 
                     utilities, water, broadband, insurance };
        } catch (err) {
            console.warn('Error using calculator logic, falling back:', err);
            console.error('Error details:', err.stack);
        }
    } else {
        console.warn('calculateInvestment function not available, using fallback calculation');
    }
    
    // Fallback to hardcoded calculation if calculator logic failed or not available
    console.log('Using fallback calculation for', calculatorType);
    purchasePrice = parseCurrency(calcData.purchase_price || allData.purchase_price);
    depositPercent = parseFloat(calcData.deposit_percent || allData.deposit_percent) || 20;
    monthlyRent = parseCurrency(calcData.monthly_rent || allData.monthly_rent);
    mortgageRate = parseFloat(calcData.mortgage_rate || allData.mortgage_rate) || 5.8;

    depositAmount = purchasePrice * (depositPercent / 100);
    annualRent = monthlyRent * 12;
    rentalYield = purchasePrice > 0 ? (annualRent / purchasePrice) * 100 : 0;

    stampDuty = parseCurrency(calcData.stamp_duty || allData.stamp_duty);
    surveyCost = parseCurrency(calcData.survey_cost || allData.survey_cost);
    legalFees = parseCurrency(calcData.legal_fees || allData.legal_fees);
    loanSetup = parseCurrency(calcData.loan_setup || allData.loan_setup);

    const totalPurchaseCosts = stampDuty + surveyCost + legalFees + loanSetup;
    totalInvestment = depositAmount + totalPurchaseCosts;

    mortgageAmount = purchasePrice - depositAmount;
    annualMortgageInterest = mortgageAmount * (mortgageRate / 100);

    councilTax = parseCurrency(calcData.council_tax || allData.council_tax);
    repairs = parseCurrency(calcData.repairs_maintenance || allData.repairs_maintenance);
    utilities = parseCurrency(calcData.utilities || allData.utilities);
    water = parseCurrency(calcData.water || allData.water);
    broadband = parseCurrency(calcData.broadband_tv || allData.broadband_tv);
    insurance = parseCurrency(calcData.insurance || allData.insurance);

    totalAnnualExpenses = annualMortgageInterest + councilTax + repairs + utilities + water + broadband + insurance;
    annualProfit = annualRent - totalAnnualExpenses;
    monthlyProfit = annualProfit / 12;
    roi = totalInvestment > 0 ? (annualProfit / totalInvestment) * 100 : 0;
    
    return { purchasePrice, depositAmount, mortgageAmount, totalInvestment, annualRent, monthlyRent, rentalYield, 
             totalAnnualExpenses, annualProfit, monthlyProfit, roi, depositPercent, mortgageRate, 
             stampDuty, surveyCost, legalFees, loanSetup, annualMortgageInterest, councilTax, repairs, 
             utilities, water, broadband, insurance };
}

// Helper function to render a single calculator section
function renderCalculatorSection(doc, calculatorType, values, startY) {
    // Map calculator type to display name
    const calculatorDisplayNames = {
        'standard-btl': 'Standard Buy to Let',
        'brr': 'Buy Refurbish Refinance',
        'flip': 'Flip',
        'holiday-let': 'Holiday Let',
        'rent-to-hmo': 'Rent to HMO',
        'rent-to-serviced': 'Rent to Serviced Accommodation',
        'purchase': 'Purchase Calculator'
    };
    const calculatorDisplayName = calculatorDisplayNames[calculatorType] || 'Standard Buy to Let';
    
    let currentY = startY;
    
    // Section title with calculator type
    doc.fontSize(24)
       .font('Helvetica-Bold')
       .fillColor('#000000')
       .text('Investment Opportunity', MARGIN, currentY);
    // Add spacing for the 24pt font (approximately 28-30 points height)
    currentY += 32;
    
    // Calculator type subtitle
    doc.fontSize(16)
       .font('Helvetica')
       .fillColor('#666666')
       .text(`${calculatorDisplayName} Calculator`, MARGIN, currentY);
    // Add spacing for the 16pt font (approximately 18-20 points height)
    currentY += 22;

    // Three key metrics boxes (gold background, horizontal) - fit exactly within content width
    const contentWidth = A4_WIDTH - (2 * MARGIN);
    const boxSpacing = 0.15 * INCH;
    const boxWidth = (contentWidth - (2 * boxSpacing)) / 3;
    const boxHeight = 1.3 * INCH; // Slightly taller

    const metricsY = currentY;
    
    // Center the three boxes within the printable content width so left/right paddings match
    const rowWidth = (3 * boxWidth) + (2 * boxSpacing);
    const rowStartX = MARGIN; // boxes span exactly the content width

    // Box 1: Purchase Price
    doc.rect(rowStartX, metricsY, boxWidth, boxHeight)
       .fillColor(ACCENT_GOLD)
       .fill();
    doc.fontSize(12)
       .font('Helvetica')
       .fillColor('#000000')
       .text('Purchase Price', rowStartX + 12, metricsY + 15, { width: boxWidth - 24, align: 'center' });
    doc.fontSize(24)
       .font('Helvetica-Bold')
       .fillColor('#FFFFFF')
        .text(formatCurrency(values.purchasePrice), rowStartX + 12, metricsY + 40, { width: boxWidth - 24, align: 'center' });

    // Box 2: Monthly Rent
    const box2X = rowStartX + boxWidth + boxSpacing;
    doc.rect(box2X, metricsY, boxWidth, boxHeight)
       .fillColor(ACCENT_GOLD)
       .fill();
    doc.fontSize(12)
       .font('Helvetica')
       .fillColor('#000000')
       .text('Estimated Monthly Rent', box2X + 12, metricsY + 15, { width: boxWidth - 24, align: 'center' });
    doc.fontSize(24)
       .font('Helvetica-Bold')
       .fillColor('#FFFFFF')
       .text(formatCurrency(values.monthlyRent) + 'pcm', box2X + 12, metricsY + 40, { width: boxWidth - 24, align: 'center' });

    // Box 3: Rental Yield
    const box3X = box2X + boxWidth + boxSpacing;
    doc.rect(box3X, metricsY, boxWidth, boxHeight)
       .fillColor(ACCENT_GOLD)
       .fill();
    doc.fontSize(12)
       .font('Helvetica')
       .fillColor('#000000')
       .text('Rental Yield', box3X + 12, metricsY + 15, { width: boxWidth - 24, align: 'center' });
    doc.fontSize(24)
       .font('Helvetica-Bold')
       .fillColor('#FFFFFF')
       .text(`${values.rentalYield.toFixed(1)}%`, box3X + 12, metricsY + 40, { width: boxWidth - 24, align: 'center' });

    currentY += boxHeight + 25;

    // Two column layout for costs and expenses
    // Match Python: colWidths=[2.2*inch, 1.3*inch, 2.2*inch, 1.3*inch]
    const leftLabelWidth = 2.2 * INCH;
    const leftValueWidth = 1.3 * INCH;
    const leftColWidth = leftLabelWidth + leftValueWidth;
    const rightLabelWidth = 2.2 * INCH;
    const rightValueWidth = 1.3 * INCH;
    const rightColWidth = rightLabelWidth + rightValueWidth;
    
    // Calculate positions - ensure no overlap
    // Total width needed: 2.2 + 1.3 + gap + 2.2 + 1.3 = 7.0 + gap
    // Available width: A4_WIDTH - 2*MARGIN = 595.28 - 108 = 487.28 points = 6.77 inches
    // So we need to reduce spacing or use available space efficiently
    const availableWidth = A4_WIDTH - (2 * MARGIN);
    const totalColumnsWidth = leftColWidth + rightColWidth;
    const columnSpacing = Math.max(0.2 * INCH, (availableWidth - totalColumnsWidth) / 2);
    
    const leftColX = MARGIN;
    const rightColX = leftColX + leftColWidth + columnSpacing;
    const tableStartY = currentY;

    // Left column: Purchase Costs
    doc.fontSize(12)
       .font('Helvetica-Bold')
       .fillColor('#000000')
       .text('Total Purchase Costs', leftColX, currentY);
    currentY += 20;

    const costs = [
        [`Deposit (${values.depositPercent}%)`, formatCurrency(values.depositAmount)],
        ['Stamp Duty', formatCurrency(values.stampDuty)],
        ['Survey', formatCurrency(values.surveyCost)],
        ['Legal Fees', formatCurrency(values.legalFees)],
        ['Loan Set-up', formatCurrency(values.loanSetup)],
        ['Total Investment Required', formatCurrency(values.totalInvestment)]
    ];

    let leftTableY = currentY;
    costs.forEach(([label, value], index) => {
        const isTotal = index === costs.length - 1;
        doc.fontSize(11)
           .font(isTotal ? 'Helvetica-Bold' : 'Helvetica')
           .fillColor('#000000');
        
        // Label (left aligned) - strictly constrain to column width
        const labelX = leftColX + 5;
        const labelMaxWidth = leftLabelWidth - 10; // Leave 5pt padding on each side
        doc.text(label, labelX, leftTableY, { 
            width: labelMaxWidth
        });
        
        // Value (right aligned) - position at end of left column
        const valueX = leftColX + leftLabelWidth;
        doc.text(value, valueX, leftTableY, { 
            width: leftValueWidth - 10, 
            align: 'right' 
        });
        
        // Draw horizontal line below (except for last row) - only within left column
        if (index < costs.length - 1) {
            doc.moveTo(leftColX, leftTableY + 12)
               .lineTo(leftColX + leftColWidth, leftTableY + 12)
               .strokeColor('#E0E0E0')
               .lineWidth(0.5)
               .stroke();
        }
        
        leftTableY += 18;
    });

    // Right column: Annual Expenses
    currentY = tableStartY;
    doc.fontSize(12)
       .font('Helvetica-Bold')
       .fillColor('#000000')
       .text('Total Annual Expenses', rightColX + 15, currentY); // 15pt left padding as per Python
    currentY += 20;

    const expenses = [
        [`Mortgage @ ${values.mortgageRate}% (Interest Only)`, formatCurrency(values.annualMortgageInterest)],
        ['Council Tax', formatCurrency(values.councilTax)],
        ['Repairs / Maintenance', formatCurrency(values.repairs)],
        ['Electric / Gas', formatCurrency(values.utilities)],
        ['Water', formatCurrency(values.water)],
        ['Broadband / TV', formatCurrency(values.broadband)],
        ['Insurance', formatCurrency(values.insurance)],
        ['Total', formatCurrency(values.totalAnnualExpenses)]
    ];

    let rightTableY = currentY;
    expenses.forEach(([label, value], index) => {
        const isTotal = index === expenses.length - 1;
        doc.fontSize(11)
           .font(isTotal ? 'Helvetica-Bold' : 'Helvetica')
           .fillColor('#000000');
        
        // Label (left aligned) - strictly constrain to column width with 15pt left padding
        const labelX = rightColX + 15; // 15pt padding as per Python
        // For the mortgage row, use smaller font to fit on one line
        if (index === 0 && label.includes('Mortgage')) {
            doc.fontSize(10);
            const labelMaxWidth = rightLabelWidth - 5; // More width for smaller font
            doc.text(label, labelX, rightTableY, { width: labelMaxWidth });
            doc.fontSize(11); // Reset font size
        } else {
            const labelMaxWidth = rightLabelWidth - 10;
            doc.text(label, labelX, rightTableY, { width: labelMaxWidth });
        }
        
        // Value (right aligned) - position at end of right column
        const valueX = rightColX + rightLabelWidth;
        doc.text(value, valueX, rightTableY, { 
            width: rightValueWidth - 10, 
            align: 'right' 
        });
        
        // Draw horizontal line below (except for last row) - only within right column
        if (index < expenses.length - 1) {
            doc.moveTo(rightColX + 15, rightTableY + 12)
               .lineTo(rightColX + rightColWidth, rightTableY + 12)
               .strokeColor('#E0E0E0')
               .lineWidth(0.5)
               .stroke();
        }
        
        rightTableY += 18;
    });

    // Profit/ROI boxes centered - single box split into label (left) and value (right)
    // Position: centered on page, below the tables
    const profitLabelWidth = 2 * INCH;
    const profitValueWidth = 2 * INCH;
    const profitTotalWidth = profitLabelWidth + profitValueWidth; // 4 inches total
    const profitBoxHeight = 0.9 * INCH;
    
    // Calculate Y position - below the tables with some spacing
    const maxTableHeight = Math.max(leftTableY, rightTableY);
    // Center the boxes horizontally
    const profitX = (A4_WIDTH - profitTotalWidth) / 2; // Center horizontally
    let profitY = maxTableHeight + (0.5 * INCH);

    const profitData = [
        ['Monthly Profit', formatCurrency(values.monthlyProfit)],
        ['Annual Profit', formatCurrency(values.annualProfit)],
        ['ROI', `${values.roi.toFixed(1)}%`]
    ];

    profitData.forEach(([label, value]) => {
        // Draw single gold box (both label and value areas)
        doc.rect(profitX, profitY, profitTotalWidth, profitBoxHeight)
           .fillColor(ACCENT_GOLD)
           .fill();
        
        // Label on left side (black text)
        doc.fontSize(13)
           .font('Helvetica')
           .fillColor('#000000')
           .text(label, profitX + 18, profitY + (profitBoxHeight / 2) - 6, {
               width: profitLabelWidth - 36,
               align: 'left'
           });
        
        // Value on right side (white bold text)
        doc.fontSize(24)
           .font('Helvetica-Bold')
           .fillColor('#FFFFFF')
           .text(value, profitX + profitLabelWidth + 12, profitY + (profitBoxHeight / 2) - 12, { 
               width: profitValueWidth - 24, 
               align: 'right' 
           });
        
        profitY += profitBoxHeight + 10;
    });
    
    return profitY; // Return the final Y position
}

// Create investment opportunity page
function createInvestmentPage(doc, data, logoPath) {
    // Get selected calculators
    let selectedCalculators = data.selected_calculators || (data.calculator_type ? [data.calculator_type] : ['standard-btl']);
    
    // Debug logging
    console.log('createInvestmentPage - data.selected_calculators:', data.selected_calculators);
    console.log('createInvestmentPage - data.calculator_type:', data.calculator_type);
    console.log('createInvestmentPage - initial selectedCalculators:', selectedCalculators);
    
    // If no calculators selected, default to standard-btl
    if (!selectedCalculators || selectedCalculators.length === 0) {
        selectedCalculators = ['standard-btl'];
    }
    
    // Ensure selectedCalculators is an array
    if (typeof selectedCalculators === 'string') {
        selectedCalculators = selectedCalculators.split(',').map(s => s.trim()).filter(s => s);
    }
    
    // Ensure it's actually an array
    if (!Array.isArray(selectedCalculators)) {
        console.warn('selectedCalculators is not an array, converting:', selectedCalculators);
        selectedCalculators = [selectedCalculators];
    }
    
    console.log('createInvestmentPage - final selectedCalculators:', selectedCalculators);
    console.log('createInvestmentPage - number of calculators:', selectedCalculators.length);
    
    // Loop through all selected calculators and create a section for each
    selectedCalculators.forEach((calculatorType, index) => {
        console.log(`Creating calculator section ${index + 1}/${selectedCalculators.length}: ${calculatorType}`);
        
        // Draw header for each calculator section
        if (index > 0) {
            // Add a new page for each additional calculator
            console.log(`Adding new page for calculator ${index + 1}`);
            doc.addPage();
        }
        drawHeader(doc, logoPath);

        // Calculate content Y position (below header)
        const estimatedLogoHeight = (1.4 * INCH) * 0.4;
        const contentY = HEADER_TOP_OFFSET + estimatedLogoHeight + 12 + 9 + 20;
        
        // Get calculator-specific data
        const calcData = data[`calculator_${calculatorType}`] || {};
        console.log(`Calculator data for ${calculatorType}:`, Object.keys(calcData).length, 'fields');
        
        // Calculate values for this calculator
        const values = calculateCalculatorValues(calculatorType, calcData, data);
        
        // Render the calculator section
        renderCalculatorSection(doc, calculatorType, values, contentY);
    });
}

// Create Key Information page
function createKeyInformationPage(doc, data, images, logoPath) {
    drawHeader(doc, logoPath);
    
    // Calculate content Y position (below header)
    const estimatedLogoHeight = (1.4 * INCH) * 0.4;
    let currentY = HEADER_TOP_OFFSET + estimatedLogoHeight + 12 + 9 + 20;
    
    doc.fontSize(24)
       .font('Helvetica-Bold')
       .fillColor('#000000')
       .text('Key Information', MARGIN, currentY);
    currentY += 30;

    // Property image (6.5 x 3.5 inch)
    const coverImages = images.cover || [];
    const galleryImages = images.property || [];
    const candidateImages = [...coverImages, ...galleryImages];
    
    let mainImgPath = null;
    for (const imgPath of candidateImages) {
        if (imgPath) {
            const filename = path.basename(imgPath).toLowerCase();
            if (filename.includes('exterior') && filename.includes('front')) {
                mainImgPath = imgPath;
                break;
            }
        }
    }
    if (!mainImgPath && candidateImages.length > 0) {
        mainImgPath = candidateImages[0];
    }
    
    addImageToPDF(doc, mainImgPath, MARGIN, currentY, 6.5 * INCH, 3.5 * INCH, 'contain');
    currentY += (3.5 * INCH) + 20;

    // Property Metrics Table (4 columns: Asking price, Bedrooms, Size, On market)
    const metricsTableY = currentY;
    const metricColWidth = 1.8 * INCH;
    const metrics = [
        ['Asking price', 'Bedrooms', 'Size', 'On the market for'],
        [
            data.asking_price || 'N/A',
            data.bedrooms || 'N/A',
            `${data.size_sqm || 'N/A'} sqm`,
            `${data.days_on_market || 'N/A'} days`
        ]
    ];

    // Draw metrics table
    for (let col = 0; col < 4; col++) {
        const colX = MARGIN + (col * metricColWidth);
        
        // Label row
        doc.fontSize(11)
           .font('Helvetica')
           .fillColor('#000000')
           .text(metrics[0][col], colX + 5, metricsTableY + 8, { width: metricColWidth - 10, align: 'center' });
        
        // Value row
        doc.fontSize(16)
           .font('Helvetica-Bold')
           .fillColor('#000000')
           .text(metrics[1][col], colX + 5, metricsTableY + 28, { width: metricColWidth - 10, align: 'center' });
    }
    
    currentY += 60;

    // Key Features (bulleted list)
    if (data.key_features) {
        doc.fontSize(14)
           .font('Helvetica-Bold')
           .fillColor('#000000')
           .text('Key Features', MARGIN, currentY);
        currentY += 20;

        const features = data.key_features.split('\n').filter(f => f.trim());
        doc.fontSize(11)
           .font('Helvetica')
           .fillColor('#000000');
        
        features.forEach(feature => {
            doc.text(`• ${feature.trim()}`, MARGIN + 20, currentY, {
                width: A4_WIDTH - (2 * MARGIN) - 20,
                align: 'left'
            });
            currentY += doc.heightOfString(`• ${feature.trim()}`, { width: A4_WIDTH - (2 * MARGIN) - 20 }) + 5;
        });
    }
}

// Create EPC chart (stepped horizontal bands with right-side markers)
function createEPCChart(doc, data, x, y) {
    const chartWidth = 4.6 * INCH;
    const barHeight = 0.18 * INCH;
    const barSpacing = 0.09 * INCH;
    // Slightly shift the bars left so the grade letters have breathing room
    const startX = x - (0.15 * INCH);
    const startY = y;

    const epcBands = [
        { grade: 'A', min: 92, max: 100, color: '#008450', range: '92+' },
        { grade: 'B', min: 81, max: 91,  color: '#22c55e', range: '81-91' },
        { grade: 'C', min: 69, max: 80,  color: '#84cc16', range: '69-80' },
        { grade: 'D', min: 55, max: 68,  color: '#eab308', range: '55-68' },
        { grade: 'E', min: 39, max: 54,  color: '#f59e0b', range: '39-54' },
        { grade: 'F', min: 21, max: 38,  color: '#ef4444', range: '21-38' },
        { grade: 'G', min: 1,  max: 20,  color: '#dc2626', range: '1-20' }
    ];

    // Column headers (left: Score, middle: Energy rating, right: Current/Potential)
    doc.fontSize(13).font('Helvetica-Bold').fillColor('#000000');
    const leftColX = startX - 1.05 * INCH;
    const leftColW = 0.95 * INCH;
    // Column positions for Current and Potential (like reference image)
    const colGap = 0.20 * INCH;
    const currentColX = startX + chartWidth + colGap;
    const colWidth = 0.9 * INCH;
    const potentialColX = currentColX + colWidth + colGap;
    const currentHeaderX = currentColX;
    const potentialHeaderX = potentialColX;
    doc.text('Score', leftColX, startY - 16, { width: leftColW, align: 'left' });
    doc.text('Energy rating', startX, startY - 16, { width: chartWidth * 0.5, align: 'left' });
    // Column headers
    doc.fontSize(13).font('Helvetica-Bold').fillColor('#000000');
    doc.text('Current', currentHeaderX, startY - 16, { width: colWidth, align: 'center' });
    doc.text('Potential', potentialHeaderX, startY - 16, { width: colWidth, align: 'center' });
    // Vertical separators
    const topLineY = startY - 6;
    const bottomLineY = startY + epcBands.length * (barHeight + barSpacing) - barSpacing + 6;
    doc.save().lineWidth(1).strokeColor('#cbd5e1');
    doc.moveTo(currentColX - (colGap / 2), topLineY).lineTo(currentColX - (colGap / 2), bottomLineY).stroke();
    doc.moveTo(potentialColX - (colGap / 2), topLineY).lineTo(potentialColX - (colGap / 2), bottomLineY).stroke();
    doc.restore();

    // Draw stepped bars (A shortest, G longest)
    const maxSteps = epcBands.length; // 7
    epcBands.forEach((band, i) => {
        const step = maxSteps - i; // A=7? we want A shortest -> use i+1
    });
    // Recompute with A shortest
    epcBands.forEach((band, i) => {
        const yPos = startY + i * (barHeight + barSpacing);
        const proportion = (i + 1) / maxSteps; // A small -> 1/7, G large -> 7/7
        let width = chartWidth * proportion;
        // Ensure there is space for the grade letter to sit to the right of the bar
        const letterPadding = 0.08 * INCH;
        const maxLetterX = currentColX - 0.32 * INCH;
        const maxBarWidthForLetter = Math.max(0, (maxLetterX - startX) - letterPadding);
        if (width > maxBarWidthForLetter) {
            width = maxBarWidthForLetter;
        }
        // Score range on the left
        doc.fontSize(11).font('Helvetica-Bold').fillColor('#1f2937');
        doc.text(band.range, leftColX, yPos + barHeight / 2 - 6, { width: leftColW, align: 'left' });
        // Bar
        doc.rect(startX, yPos, width, barHeight)
           .fillColor(band.color)
           .strokeColor('#000000')
           .lineWidth(0.5)
           .fillAndStroke();
        // Grade letter just outside the bar on the right, but keep within chart area (not into columns)
        const barEndX = startX + width;
        const letterX = Math.min(barEndX + letterPadding, maxLetterX);
        const letterY = yPos + barHeight / 2 - 7;
        doc.fontSize(14).font('Helvetica-Bold').fillColor('#111827');
        doc.text(band.grade, letterX, letterY);
    });

    // Current/Potential value markers as right-pointing small arrows aligned to the band centers
    let currentScore = parseInt(data.current_rating || data.epc_rating || '84');
    let potentialScore = parseInt(data.potential_rating || '72');
    if (isNaN(currentScore)) currentScore = 84;
    if (isNaN(potentialScore)) potentialScore = 72;

    function bandCenterY(score) {
        for (let i = 0; i < epcBands.length; i++) {
            const b = epcBands[i];
            if (score >= b.min && score <= b.max) {
                return startY + i * (barHeight + barSpacing) + barHeight / 2;
            }
        }
        return startY + (epcBands.length - 1) * (barHeight + barSpacing) + barHeight / 2;
    }

    function drawValueBadge(xLeft, yCenter, value, letter, color, maxWidth) {
        const badgeWidth = Math.min(0.8 * INCH, (maxWidth || (0.8 * INCH)));
        const badgeHeight = 0.24 * INCH;
        const x0 = xLeft, y0 = yCenter - badgeHeight / 2;
        const tip = 0.20 * INCH;
        // Badge body
        doc.save();
        doc.fillColor(color).strokeColor('#000000').lineWidth(0.5);
        doc.moveTo(x0, y0)
           .lineTo(x0 + badgeWidth, y0)
           .lineTo(x0 + badgeWidth + tip, y0 + badgeHeight / 2)
           .lineTo(x0 + badgeWidth, y0 + badgeHeight)
           .lineTo(x0, y0 + badgeHeight)
           .lineTo(x0, y0)
           .closePath()
           .fillAndStroke();
        doc.restore();
        // Text "value | letter"
        doc.fontSize(9).font('Helvetica-Bold').fillColor('#111827');
        doc.text(`${value} | ${letter}`, x0 + 5, y0 + 2, { width: badgeWidth - 7, align: 'left' });
    }

    // Determine letters for scores
    function letterFor(score) {
        for (const b of epcBands) {
            if (score >= b.min && score <= b.max) return b.grade;
        }
        return 'G';
    }

    // Place badges inside the fixed Current/Potential columns
    // Ensure badges fit inside their columns (account for tip)
    const badgeMaxWidth = colWidth - 0.14 * INCH;
    drawValueBadge(currentColX + 0.04 * INCH, bandCenterY(currentScore), currentScore, letterFor(currentScore), '#ffd86b', badgeMaxWidth); // yellow
    drawValueBadge(potentialColX + 0.04 * INCH, bandCenterY(potentialScore), potentialScore, letterFor(potentialScore), '#c9f29b', badgeMaxWidth); // light green

    // (Legend headers now provided by the Current/Potential columns)

    // Efficiency captions at bottom left/right
    const captionY = startY + epcBands.length * (barHeight + barSpacing) + 2;
    doc.fontSize(9).font('Helvetica').fillColor('#666666');
    doc.text('Very energy efficient - lower running costs', startX, captionY + 6, { width: chartWidth / 2, align: 'left' });
    doc.text('Not energy efficient - higher running costs', startX + chartWidth / 2, captionY + 6, { width: chartWidth / 2, align: 'right' });
}

// Create Other Key Information page (EPC, Broadband)
function createOtherKeyInformationPage(doc, data, images, logoPath) {
    drawHeader(doc, logoPath);
    
    // Calculate content Y position (below header)
    const estimatedLogoHeight = (1.4 * INCH) * 0.4;
    let currentY = HEADER_TOP_OFFSET + estimatedLogoHeight + 12 + 9 + 20;
    
    doc.fontSize(24)
       .font('Helvetica-Bold')
       .fillColor('#000000')
       .text('Other Key Information', MARGIN, currentY);
    currentY += 30;

    // Property image (7 x 3.3 inch)
    const coverImages = images.cover || [];
    const galleryImages = images.property || [];
    const candidateImages = [...coverImages, ...galleryImages];
    
    let mainImgPath = null;
    for (const imgPath of candidateImages) {
        if (imgPath) {
            const filename = path.basename(imgPath).toLowerCase();
            if (filename.includes('exterior') && filename.includes('front')) {
                mainImgPath = imgPath;
                break;
            }
        }
    }
    if (!mainImgPath && candidateImages.length > 0) {
        mainImgPath = candidateImages[0];
    }
    
    addImageToPDF(doc, mainImgPath, MARGIN, currentY, 7 * INCH, 3.3 * INCH, 'contain');
    currentY += (3.3 * INCH) + 20;

    // EPC Section - Title centered, chart wider and centered
    const epcStartY = currentY;
    
    // Title centered
    doc.fontSize(14)
       .font('Helvetica-Bold')
       .fillColor('#000000')
       .text('Energy Performance Certificate', MARGIN, epcStartY + 14, {
           width: A4_WIDTH - (2 * MARGIN),
           align: 'center'
       });
    
    currentY += 30;
    
    // Chart - horizontal, span across the page but slightly smaller
    const chartWidth = 5.8 * INCH;
    const chartX = (A4_WIDTH - chartWidth) / 2;
    createEPCChart(doc, data, chartX, currentY);
    
    // EPC Details below chart (if available)
    currentY += (2.4 * INCH) + 20;
    
    doc.fontSize(11)
       .font('Helvetica')
       .fillColor('#000000');
    
    if (data.inspection_date) {
        doc.font('Helvetica-Bold')
           .text('Latest available inspection date', MARGIN, currentY);
        doc.font('Helvetica')
           .text(data.inspection_date, MARGIN, currentY + 15);
        currentY += 35;
    }
    
    if (data.window_glazing) {
        doc.font('Helvetica-Bold')
           .text('Window glazing', MARGIN, currentY);
        doc.font('Helvetica')
           .text(data.window_glazing, MARGIN, currentY + 15);
        currentY += 35;
    }
    
    if (data.building_age) {
        doc.font('Helvetica-Bold')
           .text('Building construction age band', MARGIN, currentY);
        doc.font('Helvetica')
           .text(data.building_age, MARGIN, currentY + 15);
        currentY += 35;
    }

    // EPC Disclaimer (below EPC details)
    currentY += 10;
    doc.fontSize(9)
       .font('Helvetica')
       .fillColor('#666666')
       .text('This EPC data is accurate up to 6 months ago. If a more recent EPC assessment was done within this period, it will not be displayed here.', 
             MARGIN, currentY, {
                 width: A4_WIDTH - (2 * MARGIN),
                 align: 'left'
             });
    currentY += 40; // Increased spacing before broadband section

    // Broadband section - always show (matching Python structure)
    doc.fontSize(14)
       .font('Helvetica-Bold')
       .fillColor('#000000')
       .text('Internet / Broadband Availability', MARGIN, currentY);
    currentY += 20;

    // Three columns for broadband info (matching Python: 2.3, 2.3, 2.4 inches)
    const broadbandCol1Width = 2.3 * INCH;
    const broadbandCol2Width = 2.3 * INCH;
    const broadbandCol3Width = 2.4 * INCH;
    
    // Column 1: Broadband available
    doc.fontSize(11)
       .font('Helvetica')
       .fillColor('#000000')
       .text('Broadband available', MARGIN, currentY);
    doc.font('Helvetica-Bold')
       .text(data.broadband_available || 'N/A', MARGIN, currentY + 15);
    
    // Column 2: Download speed
    const col2X = MARGIN + broadbandCol1Width;
    doc.font('Helvetica')
       .text('Highest available download speed', col2X, currentY);
    doc.font('Helvetica-Bold')
       .text(data.download_speed || 'N/A', col2X, currentY + 15);
    
    // Column 3: Upload speed
    const col3X = col2X + broadbandCol2Width;
    doc.font('Helvetica')
       .text('Highest available upload speed', col3X, currentY);
    doc.font('Helvetica-Bold')
       .text(data.upload_speed || 'N/A', col3X, currentY + 15);
}

// Create City Map page
function createCityMapPage(doc, data, images, logoPath) {
    drawHeader(doc, logoPath);
    
    // Calculate content Y position (below header)
    const estimatedLogoHeight = (1.4 * INCH) * 0.4;
    let currentY = HEADER_TOP_OFFSET + estimatedLogoHeight + 12 + 9 + 20;
    
    doc.fontSize(24)
       .font('Helvetica-Bold')
       .fillColor('#000000')
       .text('City Map', MARGIN, currentY);
    currentY += 30;
 
    // Directions image - full width (match city images container width)
    const directionsImages = images.directions || [];
    let directionsPath = directionsImages[0] || null;
    
    // If not found, check sample_images folder (matching Python fallback)
    // Try multiple possible locations
    if (!directionsPath || !fs.existsSync(directionsPath)) {
        const possiblePaths = [
            path.join(__dirname, 'sample_images', 'directions.png'),
            path.join(__dirname, 'directions.png'),
            path.join(__dirname, 'sample_images', 'directions.jpg'),
            path.join(__dirname, 'directions.jpg'),
            path.join(process.cwd(), 'sample_images', 'directions.png'),
            path.join(process.cwd(), 'directions.png')
        ];
        for (const samplePath of possiblePaths) {
            if (fs.existsSync(samplePath)) {
                directionsPath = samplePath;
                break;
            }
        }
    }
    
    // Display city map image - match width of the 3 city images container below
    // City images container: 3 images at 2.3 inches each + spacing = A4_WIDTH - (2 * MARGIN)
    // Map is always 1280x768 pixels (aspect ratio 1.6667:1)
    const cityImagesContainerWidth = A4_WIDTH - (2 * MARGIN); // Same as city images container
    const directionsImgWidth = cityImagesContainerWidth; // Width in points
    const mapAspectRatio = 1280 / 768; // Known aspect ratio of composite map
    const directionsImgHeight = directionsImgWidth / mapAspectRatio; // Calculate height from aspect ratio
    
    // Use explicit width and height to ensure proper scaling
    // This ensures the map uses the full width and displays at the correct size
    try {
        if (directionsPath && fs.existsSync(directionsPath)) {
            doc.image(directionsPath, MARGIN, currentY, {
                width: directionsImgWidth,
                height: directionsImgHeight,
                align: 'center',
                valign: 'top'
            });
        } else {
            // Draw placeholder if image doesn't exist
            doc.rect(MARGIN, currentY, directionsImgWidth, directionsImgHeight)
               .fillColor('#CCCCCC')
               .fill();
        }
    } catch (error) {
        console.error('Error adding map image:', error);
        // Draw placeholder on error
        doc.rect(MARGIN, currentY, directionsImgWidth, directionsImgHeight)
           .fillColor('#CCCCCC')
           .fill();
    }
    
    // Advance by the actual image height to avoid overlap
    currentY += directionsImgHeight + 20;

    // About the City section (always render, like original)
    doc.fontSize(14)
       .font('Helvetica-Bold')
       .fillColor('#000000')
       .text('About the City', MARGIN, currentY);
    currentY += 20;

    // City name in blue/bold (matching Python highlight_style)
    const fallbackCity = 'Liverpool';
    const cityName = (data.city && String(data.city).trim().length > 0) ? data.city : fallbackCity;
    doc.fontSize(12)
       .font('Helvetica-Bold')
       .fillColor(PRIMARY_BLUE)
       .text(cityName, MARGIN, currentY);
    currentY += 20;

    // City description (fallback mock content when not provided)
    const aboutFallback = 'Liverpool is a port city and metropolitan borough in Merseyside, England. It is the administrative, cultural and economic centre of the Liverpool City Region with a population of over 1.5 million.';
    const aboutText = (data.about_city && String(data.about_city).trim().length > 0)
        ? data.about_city
        : aboutFallback;
    doc.fontSize(11)
       .font('Helvetica')
       .fillColor('#000000')
       .text(aboutText, MARGIN, currentY, {
           width: A4_WIDTH - (2 * MARGIN),
           align: 'left'
       });
    const aboutHeight = doc.heightOfString(aboutText, { width: A4_WIDTH - (2 * MARGIN) });
    currentY += aboutHeight + 15;

    // Population (matching Python format: "Population: {value}")
    doc.fontSize(11)
       .font('Helvetica-Bold')
       .fillColor('#000000')
       .text('Population: ', MARGIN, currentY);
    doc.font('Helvetica')
       .text((data.population && String(data.population).trim().length > 0) ? data.population : '508,986', MARGIN + 80, currentY);
    currentY += 20;

    // City Images (3 horizontal, 2.3 x 2.3 inch each) - always show 3 (placeholders if missing)
    const cityImages = images.city || [];
    const fixedSize = 2.3 * INCH;
    
    // Fallback to sample images if not provided (matching Python)
    let cityQueue = [...cityImages];
    if (cityQueue.length === 0) {
        const possiblePaths = [
            path.join(__dirname, 'sample_images', 'liverpool1.jpg'),
            path.join(__dirname, 'sample_images', 'liverpool2.jpg'),
            path.join(__dirname, 'sample_images', 'liverpool3.jpg'),
            path.join(process.cwd(), 'sample_images', 'liverpool1.jpg'),
            path.join(process.cwd(), 'sample_images', 'liverpool2.jpg'),
            path.join(process.cwd(), 'sample_images', 'liverpool3.jpg')
        ];
        // Try to find liverpool images
        for (let i = 1; i <= 3; i++) {
            for (const basePath of [__dirname, process.cwd()]) {
                const samplePath = path.join(basePath, 'sample_images', `liverpool${i}.jpg`);
                if (fs.existsSync(samplePath)) {
                    cityQueue.push(samplePath);
                    break;
                }
            }
        }
    }
    
    // Always show 3 images (placeholders if missing)
    while (cityQueue.length < 3) {
        cityQueue.push(null);
    }
    
    const cityImgSpacing = (A4_WIDTH - (2 * MARGIN) - (3 * fixedSize)) / 2;
    
    currentY += 20;
    for (let i = 0; i < 3; i++) {
        const imgX = MARGIN + i * (fixedSize + cityImgSpacing);
        const cityImgPath = cityQueue[i] || null;
        // Use 'cover' so all three appear identical in size and crop if needed
        addImageToPDF(doc, cityImgPath, imgX, currentY, fixedSize, fixedSize, 'cover');
    }
}

// Main PDF generation function
function generatePDF(data, images, outputPath, logoPath) {
    return new Promise((resolve, reject) => {
        try {
            // Validate output path with better error messages
            
            if (!outputPath) {
                reject(new Error('Invalid output path provided: path is null or undefined'));
                return;
            }
            
            if (typeof outputPath !== 'string') {
                reject(new Error(`Invalid output path provided: expected string, got ${typeof outputPath}`));
                return;
            }
            
            if (outputPath.trim() === '') {
                reject(new Error('Invalid output path provided: path is empty'));
                return;
            }

            // Normalize the path for cross-platform compatibility
            const normalizedOutputPath = path.normalize(outputPath);

            const doc = new PDFDocument({
                size: 'A4',
                margins: {
                    top: MARGIN,
                    bottom: MARGIN,
                    left: MARGIN,
                    right: MARGIN
                }
            });

            const stream = fs.createWriteStream(normalizedOutputPath);
            doc.pipe(stream);

            // Cover page
            createCoverPage(doc, data, images, logoPath);
            doc.addPage();

            // Investment Opportunity page
            createInvestmentPage(doc, data, logoPath);
            doc.addPage();

            // Key Information page
            createKeyInformationPage(doc, data, images, logoPath);
            doc.addPage();

            // Other Key Information page (EPC, Broadband)
            createOtherKeyInformationPage(doc, data, images, logoPath);
            doc.addPage();

            // Floor Plans
            const floorPlans = images.floor_plans || [];
            const estimatedLogoHeight = (1.4 * INCH) * 0.4;
            if (floorPlans.length > 0) {
                for (const floorPlanPath of floorPlans) {
                    drawHeader(doc, logoPath);
                    let currentY = HEADER_TOP_OFFSET + estimatedLogoHeight + 12 + 9 + 20;
                    
                    doc.fontSize(24)
                       .font('Helvetica-Bold')
                       .fillColor('#000000')
                       .text('Floor Plans', MARGIN, currentY);
                    currentY += 30;
                    
                    addImageToPDF(doc, floorPlanPath, MARGIN, currentY, 6.5 * INCH, 4.5 * INCH, 'contain');
                    
                    if (floorPlanPath !== floorPlans[floorPlans.length - 1]) {
                        doc.addPage();
                    }
                }
            } else {
                // Show placeholder if no floor plans
                drawHeader(doc, logoPath);
                let currentY = HEADER_TOP_OFFSET + estimatedLogoHeight + 12 + 9 + 20;
                
                doc.fontSize(24)
                   .font('Helvetica-Bold')
                   .fillColor('#000000')
                   .text('Floor Plans', MARGIN, currentY);
                currentY += 30;
                
                addImageToPDF(doc, null, MARGIN, currentY, 6.5 * INCH, 4.5 * INCH, 'contain');
            }
            doc.addPage();

            // Property gallery
            const galleryImages = images.property || images.cover || [];
            if (galleryImages.length > 0) {
                drawHeader(doc, logoPath);
                let currentY = HEADER_TOP_OFFSET + estimatedLogoHeight + 12 + 9 + 20;
                
                doc.fontSize(24)
                   .font('Helvetica-Bold')
                   .fillColor('#000000')
                   .text('Property Images', MARGIN, currentY);
                // Use actual title height plus extra padding to avoid overlap
                const titleHeight = doc.heightOfString('Property Images', { width: A4_WIDTH - (2 * MARGIN) });
                currentY += titleHeight + 28;

                // Use smaller fixed box to ensure clear space between images
                const galleryWidth = 5.4 * INCH;
                const galleryHeight = 2.8 * INCH;
                const gallerySpacing = 60; // 60pt (~0.83in) vertical gap
                const contentWidth = A4_WIDTH - (2 * MARGIN);
                const galleryX = MARGIN + (contentWidth - galleryWidth) / 2; // center horizontally

                for (const imgPath of galleryImages) {
                    if (currentY + galleryHeight + gallerySpacing > A4_HEIGHT - MARGIN) {
                        doc.addPage();
                        drawHeader(doc, logoPath);
                        currentY = HEADER_TOP_OFFSET + estimatedLogoHeight + 12 + 9 + 20;
                        // Re-draw the title on the new page and add same padding
                        doc.fontSize(24)
                           .font('Helvetica-Bold')
                           .fillColor('#000000')
                           .text('Property Images', MARGIN, currentY);
                        const contTitleHeight = doc.heightOfString('Property Images', { width: A4_WIDTH - (2 * MARGIN) });
                        currentY += contTitleHeight + 28;
                    }
                    
                    // Use 'cover' so every image fills the exact same width/height and is centered
                    addImageToPDF(doc, imgPath, galleryX, currentY, galleryWidth, galleryHeight, 'cover');
                    // Larger spacing between stacked images
                    currentY += galleryHeight + gallerySpacing;
                }
            }
            doc.addPage();

            // City Map page
            createCityMapPage(doc, data, images, logoPath);

            doc.end();

            stream.on('finish', () => {
                resolve();
            });

            stream.on('error', (error) => {
                reject(error);
            });
        } catch (error) {
            reject(error);
        }
    });
}

module.exports = { generatePDF };
