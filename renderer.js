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
    // Allow manual override via localStorage for testing (only in development)
    if (typeof Storage !== 'undefined' && localStorage.getItem('backend_url')) {
        const override = localStorage.getItem('backend_url');
        // Only allow localhost override if we're actually on localhost
        if (override.includes('localhost') && typeof window !== 'undefined' && window.location && 
            (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1')) {
            console.log('🔗 Using localStorage override (localhost):', override);
            return override;
        }
        // For production, ignore localhost overrides
        if (!override.includes('localhost')) {
            console.log('🔗 Using localStorage override (hosted):', override);
            return override;
        }
    }
    
    // Check if we're in Electron (local development)
    if (ipcRenderer) {
        // Electron mode - use localhost for local development
        console.log('🔗 Electron mode - using localhost:8080');
        return 'http://localhost:8080';
    }
    
    // Check if we're running on localhost (browser mode - development)
    if (typeof window !== 'undefined' && window.location && 
        (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1')) {
        console.log('🔗 Browser on localhost - using localhost:8080');
        return 'http://localhost:8080';
    }
    
    // Production: Always use hosted backend
    const hostedUrl = 'https://pdf-generator-backend-fbtb.onrender.com';
    console.warn('⚠️ Using HOSTED backend (not local):', hostedUrl);
    console.warn('⚠️ To use local backend, access the app via http://localhost:3000');
    console.warn('⚠️ Or set localStorage: localStorage.setItem("backend_url", "http://localhost:8080")');
    return hostedUrl;
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
    
    // Get selected calculators (multiple)
    const checkboxes = document.querySelectorAll('.calculator-checkbox input[type="checkbox"]:checked');
    const selectedCalculators = Array.from(checkboxes).map(cb => cb.dataset.calculator);
    data.selected_calculators = selectedCalculators;
    
    // If only one calculator selected, also set calculator_type for backward compatibility
    if (selectedCalculators.length === 1) {
        data.calculator_type = selectedCalculators[0];
    }
    
    // Explicitly set calculator_type to 'brr' if BRR is selected (even if multiple calculators)
    if (selectedCalculators.includes('brr')) {
        data.calculator_type = 'brr';
        data.is_brr_calculator = true;
    }
    
    // Explicitly set calculator_type to 'holiday-let' if Holiday Let is selected (even if multiple calculators)
    // This takes priority over BRR if both are selected
    if (selectedCalculators.includes('holiday-let')) {
        data.calculator_type = 'holiday-let';
        data.is_holiday_let_calculator = true;
    }
    
    // Explicitly set calculator_type to 'rent-to-hmo' if Rent to HMO is selected
    if (selectedCalculators.includes('rent-to-hmo')) {
        data.calculator_type = 'rent-to-hmo';
        data.is_rent_to_hmo_calculator = true;
    }
    
    // Explicitly set calculator_type to 'standard-btl' if Standard Buy to Let is selected
    if (selectedCalculators.includes('standard-btl')) {
        data.calculator_type = 'standard-btl';
        data.is_standard_btl_calculator = true;
    }
    
    // Get calculator-specific data
    selectedCalculators.forEach(calcType => {
        const calcData = {};
        const calcInputs = document.querySelectorAll(`[data-calculator="${calcType}"]`);
        calcInputs.forEach(input => {
            if (input.dataset.originalId) {
                // Use original field ID (without calculator prefix)
                calcData[input.dataset.originalId] = input.value.trim();
            }
        });
        
        // For BRR, Holiday Let, Rent to HMO, and Standard Buy to Let calculators, also include all calculated/displayed values and breakdowns
        if (calcType === 'brr' || calcType === 'holiday-let' || calcType === 'rent-to-hmo' || calcType === 'standard-btl') {
            // Get financing type (for BRR and Holiday Let only)
            let financingType = null;
            if (calcType === 'brr' || calcType === 'holiday-let') {
                const financingTypeHidden = document.getElementById(`${calcType}_financing_type_hidden`);
                financingType = financingTypeHidden?.value || (calcType === 'holiday-let' ? 'mortgage' : 'bridging');
                calcData.financing_type = financingType;
                calcData.chosen_financing = financingType === 'mortgage' ? 'Mortgage' : (financingType === 'bridging' ? 'Bridging Finance' : 'Cash');
            }
            
            // For Rent to HMO, get strategy instead
            if (calcType === 'rent-to-hmo') {
                const strategySelect = document.getElementById(`${calcType}_chosen_strategy`);
                const strategy = strategySelect?.value || 'hmo';
                calcData.chosen_strategy = strategy;
                calcData.strategy = strategy === 'holiday-let' ? 'Holiday Let' : 'HMO';
            }
            
            // Helper function to get value from element (handles both input and text content)
            const getValue = (id, isInput = false) => {
                const el = document.getElementById(id);
                if (!el) return '';
                if (isInput) {
                    return el.value?.trim() || el.textContent?.trim() || '';
                }
                return el.textContent?.trim() || el.value?.trim() || '';
            };
            
            // Get all displayed calculated values and breakdowns
            const calculatedFields = {
                // Purchase section (BRR/Holiday Let/Standard BTL) or Acquisition (Rent to HMO)
                'purchase_price': getValue(`${calcType}_purchase_price`, true),
                'purchase_price_display': getValue(`${calcType}_purchase_price`, true),
                'stamp_duty': getValue(`${calcType}_stamp_duty`),
                'survey_costs': getValue(`${calcType}_survey_costs`, true),
                'legal_fees': getValue(`${calcType}_legal_fees`, true),
                'total_investment': getValue(`${calcType}_total_investment`),
                'total_investment_required': getValue(`${calcType}_total_investment`),
                
                // Standard BTL specific - Financing
                ...(calcType === 'standard-btl' ? {
                    'financing_type': document.getElementById(`${calcType}_financing_type_hidden`)?.value || 'mortgage',
                    'mortgage_payments': getValue(`${calcType}_mortgage_payments`, true),
                    'refurb_enabled': document.getElementById(`${calcType}_refurb_enabled`)?.checked || false,
                    'refurb_cost': getValue(`${calcType}_refurb_cost`, true),
                } : {}),
                
                // Rent to HMO specific - Acquisition
                'deposit': getValue(`${calcType}_deposit`, true),
                'reference_fees': getValue(`${calcType}_reference_fees`, true),
                
                // Initial Financing - Mortgage
                'mortgage_setup_fee': getValue(`${calcType}_mortgage_setup_fee`, true),
                'mortgage_ltv': getValue(`${calcType}_mortgage_ltv`, true),
                'mortgage_interest_rate': getValue(`${calcType}_mortgage_interest_rate`, true),
                'mortgage_term_years': getValue(`${calcType}_mortgage_term_years`, true),
                'mortgage_type_interest_only': document.querySelector(`[data-mortgage-type="interest_only"][data-calculator="${calcType}"]`)?.classList.contains('pe-financing-type-active') || false,
                'mortgage_required': getValue(`${calcType}_mortgage_required`),
                'mortgage_payments': getValue(`${calcType}_mortgage_payments`),
                
                // Initial Financing - Bridging
                'bridging_setup_fee': getValue(`${calcType}_bridging_setup_fee`, true),
                'bridging_ltv': getValue(`${calcType}_bridging_ltv`, true),
                'bridging_interest_rate_monthly': getValue(`${calcType}_bridging_interest_rate_monthly`, true),
                'bridging_interest_display': getValue(`${calcType}_bridging_interest`, true),
                'bridging_finance_required': getValue(`${calcType}_bridging_finance_required`),
                
                // Refurb
                'refurb_enabled': document.getElementById(`${calcType}_refurb_enabled`)?.checked !== false, // Default to true if checkbox exists
                'refurb_cost': getValue(`${calcType}_refurb_cost`, true),
                'include_in_bridging': document.getElementById(`${calcType}_include_in_bridging`)?.classList.contains('pe-toggle-button-active') || false,
                'vacant_period': getValue(`${calcType}_vacant_period`, true),
                'refurb_council_tax': getValue(`${calcType}_refurb_council_tax`, true),
                'refurb_council_tax_monthly': getValue(`${calcType}_refurb_council_tax_monthly`, true),
                'refurb_electric_gas': getValue(`${calcType}_refurb_electric_gas`, true),
                'refurb_water': getValue(`${calcType}_refurb_water`, true),
                'refurb_insurance': getValue(`${calcType}_refurb_insurance`, true),
                
                // Refinance
                'estimated_market_value': getValue(`${calcType}_estimated_market_value`),
                'refinance_setup_fee': getValue(`${calcType}_refinance_setup_fee`, true),
                'refinance_ltv': getValue(`${calcType}_refinance_ltv`, true),
                'refinance_mortgage_type_interest_only': document.querySelector(`[data-refinance-mortgage-type="interest_only"][data-calculator="${calcType}"]`)?.classList.contains('pe-financing-type-active') || false,
                'refinance_interest_rate': getValue(`${calcType}_refinance_interest_rate`, true),
                'refinance_mortgage_term_years': getValue(`${calcType}_refinance_mortgage_term_years`, true),
                'refinance_mortgage_payments': getValue(`${calcType}_refinance_mortgage_payments`),
                'locked_in_equity': getValue(`${calcType}_locked_in_equity`),
                'money_left_in': getValue(`${calcType}_money_left_in`),
                'ideal_purchase_price': getValue(`${calcType}_ideal_purchase_price`),
                
                // Rental Income / Income
                'monthly_rent': getValue(`${calcType}_monthly_rent`, true),
                'nightly_rate': getValue(`${calcType}_nightly_rate`, true), // Holiday Let / Rent to HMO specific
                'occupancy_rate': getValue(`${calcType}_occupancy_rate`, true), // Holiday Let / Rent to HMO specific
                'gross_yield': getValue(`${calcType}_gross_yield`),
                'total_rental_income': getValue(`${calcType}_total_rental_income`), // Rent to HMO specific
                
                // Rent to HMO specific - Room rents (collect all room rents)
                ...(calcType === 'rent-to-hmo' ? (() => {
                    const roomRents = {};
                    for (let i = 1; i <= 10; i++) {
                        const roomRent = getValue(`${calcType}_room_${i}_rent`, true);
                        if (roomRent) {
                            roomRents[`room_${i}_rent`] = roomRent;
                        }
                    }
                    return roomRents;
                })() : {}),
                
                // Rent to HMO specific - Strategy and Rent to Owner
                'chosen_strategy': calcType === 'rent-to-hmo' ? getValue(`${calcType}_chosen_strategy`, true) : undefined,
                'monthly_rent_to_owner': calcType === 'rent-to-hmo' ? getValue(`${calcType}_monthly_rent_to_owner`, true) : undefined,
                
                // Ongoing Costs
                'maintenance_percent': getValue(`${calcType}_maintenance_percent`, true),
                'ongoing_insurance': getValue(`${calcType}_ongoing_insurance`, true),
                'agent_fees': getValue(`${calcType}_agent_fees`, true),
                'management_fee': getValue(`${calcType}_management_fee`, true), // Holiday Let specific
                'booking_fees': getValue(`${calcType}_booking_fees`, true), // Holiday Let specific
                'cleaning_costs': getValue(`${calcType}_cleaning_costs`, true), // Holiday Let specific
                'cleaning_fee': getValue(`${calcType}_cleaning_fee`, true), // Holiday Let specific (alternative field name)
                'tv_license': document.getElementById(`${calcType}_tv_license`)?.checked || false, // Holiday Let specific
                'communal_tv_license': document.getElementById(`${calcType}_communal_tv_license`)?.checked || false, // Rent to HMO specific
                'ongoing_mortgage_payments': getValue(`${calcType}_ongoing_mortgage_payments`),
                
                // Rent to HMO specific - Ongoing Costs
                'council_tax': getValue(`${calcType}_council_tax`, true), // Rent to HMO specific
                'utilities': getValue(`${calcType}_utilities`, true), // Rent to HMO specific
                'water': getValue(`${calcType}_water`, true), // Rent to HMO specific
                'broadband_tv': getValue(`${calcType}_broadband_tv`, true), // Rent to HMO specific
                'insurance': getValue(`${calcType}_insurance`, true), // Rent to HMO specific
                
                // Summary
                'total_annual_expenses': getValue(`${calcType}_total_annual_expenses`, true),
                'annual_profit': getValue(`${calcType}_annual_profit`),
                'monthly_profit': getValue(`${calcType}_monthly_profit`),
                'total_annual_profit': getValue(`${calcType}_annual_profit`),
                'total_monthly_profit': getValue(`${calcType}_monthly_profit`),
                
                // Metrics
                'roi': getValue(`${calcType}_roi`),
                'return_on_investment': getValue(`${calcType}_roi`),
                'roi_display': getValue(`${calcType}_roi_display`),
                'roce': getValue(`${calcType}_roce`),
                'return_on_capital_employed': getValue(`${calcType}_roce`),
                'net_yield': getValue(`${calcType}_net_yield`),
                'equity_10_years': getValue(`${calcType}_equity_10_years`),
                'appreciation': getValue(`${calcType}_appreciation`, true),
                'annual_property_appreciation': getValue(`${calcType}_appreciation`, true),
            };
            
            // For Standard BTL, add simplified calculations
            if (calcType === 'standard-btl') {
                const purchasePriceNum = parseFloat(getValue(`${calcType}_purchase_price`, true).replace(/[£,\s]/g, '')) || 0;
                const monthlyRentNum = parseFloat(getValue(`${calcType}_monthly_rent`, true).replace(/[£,\s]/g, '')) || 0;
                const annualRent = monthlyRentNum * 12;
                const mortgagePaymentsNum = parseFloat(getValue(`${calcType}_mortgage_payments`, true).replace(/[£,\s]/g, '')) || 0;
                const annualMortgagePayments = mortgagePaymentsNum * 12;
                const totalAnnualExpensesNum = annualMortgagePayments; // Simplified
                const annualProfitNum = annualRent - totalAnnualExpensesNum;
                const totalInvestmentNum = parseFloat(getValue(`${calcType}_total_investment`).replace(/[£,\s]/g, '')) || 0;
                const roiNum = totalInvestmentNum > 0 ? (annualProfitNum / totalInvestmentNum) * 100 : 0;
                const grossYieldNum = purchasePriceNum > 0 ? (annualRent / purchasePriceNum) * 100 : 0;
                const netYieldNum = purchasePriceNum > 0 ? (annualProfitNum / purchasePriceNum) * 100 : 0;
                
                calculatedFields.total_annual_income = annualRent;
                calculatedFields.gross_yield = `${grossYieldNum.toFixed(1)}%`;
                calculatedFields.net_yield = `${netYieldNum.toFixed(1)}%`;
                calculatedFields.roi = `${roiNum.toFixed(1)}%`;
                calculatedFields.return_on_investment = `${roiNum.toFixed(1)}%`;
            }
            
            // Calculate additional breakdown values for PDF
            // Only calculate financing-related values for BRR and Holiday Let
            if (calcType === 'brr' || calcType === 'holiday-let') {
                // Parse numeric values for calculations
                const purchasePriceNum = parseFloat(getValue(`${calcType}_purchase_price`, true).replace(/[£,\s]/g, '')) || 0;
                const mortgageRequiredStr = getValue(`${calcType}_mortgage_required`);
                const bridgingRequiredStr = getValue(`${calcType}_bridging_finance_required`);
                const initialMortgageNum = parseFloat((financingType === 'mortgage' ? mortgageRequiredStr : bridgingRequiredStr).replace(/[£,\s]/g, '')) || 0;
                const depositAmount = purchasePriceNum - initialMortgageNum;
                
                // Add calculated breakdown values
                calculatedFields.deposit_amount = depositAmount;
                calculatedFields.deposit_percent = purchasePriceNum > 0 ? ((depositAmount / purchasePriceNum) * 100).toFixed(1) : '0';
                calculatedFields.finance_required = financingType === 'mortgage' ? getValue(`${calcType}_mortgage_required`) : getValue(`${calcType}_bridging_finance_required`);
                calculatedFields.monthly_payments = financingType === 'mortgage' ? getValue(`${calcType}_mortgage_payments`) : getValue(`${calcType}_bridging_interest`, true);
                calculatedFields.outstanding_finance_balance = initialMortgageNum; // Initial mortgage/bridging amount
                calculatedFields.new_mortgage_amount = parseFloat(getValue(`${calcType}_estimated_market_value`).replace(/[£,\s]/g, '')) * (parseFloat(getValue(`${calcType}_refinance_ltv`, true).replace(/[%\s]/g, '')) || 75) / 100;
                calculatedFields.new_monthly_payments = getValue(`${calcType}_refinance_mortgage_payments`);
                calculatedFields.total_annual_income = (parseFloat(getValue(`${calcType}_monthly_rent`, true).replace(/[£,\s]/g, '')) || 0) * 12;
            } else if (calcType === 'rent-to-hmo') {
                // For Rent to HMO, calculate deposit from deposit field
                const depositNum = parseFloat(getValue(`${calcType}_deposit`, true).replace(/[£,\s]/g, '')) || 0;
                calculatedFields.deposit_amount = depositNum;
                
                // Calculate total annual income based on strategy
                const strategy = calcData.chosen_strategy || 'hmo';
                if (strategy === 'holiday-let') {
                    const nightlyRate = parseFloat(getValue(`${calcType}_nightly_rate`, true).replace(/[£,\s]/g, '')) || 0;
                    const occupancyRateValue = getValue(`${calcType}_occupancy_rate`, true) || '70 %';
                    const occupancyRate = parseFloat(occupancyRateValue.replace(/[%\s]/g, '')) || 70;
                    calculatedFields.total_annual_income = nightlyRate * 365 * (occupancyRate / 100);
                } else {
                    // HMO: sum of all room rents * 12
                    let totalMonthlyRent = 0;
                    for (let i = 1; i <= 10; i++) {
                        const roomRent = parseFloat(getValue(`${calcType}_room_${i}_rent`, true).replace(/[£,\s]/g, '')) || 0;
                        totalMonthlyRent += roomRent;
                    }
                    calculatedFields.total_annual_income = totalMonthlyRent * 12;
                }
            }
            
            // Expenses during refurb breakdown
            const vacantPeriodNum = parseFloat(getValue(`${calcType}_vacant_period`, true)) || 0;
            const councilTaxAnnual = parseFloat(getValue(`${calcType}_refurb_council_tax`, true).replace(/[£,\s]/g, '')) || 0;
            const councilTaxDuringRefurb = (councilTaxAnnual / 12) * vacantPeriodNum;
            const electricGasDuringRefurb = (parseFloat(getValue(`${calcType}_refurb_electric_gas`, true).replace(/[£,\s]/g, '')) || 0) * vacantPeriodNum;
            const waterDuringRefurb = (parseFloat(getValue(`${calcType}_refurb_water`, true).replace(/[£,\s]/g, '')) || 0) * vacantPeriodNum;
            const insuranceDuringRefurb = (parseFloat(getValue(`${calcType}_refurb_insurance`, true).replace(/[£,\s]/g, '')) || 0) * vacantPeriodNum;
            
            calculatedFields.council_tax_during_refurb = councilTaxDuringRefurb;
            calculatedFields.electric_gas_during_refurb = electricGasDuringRefurb;
            calculatedFields.water_during_refurb = waterDuringRefurb;
            calculatedFields.insurance_during_refurb = insuranceDuringRefurb;
            
            // Interest during refurb (mortgage payments or bridging interest) - only for BRR and Holiday Let
            if (calcType === 'brr' || calcType === 'holiday-let') {
                if (financingType === 'mortgage') {
                    const mortgagePayments = parseFloat(getValue(`${calcType}_mortgage_payments`).replace(/[£,\s]/g, '')) || 0;
                    calculatedFields.interest_during_refurb = mortgagePayments * vacantPeriodNum;
                } else if (financingType === 'bridging') {
                    const bridgingInterest = parseFloat(getValue(`${calcType}_bridging_interest`, true).replace(/[£,\s]/g, '')) || 0;
                    calculatedFields.interest_during_refurb = bridgingInterest * vacantPeriodNum;
                } else {
                    calculatedFields.interest_during_refurb = 0;
                }
                
                // Loan set-up fee (mortgage or bridging) - for "Loan Set-up" field in PDF
                let loanSetupFee = 0;
                const purchasePriceNum = parseFloat(getValue(`${calcType}_purchase_price`, true).replace(/[£,\s]/g, '')) || 0;
                const mortgageRequiredStr = getValue(`${calcType}_mortgage_required`);
                const bridgingRequiredStr = getValue(`${calcType}_bridging_finance_required`);
                const initialMortgageNum = parseFloat((financingType === 'mortgage' ? mortgageRequiredStr : bridgingRequiredStr).replace(/[£,\s]/g, '')) || 0;
                
                if (financingType === 'mortgage') {
                    const mortgageSetupFeeValue = getValue(`${calcType}_mortgage_setup_fee`, true);
                    if (mortgageSetupFeeValue.includes('%')) {
                        const percent = parseFloat(mortgageSetupFeeValue.replace(/[%\s]/g, '')) || 0;
                        loanSetupFee = initialMortgageNum * (percent / 100);
                    } else {
                        loanSetupFee = parseFloat(mortgageSetupFeeValue.replace(/[£,\s]/g, '')) || 0;
                    }
                } else if (financingType === 'bridging') {
                    const bridgingSetupFeeValue = getValue(`${calcType}_bridging_setup_fee`, true);
                    if (bridgingSetupFeeValue.includes('%')) {
                        const percent = parseFloat(bridgingSetupFeeValue.replace(/[%\s]/g, '')) || 0;
                        loanSetupFee = initialMortgageNum * (percent / 100);
                    } else {
                        loanSetupFee = parseFloat(bridgingSetupFeeValue.replace(/[£,\s]/g, '')) || 0;
                    }
                }
                calculatedFields.loan_setup = loanSetupFee;
                calculatedFields.loan_setup_fee = loanSetupFee;
            } else {
                // Rent to HMO doesn't have financing, so no interest or loan setup fees
                calculatedFields.interest_during_refurb = 0;
                calculatedFields.loan_setup = 0;
                calculatedFields.loan_setup_fee = 0;
            }
            
            // Add calculated fields to calcData
            Object.assign(calcData, calculatedFields);
            
            // Also add to root level for easier access by backend (with calculator prefix)
            const prefix = calcType === 'holiday-let' ? 'holiday_let_' : 'brr_';
            Object.keys(calculatedFields).forEach(key => {
                data[`${prefix}${key}`] = calculatedFields[key];
            });
            
            // Also add refurb_enabled to root level with prefix
            if (calcData.refurb_enabled !== undefined) {
                data[`${prefix}refurb_enabled`] = calcData.refurb_enabled;
            }
            
            // For Holiday Let, also add specific fields with prefix
            if (calcType === 'holiday-let') {
                if (calcData.nightly_rate !== undefined) {
                    data[`${prefix}nightly_rate`] = calcData.nightly_rate;
                }
                if (calcData.occupancy_rate !== undefined) {
                    data[`${prefix}occupancy_rate`] = calcData.occupancy_rate;
                }
                if (calcData.maintenance_percent !== undefined) {
                    data[`${prefix}maintenance_percent`] = calcData.maintenance_percent;
                }
                if (calcData.management_fee !== undefined) {
                    data[`${prefix}management_fee`] = calcData.management_fee;
                }
                if (calcData.booking_fees !== undefined) {
                    data[`${prefix}booking_fees`] = calcData.booking_fees;
                }
                if (calcData.cleaning_costs !== undefined) {
                    data[`${prefix}cleaning_costs`] = calcData.cleaning_costs;
                }
                if (calcData.cleaning_fee !== undefined) {
                    data[`${prefix}cleaning_fee`] = calcData.cleaning_fee;
                }
                if (calcData.tv_license !== undefined) {
                    data[`${prefix}tv_license`] = calcData.tv_license;
                }
            }
            
            // For Rent to HMO, also add specific fields with prefix
            if (calcType === 'rent-to-hmo') {
                if (calcData.deposit !== undefined) {
                    data[`${prefix}deposit`] = calcData.deposit;
                }
                if (calcData.survey_costs !== undefined) {
                    data[`${prefix}survey_costs`] = calcData.survey_costs;
                }
                if (calcData.legal_fees !== undefined) {
                    data[`${prefix}legal_fees`] = calcData.legal_fees;
                }
                if (calcData.reference_fees !== undefined) {
                    data[`${prefix}reference_fees`] = calcData.reference_fees;
                }
                if (calcData.monthly_rent_to_owner !== undefined) {
                    data[`${prefix}monthly_rent_to_owner`] = calcData.monthly_rent_to_owner;
                }
                if (calcData.chosen_strategy !== undefined) {
                    data[`${prefix}chosen_strategy`] = calcData.chosen_strategy;
                }
                if (calcData.strategy !== undefined) {
                    data[`${prefix}strategy`] = calcData.strategy;
                }
                if (calcData.nightly_rate !== undefined) {
                    data[`${prefix}nightly_rate`] = calcData.nightly_rate;
                }
                if (calcData.occupancy_rate !== undefined) {
                    data[`${prefix}occupancy_rate`] = calcData.occupancy_rate;
                }
                if (calcData.total_rental_income !== undefined) {
                    data[`${prefix}total_rental_income`] = calcData.total_rental_income;
                }
                // Add all room rents
                for (let i = 1; i <= 10; i++) {
                    const roomRent = calcData[`room_${i}_rent`];
                    if (roomRent !== undefined) {
                        data[`${prefix}room_${i}_rent`] = roomRent;
                    }
                }
                if (calcData.council_tax !== undefined) {
                    data[`${prefix}council_tax`] = calcData.council_tax;
                }
                if (calcData.utilities !== undefined) {
                    data[`${prefix}utilities`] = calcData.utilities;
                }
                if (calcData.water !== undefined) {
                    data[`${prefix}water`] = calcData.water;
                }
                if (calcData.broadband_tv !== undefined) {
                    data[`${prefix}broadband_tv`] = calcData.broadband_tv;
                }
                if (calcData.insurance !== undefined) {
                    data[`${prefix}insurance`] = calcData.insurance;
                }
                if (calcData.communal_tv_license !== undefined) {
                    data[`${prefix}communal_tv_license`] = calcData.communal_tv_license;
                }
                if (calcData.booking_fees !== undefined) {
                    data[`${prefix}booking_fees`] = calcData.booking_fees;
                }
            }
        }
        
        data[`calculator_${calcType}`] = calcData;
    });
    
    // Get all other input fields (property info, EPC, location, etc.)
    const allInputs = document.querySelectorAll('input, textarea, select');
    allInputs.forEach(input => {
        // Skip calculator-specific fields (they're handled above) and checkboxes
        if (input.type === 'checkbox' || input.dataset.calculator) {
            return;
        }
        if (input.id && !input.id.startsWith('calc-')) {
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
        // Property Info
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
        
        // Standard BTL / Purchase Calculator (default)
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
        
        // BRR Calculator
        'refurb_cost': '£30,000',
        'after_refurb_value': '£350,000',
        'refinance_ltv': '75',
        'bridging_interest': '£500',
        'vacant_period': '3',
        
        // Flip Calculator
        'sale_price': '£320,000',
        'holding_period': '6',
        'legal_fees_sale': '£1,200',
        'estate_agent_fees': '£3,200',
        'finance_cost': '£2,500',
        
        // Holiday Let Calculator
        'weekly_rent': '£450',
        'occupancy_rate': '65', // Default, will be overridden per calculator
        'management_fee': '20', // Percentage for Holiday Let
        'cleaning_fee': '£50',
        
        // Rent to HMO Calculator
        'monthly_rent_paid': '£1,200',
        'number_of_rooms': '5',
        'rent_per_room': '£550',
        'management_fee_annual': '£1,200', // Annual currency for Rent to HMO
        
        // Rent to Serviced Accommodation Calculator
        'daily_rate': '£85',
        // Note: occupancy_rate and management_fee values will be set per calculator below
        
        // EPC
        'epc_rating': 'C',
        'current_energy_cost': '£1,200',
        'potential_energy_cost': '£800',
        'co2_current': '3.2',
        'co2_potential': '1.8',
        
        // Internet / Broadband
        'broadband_available': 'Yes (FTTP available)',
        'download_speed': '1 Gbps',
        'upload_speed': '100 Mbps'
    };

    // Store mock data globally for use in dynamically created calculator fields
    window.mockDataStore = mockData;
    
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
    
    // Auto-select Standard BTL calculator and show fields
    setTimeout(() => {
        const standardBTLCheckbox = document.getElementById('calc-standard-btl');
        if (standardBTLCheckbox) {
            standardBTLCheckbox.checked = true;
            updateCalculatorSelection();
        }
    }, 100);
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

            // Debug logging
            console.log('[Frontend] Sending to backend:', {
                selected_calculators: data.selected_calculators,
                calculator_type: data.calculator_type,
                hasHolidayLet: data.selected_calculators?.includes('holiday-let'),
                hasBRR: data.selected_calculators?.includes('brr')
            });
            console.log('Frontend - is_brr_calculator:', data.is_brr_calculator);
            console.log('Frontend - calculator data keys:', Object.keys(data).filter(k => k.startsWith('calculator_') || k.startsWith('brr_')));
            console.log('Frontend - BRR fields count:', Object.keys(data).filter(k => k.startsWith('brr_')).length);
            if (data.calculator_brr) {
                console.log('Frontend - calculator_brr keys:', Object.keys(data.calculator_brr));
                console.log('Frontend - calculator_brr sample:', {
                    purchase_price: data.calculator_brr.purchase_price,
                    total_investment: data.calculator_brr.total_investment,
                    financing_type: data.calculator_brr.financing_type,
                    money_left_in: data.calculator_brr.money_left_in,
                    roi: data.calculator_brr.roi
                });
            }
            
            // Debug: Log which backend URL is being used
            console.log('🔗 Using backend URL:', BACKEND_URL);
            console.log('🔗 localStorage backend_url:', localStorage.getItem('backend_url'));
            console.log('🔗 Current hostname:', window.location.hostname);
            
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
                
                const z = 15; // Zoom level 15 for more detailed, zoomed-in view
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
                
                // Fetch city-specific images using backend search endpoint
                const fetchCityImages = async () => {
                    const citySearch = city.toLowerCase();
                    const cityImages = [];
                    
                    try {
                        // Get image URLs from backend
                        const searchUrl = `${BACKEND_URL}/search-city-images?city=${encodeURIComponent(citySearch)}`;
                        const searchResponse = await fetch(searchUrl);
                        
                        if (searchResponse.ok) {
                            const imageData = await searchResponse.json();
                            if (imageData.images && imageData.images.length > 0) {
                                // Fetch each image through proxy and convert to blob URL
                                // Try up to 3 images, with fallback options if one fails
                                for (let idx = 0; idx < 3; idx++) {
                                    let imageLoaded = false;
                                    const imageName = ['Landmark', 'City View', 'Architecture'][idx];
                                    
                                    // Try primary image, then fallbacks if available
                                    for (let attempt = 0; attempt < 3 && !imageLoaded; attempt++) {
                                        const imgIndex = idx + attempt; // Use primary or fallback
                                        if (imgIndex >= imageData.images.length) break;
                                        
                                        const imgUrl = imageData.images[imgIndex];
                                        const proxyUrl = `${BACKEND_URL}/proxy-image?url=${encodeURIComponent(imgUrl)}`;
                                        
                                        try {
                                            console.log(`Fetching city image ${idx + 1} (attempt ${attempt + 1}) from:`, imgUrl);
                                            const response = await fetch(proxyUrl);
                                            if (response.ok) {
                                                const blob = await response.blob();
                                                const blobUrl = URL.createObjectURL(blob);
                                                cityImages.push({
                                                    url: blobUrl,
                                                    name: `${city} - ${imageName}`
                                                });
                                                console.log(`Successfully loaded city image ${idx + 1}`);
                                                imageLoaded = true;
                                            } else {
                                                console.warn(`Failed to fetch image ${idx + 1} (attempt ${attempt + 1}):`, response.status);
                                            }
                                        } catch (err) {
                                            console.warn(`Error loading image ${idx + 1} (attempt ${attempt + 1}):`, err);
                                        }
                                    }
                                    
                                    if (!imageLoaded) {
                                        console.warn(`Could not load image ${idx + 1} after all attempts`);
                                    }
                                }
                            }
                        } else {
                            console.warn('Failed to search for city images:', searchResponse.status);
                        }
                    } catch (err) {
                        console.warn('Error fetching city images:', err);
                    }
                    
                    return cityImages;
                };
                
                // Add city images to city section (replace any existing)
                fetchCityImages().then(cityImages => {
                    if (cityImages.length > 0) {
                        imageSections.city = cityImages; // Replace all existing images
                        updateImageList('city');
                        console.log('Added city images to city section:', cityImages.length);
                    } else {
                        console.warn('No city images were loaded');
                    }
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

// Calculator selection and field management
const calculatorConfigs = {
    'standard-btl': {
        title: 'Standard Buy to Let',
        fields: [
            { label: 'Purchase Price', id: 'purchase_price', type: 'currency', required: true },
            { label: 'Deposit (%)', id: 'deposit_percent', type: 'number', required: true },
            { label: 'Mortgage Rate (%)', id: 'mortgage_rate', type: 'number', required: true },
            { label: 'Monthly Rent', id: 'monthly_rent', type: 'currency', required: true },
            { label: 'Stamp Duty', id: 'stamp_duty', type: 'currency' },
            { label: 'Survey Cost', id: 'survey_cost', type: 'currency' },
            { label: 'Legal Fees', id: 'legal_fees', type: 'currency' },
            { label: 'Loan Set-up', id: 'loan_setup', type: 'currency' },
            { label: 'Council Tax (Annual)', id: 'council_tax', type: 'currency' },
            { label: 'Repairs / Maintenance (Annual)', id: 'repairs_maintenance', type: 'currency' },
            { label: 'Utilities (Annual)', id: 'utilities', type: 'currency' },
            { label: 'Water (Annual)', id: 'water', type: 'currency' },
            { label: 'Broadband / TV (Annual)', id: 'broadband_tv', type: 'currency' },
            { label: 'Insurance (Annual)', id: 'insurance', type: 'currency' }
        ]
    },
    'brr': {
        title: 'Buy Refurbish Refinance',
        fields: [
            { label: 'Purchase Price', id: 'purchase_price', type: 'currency', required: true },
            { label: 'Refurbishment Cost', id: 'refurb_cost', type: 'currency', required: true },
            { label: 'Deposit (%)', id: 'deposit_percent', type: 'number', required: true },
            { label: 'Mortgage Rate (%)', id: 'mortgage_rate', type: 'number', required: true },
            { label: 'Refinance LTV (%)', id: 'refinance_ltv', type: 'number', required: true },
            { label: 'Monthly Rent', id: 'monthly_rent', type: 'currency', required: true },
            { label: 'Stamp Duty', id: 'stamp_duty', type: 'currency' },
            { label: 'Survey Cost', id: 'survey_cost', type: 'currency' },
            { label: 'Legal Fees', id: 'legal_fees', type: 'currency' },
            { label: 'Council Tax (Annual)', id: 'council_tax', type: 'currency' },
            { label: 'Repairs / Maintenance (Annual)', id: 'repairs_maintenance', type: 'currency' },
            { label: 'Insurance (Annual)', id: 'insurance', type: 'currency' }
        ]
    },
    'flip': {
        title: 'Flip',
        fields: [
            { label: 'Purchase Price', id: 'purchase_price', type: 'currency', required: true },
            { label: 'Refurbishment Cost', id: 'refurb_cost', type: 'currency', required: true },
            { label: 'Expected Sale Price', id: 'sale_price', type: 'currency', required: true },
            { label: 'Holding Period (Months)', id: 'holding_period', type: 'number', required: true },
            { label: 'Stamp Duty', id: 'stamp_duty', type: 'currency' },
            { label: 'Survey Cost', id: 'survey_cost', type: 'currency' },
            { label: 'Legal Fees (Purchase)', id: 'legal_fees', type: 'currency' },
            { label: 'Legal Fees (Sale)', id: 'legal_fees_sale', type: 'currency' },
            { label: 'Estate Agent Fees', id: 'estate_agent_fees', type: 'currency' },
            { label: 'Finance Cost (if applicable)', id: 'finance_cost', type: 'currency' }
        ]
    },
    'holiday-let': {
        title: 'Holiday Let',
        fields: [
            { label: 'Purchase Price', id: 'purchase_price', type: 'currency', required: true },
            { label: 'Deposit (%)', id: 'deposit_percent', type: 'number', required: true },
            { label: 'Mortgage Rate (%)', id: 'mortgage_rate', type: 'number', required: true },
            { label: 'Weekly Rental Income', id: 'weekly_rent', type: 'currency', required: true },
            { label: 'Occupancy Rate (%)', id: 'occupancy_rate', type: 'number', required: true },
            { label: 'Management Fee (%)', id: 'management_fee', type: 'number' },
            { label: 'Cleaning Fee (per booking)', id: 'cleaning_fee', type: 'currency' },
            { label: 'Council Tax (Annual)', id: 'council_tax', type: 'currency' },
            { label: 'Insurance (Annual)', id: 'insurance', type: 'currency' },
            { label: 'Utilities (Annual)', id: 'utilities', type: 'currency' }
        ]
    },
    'rent-to-hmo': {
        title: 'Rent to HMO',
        fields: [
            { label: 'Monthly Rent (to Landlord)', id: 'monthly_rent_paid', type: 'currency', required: true },
            { label: 'Number of Rooms', id: 'number_of_rooms', type: 'number', required: true },
            { label: 'Monthly Rent per Room', id: 'rent_per_room', type: 'currency', required: true },
            { label: 'Occupancy Rate (%)', id: 'occupancy_rate', type: 'number', required: true },
            { label: 'Council Tax (Annual)', id: 'council_tax', type: 'currency' },
            { label: 'Utilities (Annual)', id: 'utilities', type: 'currency' },
            { label: 'Insurance (Annual)', id: 'insurance', type: 'currency' },
            { label: 'Management Fee (Annual)', id: 'management_fee_annual', type: 'currency' }
        ]
    },
    'rent-to-serviced': {
        title: 'Rent to Serviced Accommodation',
        fields: [
            { label: 'Monthly Rent (to Landlord)', id: 'monthly_rent_paid', type: 'currency', required: true },
            { label: 'Daily Rate', id: 'daily_rate', type: 'currency', required: true },
            { label: 'Occupancy Rate (%)', id: 'occupancy_rate', type: 'number', required: true },
            { label: 'Cleaning Fee (per booking)', id: 'cleaning_fee', type: 'currency' },
            { label: 'Management Fee (%)', id: 'management_fee', type: 'number' },
            { label: 'Council Tax (Annual)', id: 'council_tax', type: 'currency' },
            { label: 'Utilities (Annual)', id: 'utilities', type: 'currency' },
            { label: 'Insurance (Annual)', id: 'insurance', type: 'currency' }
        ]
    },
    'purchase': {
        title: 'Purchase Calculator',
        fields: [
            { label: 'Purchase Price', id: 'purchase_price', type: 'currency', required: true },
            { label: 'Deposit (%)', id: 'deposit_percent', type: 'number', required: true },
            { label: 'Mortgage Rate (%)', id: 'mortgage_rate', type: 'number', required: true },
            { label: 'Monthly Rent', id: 'monthly_rent', type: 'currency', required: true },
            { label: 'Stamp Duty', id: 'stamp_duty', type: 'currency' },
            { label: 'Survey Cost', id: 'survey_cost', type: 'currency' },
            { label: 'Legal Fees', id: 'legal_fees', type: 'currency' },
            { label: 'Loan Set-up', id: 'loan_setup', type: 'currency' },
            { label: 'Council Tax (Annual)', id: 'council_tax', type: 'currency' },
            { label: 'Repairs / Maintenance (Annual)', id: 'repairs_maintenance', type: 'currency' },
            { label: 'Utilities (Annual)', id: 'utilities', type: 'currency' },
            { label: 'Water (Annual)', id: 'water', type: 'currency' },
            { label: 'Broadband / TV (Annual)', id: 'broadband_tv', type: 'currency' },
            { label: 'Insurance (Annual)', id: 'insurance', type: 'currency' }
        ]
    }
};

function initializeCalculatorSelection() {
    const calculatorOptions = document.querySelectorAll('.calculator-option');
    const checkboxes = document.querySelectorAll('.calculator-checkbox input[type="checkbox"]');
    
    // Set BRR (Buy Refurbish Refinance) as default selected calculator
    // Unselect all other calculators first
    checkboxes.forEach(checkbox => {
        if (checkbox.dataset.calculator !== 'brr') {
            checkbox.checked = false;
        }
    });
    
    const brrCheckbox = document.querySelector('input[data-calculator="brr"]');
    if (brrCheckbox) {
        brrCheckbox.checked = true;
    }
    
    // Trigger update to show BRR fields
    updateCalculatorSelection();
    
    // Handle checkbox clicks
    checkboxes.forEach(checkbox => {
        checkbox.addEventListener('change', (e) => {
            e.stopPropagation(); // Prevent option click
            updateCalculatorSelection();
        });
    });
    
    // Handle option clicks (toggle checkbox)
    calculatorOptions.forEach(option => {
        option.addEventListener('click', (e) => {
            // Don't toggle if clicking directly on checkbox
            if (e.target.type === 'checkbox') return;
            
            const checkbox = option.querySelector('input[type="checkbox"]');
            if (checkbox) {
                checkbox.checked = !checkbox.checked;
                updateCalculatorSelection();
            }
        });
    });
}

function updateCalculatorSelection() {
    const checkboxes = document.querySelectorAll('.calculator-checkbox input[type="checkbox"]:checked');
    const selectedCalculators = Array.from(checkboxes).map(cb => cb.dataset.calculator);
    const calculatorOptions = document.querySelectorAll('.calculator-option');
    
    // Update visual selection
    calculatorOptions.forEach(option => {
        const checkbox = option.querySelector('input[type="checkbox"]');
        if (checkbox && checkbox.checked) {
            option.classList.add('selected');
        } else {
            option.classList.remove('selected');
        }
    });
    
    // Show/hide calculator fields
    showCalculatorFields(selectedCalculators);
}

// Create Simple Buy to Let calculator
function createSimpleBuyToLetCalculator(calculatorType) {
    console.log('[Frontend] createSimpleBuyToLetCalculator called for:', calculatorType);
    const section = document.createElement('div');
    section.className = 'calculator-fields-section propertyengine-calculator';
    section.dataset.calculator = calculatorType;
    
    section.innerHTML = `
        <div class="pe-calculator-topbar">
            <h2 class="pe-calculator-title">Buy to Let Calculator</h2>
            <div class="pe-calculator-actions">
                <label class="pe-switch-label">
                    <input type="checkbox" class="pe-switch-detailed" id="${calculatorType}_detailed_view">
                    <span>Switch to detailed view</span>
                </label>
                <button class="pe-btn-save">Save Calculator</button>
                <div class="pe-roi-display">
                    <div class="pe-roi-label">Return on Investment</div>
                    <div class="pe-roi-value" id="${calculatorType}_roi">-%</div>
                </div>
            </div>
        </div>
        
        <div class="pe-calculator-content">
            <div class="pe-calculator-left">
                <div class="pe-section">
                    <h3 class="pe-section-title">Purchase</h3>
                    <div class="pe-field">
                        <label class="pe-field-label required">Purchase Price</label>
                        <input type="text" class="pe-input" id="${calculatorType}_purchase_price" data-original-id="purchase_price" data-calculator="${calculatorType}" placeholder="Required">
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Stamp Duty</label>
                        <input type="text" class="pe-input" id="${calculatorType}_stamp_duty" data-original-id="stamp_duty" data-calculator="${calculatorType}" value="£ 0" readonly>
                    </div>
                    <div class="pe-field pe-field-detailed" style="display: none;">
                        <label class="pe-field-label">Survey Costs</label>
                        <input type="text" class="pe-input" id="${calculatorType}_survey_costs" data-original-id="survey_costs" data-calculator="${calculatorType}" value="£ 500">
                    </div>
                    <div class="pe-field pe-field-detailed" style="display: none;">
                        <label class="pe-field-label">Legal Fees</label>
                        <input type="text" class="pe-input" id="${calculatorType}_legal_fees" data-original-id="legal_fees" data-calculator="${calculatorType}" value="£ 1,500">
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Total Investment Required</label>
                        <div class="pe-field-with-action">
                            <input type="text" class="pe-input" id="${calculatorType}_total_investment" value="£ 0" readonly>
                            <button class="pe-btn-detail">Detail</button>
                        </div>
                    </div>
                    <a href="#" class="pe-link-add pe-field-detailed" style="display: none;" id="${calculatorType}_add_purchase_cost">+ Add additional purchase cost</a>
                </div>
                
                <div class="pe-section">
                    <h3 class="pe-section-title">Financing</h3>
                    <div class="pe-field pe-field-inline">
                        <label class="pe-field-label">Type</label>
                        <div class="pe-financing-type-group">
                            <button type="button" class="pe-financing-type pe-financing-type-active" data-type="mortgage" data-calculator="${calculatorType}">
                                <span class="pe-financing-check">✓</span>
                                <span>Mortgage</span>
                            </button>
                            <button type="button" class="pe-financing-type" data-type="bridging" data-calculator="${calculatorType}">
                                <span class="pe-financing-check">✓</span>
                                <span>Bridging Finance</span>
                            </button>
                            <button type="button" class="pe-financing-type" data-type="cash" data-calculator="${calculatorType}">
                                <span class="pe-financing-check">✓</span>
                                <span>Cash</span>
                            </button>
                            <input type="hidden" name="${calculatorType}_financing_type" value="mortgage" id="${calculatorType}_financing_type_hidden">
                        </div>
                    </div>
                    <div class="pe-field pe-field-detailed pe-field-mortgage" style="display: none;">
                        <label class="pe-field-label">Mortgage Set-up Fee</label>
                        <div class="pe-field-with-action">
                            <input type="text" class="pe-input" id="${calculatorType}_mortgage_setup_fee" data-original-id="mortgage_setup_fee" data-calculator="${calculatorType}" value="£ 1,000">
                            <div class="pe-toggle-group">
                                <button type="button" class="pe-toggle-small pe-toggle-small-active" data-fee-type="currency" data-fee-for="mortgage" data-calculator="${calculatorType}">£</button>
                                <button type="button" class="pe-toggle-small" data-fee-type="percent" data-fee-for="mortgage" data-calculator="${calculatorType}">%</button>
                            </div>
                        </div>
                    </div>
                    <div class="pe-field pe-field-detailed pe-field-mortgage" style="display: none;">
                        <label class="pe-field-label">Mortgage Loan to Value</label>
                        <input type="text" class="pe-input" id="${calculatorType}_mortgage_ltv" data-original-id="mortgage_ltv" data-calculator="${calculatorType}" value="75 %">
                    </div>
                    <div class="pe-field pe-field-detailed pe-field-mortgage pe-field-inline" style="display: none;">
                        <label class="pe-field-label">Mortgage Type</label>
                        <div class="pe-financing-type-group">
                            <button type="button" class="pe-financing-type pe-financing-type-active" data-mortgage-type="interest_only" data-calculator="${calculatorType}">
                                <span class="pe-financing-check">✓</span>
                                <span>Interest Only</span>
                            </button>
                            <button type="button" class="pe-financing-type" data-mortgage-type="repayment" data-calculator="${calculatorType}">
                                <span class="pe-financing-check">✓</span>
                                <span>Repayment</span>
                            </button>
                        </div>
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Mortgage Payments / pcm</label>
                        <input type="text" class="pe-input" id="${calculatorType}_mortgage_payments" data-original-id="mortgage_payments" data-calculator="${calculatorType}" value="£ 0">
                    </div>
                    <div class="pe-field pe-field-detailed pe-field-mortgage" style="display: none;">
                        <label class="pe-field-label">Mortgage Interest Rate (APR)</label>
                        <input type="text" class="pe-input" id="${calculatorType}_mortgage_interest_rate" data-original-id="mortgage_interest_rate" data-calculator="${calculatorType}" value="5.5 %">
                    </div>
                    <div class="pe-field pe-field-detailed pe-field-mortgage pe-field-repayment" style="display: none;">
                        <label class="pe-field-label">Mortgage Term (Years)</label>
                        <input type="text" class="pe-input" id="${calculatorType}_mortgage_term_years" data-original-id="mortgage_term_years" data-calculator="${calculatorType}" value="25">
                    </div>
                    <div class="pe-field pe-field-detailed pe-field-mortgage" style="display: none;">
                        <label class="pe-field-label">Mortgage Required</label>
                        <input type="text" class="pe-input" id="${calculatorType}_mortgage_required" value="£ 0" readonly>
                    </div>
                </div>
                
                <div class="pe-section">
                    <div class="pe-section-header">
                        <h3 class="pe-section-title">Refurb</h3>
                        <label class="pe-toggle">
                            <input type="checkbox" class="pe-toggle-input" id="${calculatorType}_refurb_enabled">
                            <span class="pe-toggle-slider"></span>
                        </label>
                    </div>
                    <div class="pe-section-content" id="${calculatorType}_refurb_content" style="display: none;">
                        <div class="pe-field">
                            <label class="pe-field-label">Refurb Cost</label>
                            <input type="text" class="pe-input" id="${calculatorType}_refurb_cost" data-original-id="refurb_cost" data-calculator="${calculatorType}" value="£ 0">
                        </div>
                    </div>
                </div>
            </div>
            
            <div class="pe-calculator-right">
                <div class="pe-section pe-field-detailed" style="display: none;">
                    <h3 class="pe-section-title">Refinance</h3>
                    <div class="pe-field">
                        <label class="pe-field-label">Estimated Market Value</label>
                        <input type="text" class="pe-input" id="${calculatorType}_estimated_market_value" data-original-id="estimated_market_value" data-calculator="${calculatorType}" placeholder="£ 0">
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Mortgage Set-up Fee</label>
                        <div class="pe-field-with-action">
                            <input type="text" class="pe-input" id="${calculatorType}_refinance_setup_fee" data-original-id="refinance_setup_fee" data-calculator="${calculatorType}" value="£ 0">
                            <div class="pe-toggle-group">
                                <button type="button" class="pe-toggle-small pe-toggle-small-active" data-fee-type="currency" data-fee-for="refinance" data-calculator="${calculatorType}">£</button>
                                <button type="button" class="pe-toggle-small" data-fee-type="percent" data-fee-for="refinance" data-calculator="${calculatorType}">%</button>
                            </div>
                        </div>
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Mortgage Loan to Value</label>
                        <input type="text" class="pe-input" id="${calculatorType}_refinance_ltv" data-original-id="refinance_ltv" data-calculator="${calculatorType}" value="75 %">
                    </div>
                    <div class="pe-field pe-field-inline">
                        <label class="pe-field-label">Mortgage Type</label>
                        <div class="pe-financing-type-group">
                            <button type="button" class="pe-financing-type pe-financing-type-active" data-refinance-mortgage-type="interest_only" data-calculator="${calculatorType}">
                                <span class="pe-financing-check">✓</span>
                                <span>Interest Only</span>
                            </button>
                            <button type="button" class="pe-financing-type" data-refinance-mortgage-type="repayment" data-calculator="${calculatorType}">
                                <span class="pe-financing-check">✓</span>
                                <span>Repayment</span>
                            </button>
                        </div>
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Mortgage Payments / pcm</label>
                        <input type="text" class="pe-input" id="${calculatorType}_refinance_mortgage_payments" value="£ 0" readonly>
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Mortgage Interest Rate (APR)</label>
                        <input type="text" class="pe-input" id="${calculatorType}_refinance_interest_rate" data-original-id="refinance_interest_rate" data-calculator="${calculatorType}" placeholder="5.5 %">
                    </div>
                    <div class="pe-field pe-field-refinance-repayment" style="display: none;">
                        <label class="pe-field-label">Mortgage Term (Years)</label>
                        <input type="text" class="pe-input" id="${calculatorType}_refinance_mortgage_term_years" data-original-id="refinance_mortgage_term_years" data-calculator="${calculatorType}" value="25">
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Locked In Equity</label>
                        <input type="text" class="pe-input" id="${calculatorType}_locked_in_equity" value="£ 0" readonly>
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Money Left In</label>
                        <input type="text" class="pe-input" id="${calculatorType}_money_left_in" value="£ 0" readonly>
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Ideal Purchase Price (max)</label>
                        <input type="text" class="pe-input" id="${calculatorType}_ideal_purchase_price" value="£ 0" readonly>
                    </div>
                    <div id="${calculatorType}_additional_refinance_costs"></div>
                    <a href="#" class="pe-link-add" id="${calculatorType}_add_refinance_cost">+ Add additional refinance cost</a>
                </div>
                
                <div class="pe-section">
                    <h3 class="pe-section-title">Rental Income</h3>
                    <div class="pe-field">
                        <label class="pe-field-label required">Monthly Rent</label>
                        <input type="text" class="pe-input" id="${calculatorType}_monthly_rent" data-original-id="monthly_rent" data-calculator="${calculatorType}" placeholder="Required">
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Gross Yield</label>
                        <input type="text" class="pe-input" id="${calculatorType}_gross_yield" value="-" readonly>
                    </div>
                </div>
                
                <div class="pe-section pe-field-detailed" style="display: none;">
                    <h3 class="pe-section-title">Ongoing Costs</h3>
                    <div class="pe-subsection">
                        <h4 class="pe-subsection-title">Annual Expenses</h4>
                        <div class="pe-field">
                            <label class="pe-field-label">% of Income on Maintenance</label>
                            <input type="text" class="pe-input" id="${calculatorType}_maintenance_percent" data-original-id="maintenance_percent" data-calculator="${calculatorType}" value="10 %">
                        </div>
                        <div id="${calculatorType}_additional_annual_expenses"></div>
                        <a href="#" class="pe-link-add" id="${calculatorType}_add_annual_expense">+ Add additional annual expense</a>
                    </div>
                    <div class="pe-subsection">
                        <h4 class="pe-subsection-title">Monthly Expenses</h4>
                        <div class="pe-field">
                            <label class="pe-field-label">Mortgage Payments</label>
                            <input type="text" class="pe-input" id="${calculatorType}_ongoing_mortgage_payments" value="£ 0" readonly>
                        </div>
                        <div class="pe-field">
                            <label class="pe-field-label">Insurance</label>
                            <input type="text" class="pe-input" id="${calculatorType}_ongoing_insurance" data-original-id="ongoing_insurance" data-calculator="${calculatorType}" value="£ 40">
                        </div>
                        <div class="pe-field">
                            <label class="pe-field-label">Agent Fees</label>
                            <input type="text" class="pe-input" id="${calculatorType}_agent_fees" data-original-id="agent_fees" data-calculator="${calculatorType}" value="10 %">
                            <div class="pe-toggle-group" style="display: inline-block; margin-left: 8px;">
                                <button type="button" class="pe-toggle-small pe-toggle-small-active" data-fee-type="percent" data-fee-for="agent" data-calculator="${calculatorType}">%</button>
                                <button type="button" class="pe-toggle-small" data-fee-type="currency" data-fee-for="agent" data-calculator="${calculatorType}">£</button>
                            </div>
                        </div>
                        <div id="${calculatorType}_additional_monthly_expenses"></div>
                        <a href="#" class="pe-link-add" id="${calculatorType}_add_monthly_expense">+ Add additional monthly expense</a>
                    </div>
                </div>
                
                <div class="pe-section">
                    <h3 class="pe-section-title">Summary</h3>
                    <div class="pe-field">
                        <label class="pe-field-label">Total Annual Expenses</label>
                        <div class="pe-field-with-action">
                            <input type="text" class="pe-input" id="${calculatorType}_total_annual_expenses" value="£ 0" readonly>
                            <button class="pe-btn-detail">Detail</button>
                        </div>
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Annual Profit</label>
                        <input type="text" class="pe-input" id="${calculatorType}_annual_profit" value="£ 0" readonly>
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Monthly Profit</label>
                        <input type="text" class="pe-input" id="${calculatorType}_monthly_profit" value="£ 0" readonly>
                    </div>
                    <a href="#" class="pe-link-add pe-field-detailed" style="display: none;" id="${calculatorType}_edit_thresholds">Edit target thresholds</a>
                </div>
            </div>
        </div>
        
        <div class="pe-metrics">
            <h4 class="pe-metrics-section-title">Short Term Metrics</h4>
            <div class="pe-metrics-short">
                <div class="pe-metric-box">
                    <div class="pe-metric-label">Gross yield</div>
                    <div class="pe-metric-value" id="${calculatorType}_gross_yield_metric">-%</div>
                    <span class="pe-info-icon" title="Gross yield">i</span>
                </div>
                <div class="pe-metric-box">
                    <div class="pe-metric-label">Return on capital employed</div>
                    <div class="pe-metric-value" id="${calculatorType}_roce">-%</div>
                    <span class="pe-info-icon" title="Return on capital employed">i</span>
                </div>
                <div class="pe-metric-box">
                    <div class="pe-metric-label">Net yield</div>
                    <div class="pe-metric-value" id="${calculatorType}_net_yield">-%</div>
                    <span class="pe-info-icon" title="Net yield">i</span>
                </div>
            </div>
            <h4 class="pe-metrics-section-title">Long Term Rental Metrics</h4>
            <div class="pe-metrics-long">
                <div class="pe-metric-box pe-metric-box-blue">
                    <div class="pe-metric-label">Equity in 10 years</div>
                    <div class="pe-metric-value" id="${calculatorType}_equity_10_years">£0</div>
                    <span class="pe-info-icon" title="Projected equity after 10 years">i</span>
                </div>
                <div class="pe-metrics-appreciation">
                    <label class="pe-field-label">Annual Property Appreciation</label>
                    <input type="text" class="pe-input pe-input-small" id="${calculatorType}_appreciation" value="5.5%" data-calculator="${calculatorType}">
                </div>
            </div>
        </div>
    `;
    
    // Setup event handlers
    setupSimpleBuyToLetCalculatorEvents(section, calculatorType);
    
    return section;
}

function setupSimpleBuyToLetCalculatorEvents(section, calculatorType) {
    console.log('[Frontend] setupSimpleBuyToLetCalculatorEvents called for:', calculatorType);
    
    // Detailed view toggle
    const detailedViewToggle = section.querySelector(`#${calculatorType}_detailed_view`);
    if (detailedViewToggle) {
        detailedViewToggle.addEventListener('change', (e) => {
            const isDetailed = e.target.checked;
            const label = detailedViewToggle.parentElement.querySelector('span');
            if (label) {
                label.textContent = isDetailed ? 'Switch to simple view' : 'Switch to detailed view';
            }
            
            // Show/hide detailed fields
            const detailedFields = section.querySelectorAll('.pe-field-detailed');
            detailedFields.forEach(field => {
                field.style.display = isDetailed ? 'block' : 'none';
            });
            
            // Show/hide detailed sections
            const detailedSections = section.querySelectorAll('.pe-section.pe-field-detailed');
            detailedSections.forEach(sec => {
                sec.style.display = isDetailed ? 'block' : 'none';
            });
            
            // Show/hide mortgage-specific fields based on financing type
            const financingTypeHidden = section.querySelector(`#${calculatorType}_financing_type_hidden`);
            const financingType = financingTypeHidden?.value || 'mortgage';
            const mortgageFields = section.querySelectorAll('.pe-field-mortgage');
            mortgageFields.forEach(field => {
                if (isDetailed && financingType === 'mortgage') {
                    field.style.display = 'block';
                } else if (!isDetailed) {
                    field.style.display = 'none';
                } else {
                    field.style.display = 'none';
                }
            });
            
            // Recalculate when toggling view
            calculateSimpleBuyToLetValues(calculatorType);
        });
    }
    
    // Financing type buttons
    const financingTypeButtons = section.querySelectorAll(`[data-type][data-calculator="${calculatorType}"]`);
    financingTypeButtons.forEach(btn => {
        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            financingTypeButtons.forEach(b => b.classList.remove('pe-financing-type-active'));
            btn.classList.add('pe-financing-type-active');
            
            const financingType = btn.dataset.type;
            const hiddenInput = section.querySelector(`#${calculatorType}_financing_type_hidden`);
            if (hiddenInput) {
                hiddenInput.value = financingType;
            }
            
            calculateSimpleBuyToLetValues(calculatorType);
        });
    });
    
    // Refurb toggle
    const refurbToggle = section.querySelector(`#${calculatorType}_refurb_enabled`);
    if (refurbToggle) {
        refurbToggle.addEventListener('change', (e) => {
            const content = section.querySelector(`#${calculatorType}_refurb_content`);
            if (content) {
                content.style.display = e.target.checked ? 'block' : 'none';
            }
            calculateSimpleBuyToLetValues(calculatorType);
        });
    }
    
    // Mortgage type buttons (Initial Financing)
    const initialMortgageTypeButtons = section.querySelectorAll(`[data-mortgage-type][data-calculator="${calculatorType}"]`);
    initialMortgageTypeButtons.forEach(btn => {
        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            initialMortgageTypeButtons.forEach(b => b.classList.remove('pe-financing-type-active'));
            btn.classList.add('pe-financing-type-active');
            
            const mortgageType = btn.dataset.mortgageType;
            const repaymentFields = section.querySelectorAll('.pe-field-repayment');
            const detailedViewToggle = section.querySelector(`#${calculatorType}_detailed_view`);
            const isDetailed = detailedViewToggle?.checked || false;
            
            if (isDetailed) {
                repaymentFields.forEach(field => {
                    field.style.display = mortgageType === 'repayment' ? 'block' : 'none';
                });
            }
            
            calculateSimpleBuyToLetValues(calculatorType);
        });
    });
    
    // Refinance mortgage type buttons
    const refinanceMortgageTypeButtons = section.querySelectorAll(`[data-refinance-mortgage-type][data-calculator="${calculatorType}"]`);
    refinanceMortgageTypeButtons.forEach(btn => {
        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            refinanceMortgageTypeButtons.forEach(b => b.classList.remove('pe-financing-type-active'));
            btn.classList.add('pe-financing-type-active');
            
            const refinanceMortgageType = btn.dataset.refinanceMortgageType;
            const refinanceRepaymentFields = section.querySelectorAll('.pe-field-refinance-repayment');
            
            refinanceRepaymentFields.forEach(field => {
                field.style.display = refinanceMortgageType === 'repayment' ? 'block' : 'none';
            });
            
            calculateSimpleBuyToLetValues(calculatorType);
        });
    });
    
    // Fee type toggles (Mortgage Set-up Fee, Refinance Set-up Fee, Agent Fees)
    const feeTypeButtons = section.querySelectorAll(`[data-fee-type][data-calculator="${calculatorType}"]`);
    feeTypeButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const feeFor = btn.dataset.feeFor;
            const feeType = btn.dataset.feeType;
            const feeButtons = section.querySelectorAll(`[data-fee-for="${feeFor}"][data-calculator="${calculatorType}"]`);
            
            feeButtons.forEach(b => b.classList.remove('pe-toggle-small-active'));
            btn.classList.add('pe-toggle-small-active');
            
            // Update input value format based on selected type
            let inputId = '';
            if (feeFor === 'mortgage') {
                inputId = `${calculatorType}_mortgage_setup_fee`;
            } else if (feeFor === 'refinance') {
                inputId = `${calculatorType}_refinance_setup_fee`;
            } else if (feeFor === 'agent') {
                inputId = `${calculatorType}_agent_fees`;
            }
            
            const input = document.getElementById(inputId);
            if (input) {
                const currentValue = input.value.replace(/[£%\s,]/g, '');
                if (feeType === 'currency') {
                    input.value = `£ ${currentValue}`;
                } else if (feeType === 'percent') {
                    input.value = `${currentValue} %`;
                }
                input.dataset.feeType = feeType;
            }
            
            calculateSimpleBuyToLetValues(calculatorType);
        });
    });
    
    // Update mortgage fields visibility when financing type changes
    const updateMortgageFieldsVisibility = () => {
        const financingTypeHidden = section.querySelector(`#${calculatorType}_financing_type_hidden`);
        const financingType = financingTypeHidden?.value || 'mortgage';
        const detailedViewToggle = section.querySelector(`#${calculatorType}_detailed_view`);
        const isDetailed = detailedViewToggle?.checked || false;
        const mortgageFields = section.querySelectorAll('.pe-field-mortgage');
        
        mortgageFields.forEach(field => {
            if (isDetailed && financingType === 'mortgage') {
                field.style.display = 'block';
            } else {
                field.style.display = 'none';
            }
        });
    };
    
    // Update mortgage fields when financing type changes
    financingTypeButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            setTimeout(updateMortgageFieldsVisibility, 50);
        });
    });
    
    // Add event listeners to all inputs
    const allInputs = section.querySelectorAll('input, select');
    allInputs.forEach(input => {
        if (input.id && !input.readOnly && input.type !== 'checkbox') {
            input.addEventListener('input', () => calculateSimpleBuyToLetValues(calculatorType));
            input.addEventListener('change', () => calculateSimpleBuyToLetValues(calculatorType));
        }
    });
}

function calculateSimpleBuyToLetValues(calculatorType) {
    console.log('[Frontend] calculateSimpleBuyToLetValues called for:', calculatorType);
    
    // Get input values
    const purchasePrice = parseCurrency(document.getElementById(`${calculatorType}_purchase_price`)?.value || '0');
    const surveyCosts = parseCurrency(document.getElementById(`${calculatorType}_survey_costs`)?.value || '0');
    const legalFees = parseCurrency(document.getElementById(`${calculatorType}_legal_fees`)?.value || '0');
    const monthlyRent = parseCurrency(document.getElementById(`${calculatorType}_monthly_rent`)?.value || '0');
    
    // Calculate stamp duty automatically (Buy to Let is an additional property)
    let stampDuty = 0;
    let stampDutyBreakdown = [];
    if (purchasePrice > 0) {
        // For Buy to Let, it's always an additional property (buy-to-let/second home)
        // Use current rates (from 1st April 2025)
        const isAdditionalProperty = true; // Buy to Let is always additional property
        stampDuty = calculateStampDuty(purchasePrice, 'current', false, false, false, stampDutyBreakdown, isAdditionalProperty);
    }
    
    // Update stamp duty field
    const stampDutyEl = document.getElementById(`${calculatorType}_stamp_duty`);
    if (stampDutyEl) {
        stampDutyEl.value = formatCurrency(stampDuty);
    }
    
    // Get financing type
    const hiddenInput = document.getElementById(`${calculatorType}_financing_type_hidden`);
    const financingType = hiddenInput?.value || 'mortgage';
    
    // Get refurb info
    const refurbToggle = document.getElementById(`${calculatorType}_refurb_enabled`);
    const isRefurbEnabled = refurbToggle ? refurbToggle.checked : false;
    const refurbCost = isRefurbEnabled ? parseCurrency(document.getElementById(`${calculatorType}_refurb_cost`)?.value || '0') : 0;
    
    // Calculate mortgage amount based on LTV
    let mortgageAmount = 0;
    let mortgageSetupFee = 0;
    if (financingType === 'mortgage') {
        const mortgageLTV = parseFloat(document.getElementById(`${calculatorType}_mortgage_ltv`)?.value?.replace(/[%\s]/g, '') || '75');
        mortgageAmount = purchasePrice * (mortgageLTV / 100);
        
        // Get mortgage setup fee
        const mortgageSetupFeeValue = document.getElementById(`${calculatorType}_mortgage_setup_fee`)?.value || '£ 1,000';
        if (mortgageSetupFeeValue.includes('%')) {
            const percent = parseFloat(mortgageSetupFeeValue.replace(/[%\s]/g, '')) || 0;
            mortgageSetupFee = mortgageAmount * (percent / 100);
        } else {
            mortgageSetupFee = parseCurrency(mortgageSetupFeeValue);
        }
    }
    
    // Calculate mortgage payments
    let monthlyMortgagePayment = 0;
    if (mortgageAmount > 0) {
        const mortgageInterestRateValue = document.getElementById(`${calculatorType}_mortgage_interest_rate`)?.value || '5.5 %';
        const mortgageInterestRate = parseFloat(mortgageInterestRateValue.replace(/[%\s]/g, '')) || 5.5;
        
        const mortgageTypeBtn = document.querySelector(`[data-mortgage-type="repayment"][data-calculator="${calculatorType}"]`);
        const isRepayment = mortgageTypeBtn?.classList.contains('pe-financing-type-active') || false;
        
        if (isRepayment) {
            const mortgageTermYears = parseFloat(document.getElementById(`${calculatorType}_mortgage_term_years`)?.value || '25');
            const monthlyRate = (mortgageInterestRate / 100) / 12;
            const numberOfPayments = mortgageTermYears * 12;
            
            if (monthlyRate > 0) {
                monthlyMortgagePayment = mortgageAmount * (monthlyRate * Math.pow(1 + monthlyRate, numberOfPayments)) /
                                        (Math.pow(1 + monthlyRate, numberOfPayments) - 1);
            } else {
                monthlyMortgagePayment = mortgageAmount / numberOfPayments;
            }
        } else {
            // Interest-only
            monthlyMortgagePayment = mortgageAmount * (mortgageInterestRate / 100) / 12;
        }
    }
    
    // Calculate deposit
    const deposit = purchasePrice - mortgageAmount;
    
    // Calculate total investment = deposit + stamp duty + survey costs + legal fees + mortgage setup fee + refurb cost
    const totalInvestment = deposit + stampDuty + surveyCosts + legalFees + mortgageSetupFee + refurbCost;
    
    // Calculate annual rent
    const annualRent = monthlyRent * 12;
    
    // Calculate gross yield
    const grossYield = purchasePrice > 0 ? (annualRent / purchasePrice) * 100 : 0;
    
    // Calculate annual expenses
    const annualMortgagePayments = monthlyMortgagePayment * 12;
    
    // Get ongoing costs (from detailed view if available)
    const maintenancePercentValue = document.getElementById(`${calculatorType}_maintenance_percent`)?.value || '10 %';
    const maintenancePercent = parseFloat(maintenancePercentValue.replace(/[%\s]/g, '')) || 10;
    const annualMaintenance = annualRent * (maintenancePercent / 100);
    
    const ongoingInsuranceValue = document.getElementById(`${calculatorType}_ongoing_insurance`)?.value || '£ 40';
    const monthlyInsurance = parseCurrency(ongoingInsuranceValue);
    const annualInsurance = monthlyInsurance * 12;
    
    // Agent fees
    let annualAgentFees = 0;
    const agentFeesValue = document.getElementById(`${calculatorType}_agent_fees`)?.value || '10 %';
    const agentFeesInput = document.getElementById(`${calculatorType}_agent_fees`);
    let agentFeeType = 'percent';
    if (agentFeesInput) {
        agentFeeType = agentFeesInput.dataset.feeType || 'percent';
        const activeToggle = document.querySelector(`[data-fee-for="agent"][data-calculator="${calculatorType}"].pe-toggle-small-active`);
        if (activeToggle) {
            agentFeeType = activeToggle.dataset.feeType || 'percent';
        }
    }
    
    if (agentFeeType === 'percent') {
        const agentFeePercent = parseFloat(agentFeesValue.replace(/[%\s]/g, '')) || 10;
        annualAgentFees = annualRent * (agentFeePercent / 100);
    } else {
        const monthlyAgentFees = parseCurrency(agentFeesValue);
        annualAgentFees = monthlyAgentFees * 12;
    }
    
    const totalAnnualExpenses = annualMortgagePayments + annualMaintenance + annualInsurance + annualAgentFees;
    
    // Calculate profits
    const annualProfit = annualRent - totalAnnualExpenses;
    const monthlyProfit = annualProfit / 12;
    
    // Calculate ROI and ROCE (same for BTL - based on total investment)
    const roi = totalInvestment > 0 ? (annualProfit / totalInvestment) * 100 : 0;
    const roce = roi; // For BTL, ROCE is the same as ROI
    
    // Calculate net yield
    const netYield = purchasePrice > 0 ? (annualProfit / purchasePrice) * 100 : 0;
    
    // Calculate equity in 10 years
    const appreciationValue = document.getElementById(`${calculatorType}_appreciation`)?.value || '5.5%';
    const appreciation = parseFloat(appreciationValue.replace(/[%\s]/g, '')) || 5.5;
    const futurePropertyValue = purchasePrice * Math.pow(1 + appreciation / 100, 10);
    
    // Mortgage balance after 10 years
    let mortgageBalance10Years = mortgageAmount; // Default for interest-only
    if (mortgageAmount > 0) {
        const mortgageTypeBtn = document.querySelector(`[data-mortgage-type="repayment"][data-calculator="${calculatorType}"]`);
        const isRepayment = mortgageTypeBtn?.classList.contains('pe-financing-type-active') || false;
        
        if (isRepayment) {
            const mortgageInterestRateValue = document.getElementById(`${calculatorType}_mortgage_interest_rate`)?.value || '5.5 %';
            const mortgageInterestRate = parseFloat(mortgageInterestRateValue.replace(/[%\s]/g, '')) || 5.5;
            const mortgageTermYears = parseFloat(document.getElementById(`${calculatorType}_mortgage_term_years`)?.value || '25');
            
            const monthlyRate = mortgageInterestRate / 100 / 12;
            const totalMonths = mortgageTermYears * 12;
            const monthsPaid = 10 * 12;
            
            if (totalMonths > monthsPaid && monthlyRate > 0) {
                const balanceFactor = (Math.pow(1 + monthlyRate, totalMonths) - Math.pow(1 + monthlyRate, monthsPaid)) / 
                                     (Math.pow(1 + monthlyRate, totalMonths) - 1);
                mortgageBalance10Years = mortgageAmount * balanceFactor;
            } else if (totalMonths <= monthsPaid) {
                mortgageBalance10Years = 0;
            }
        }
    }
    
    const equity10Years = futurePropertyValue - mortgageBalance10Years;
    
    // Update display fields
    const totalInvestmentEl = document.getElementById(`${calculatorType}_total_investment`);
    if (totalInvestmentEl) totalInvestmentEl.value = formatCurrency(totalInvestment);
    
    const mortgageRequiredEl = document.getElementById(`${calculatorType}_mortgage_required`);
    if (mortgageRequiredEl) mortgageRequiredEl.value = formatCurrency(mortgageAmount);
    
    const mortgagePaymentsEl = document.getElementById(`${calculatorType}_mortgage_payments`);
    if (mortgagePaymentsEl) mortgagePaymentsEl.value = formatCurrency(monthlyMortgagePayment);
    
    const ongoingMortgagePaymentsEl = document.getElementById(`${calculatorType}_ongoing_mortgage_payments`);
    if (ongoingMortgagePaymentsEl) ongoingMortgagePaymentsEl.value = formatCurrency(monthlyMortgagePayment);
    
    const grossYieldEl = document.getElementById(`${calculatorType}_gross_yield`);
    if (grossYieldEl) grossYieldEl.value = `${grossYield.toFixed(1)}%`;
    
    const totalAnnualExpensesEl = document.getElementById(`${calculatorType}_total_annual_expenses`);
    if (totalAnnualExpensesEl) totalAnnualExpensesEl.value = formatCurrency(totalAnnualExpenses);
    
    const annualProfitEl = document.getElementById(`${calculatorType}_annual_profit`);
    if (annualProfitEl) annualProfitEl.value = formatCurrency(annualProfit);
    
    const monthlyProfitEl = document.getElementById(`${calculatorType}_monthly_profit`);
    if (monthlyProfitEl) monthlyProfitEl.value = formatCurrency(monthlyProfit);
    
    const roiDisplayEl = document.getElementById(`${calculatorType}_roi_display`);
    if (roiDisplayEl) roiDisplayEl.value = `${roi.toFixed(1)}%`;
    
    const roiTopEl = document.getElementById(`${calculatorType}_roi`);
    if (roiTopEl) roiTopEl.textContent = `${roi.toFixed(1)}%`;
    
    // Update metrics
    const roceEl = document.getElementById(`${calculatorType}_roce`);
    if (roceEl) roceEl.textContent = `${roce.toFixed(1)}%`;
    
    const grossYieldMetricEl = document.getElementById(`${calculatorType}_gross_yield_metric`);
    if (grossYieldMetricEl) grossYieldMetricEl.textContent = `${grossYield.toFixed(1)}%`;
    
    const netYieldEl = document.getElementById(`${calculatorType}_net_yield`);
    if (netYieldEl) netYieldEl.textContent = `${netYield.toFixed(2)}%`;
    
    const equity10YearsEl = document.getElementById(`${calculatorType}_equity_10_years`);
    if (equity10YearsEl) equity10YearsEl.textContent = formatCurrency(equity10Years);
}

function createBRRCalculator(calculatorType) {
    console.log('[Frontend] createBRRCalculator called for:', calculatorType);
    const section = document.createElement('div');
    section.className = 'calculator-fields-section propertyengine-calculator';
    section.dataset.calculator = calculatorType;
    
    section.innerHTML = `
        <div class="pe-calculator-topbar">
            <h2 class="pe-calculator-title">Buy Refurbish Refinance Calculator</h2>
            <div class="pe-calculator-actions">
                <label class="pe-switch-label">
                    <input type="checkbox" class="pe-switch-detailed" id="${calculatorType}_detailed_view">
                    <span>Switch to detailed view</span>
                </label>
                <button class="pe-btn-save">Save Calculator</button>
                <div class="pe-roi-display">
                    <div class="pe-roi-label">Return on Investment</div>
                    <div class="pe-roi-value" id="${calculatorType}_roi">-%</div>
                </div>
            </div>
        </div>
        
        <div class="pe-calculator-content">
            <div class="pe-calculator-left">
                <div class="pe-section">
                    <h3 class="pe-section-title">Purchase</h3>
                    <div class="pe-field">
                        <label class="pe-field-label required">Purchase Price</label>
                        <input type="text" class="pe-input" id="${calculatorType}_purchase_price" data-original-id="purchase_price" data-calculator="${calculatorType}" placeholder="£ 0">
                    </div>
                    <div class="pe-field">
                        <div class="pe-field-header">
                            <label class="pe-field-label">Stamp Duty</label>
                            <button type="button" class="pe-toggle-expand" id="${calculatorType}_stamp_duty_toggle" data-target="${calculatorType}_stamp_duty_content">
                                <span class="pe-toggle-expand-icon">▼</span>
                            </button>
                        </div>
                        <div class="pe-field">
                            <input type="text" class="pe-input" id="${calculatorType}_stamp_duty" data-original-id="stamp_duty" data-calculator="${calculatorType}" value="£ 0" readonly>
                        </div>
                        <div class="pe-stamp-duty-container" id="${calculatorType}_stamp_duty_content" style="display: none;">
                            <div class="pe-stamp-duty-section">
                                <label class="pe-stamp-duty-section-label">Buyer Type</label>
                                <div class="pe-stamp-duty-options">
                                    <button type="button" class="pe-option-button" id="${calculatorType}_individual_moving" data-calculator="${calculatorType}">
                                        <span class="pe-option-check">✓</span>
                                        <span>Buying as an Individual Moving Home</span>
                                    </button>
                                    <button type="button" class="pe-option-button" id="${calculatorType}_first_time_buyer" data-calculator="${calculatorType}" style="display: none;">
                                        <span class="pe-option-check">✓</span>
                                        <span>First Time Buyer</span>
                                    </button>
                                    <button type="button" class="pe-option-button" id="${calculatorType}_overseas_buyer" data-calculator="${calculatorType}">
                                        <span class="pe-option-check">✓</span>
                                        <span>Overseas Buyer</span>
                                    </button>
                                </div>
                            </div>
                            <div class="pe-stamp-duty-section">
                                <label class="pe-stamp-duty-section-label">Rate Period</label>
                                <select class="pe-select" id="${calculatorType}_stamp_duty_period" data-calculator="${calculatorType}">
                                    <option value="current">Current (from 1st April 2025)</option>
                                    <option value="2024-2025">31st Oct 2024 - 31st Mar 2025</option>
                                    <option value="2022-2024">23rd Sep 2022 - 31st Oct 2024</option>
                                    <option value="2021-2022">1st Oct 2021 - 23rd Sep 2022</option>
                                    <option value="2021-mid">1st Jul 2021 - 30th Sep 2021</option>
                                    <option value="none">None</option>
                                </select>
                            </div>
                            <div class="pe-stamp-duty-breakdown" id="${calculatorType}_stamp_duty_breakdown" style="display: none;">
                                <!-- Breakdown will be populated dynamically -->
                            </div>
                        </div>
                    </div>
                    <div class="pe-field pe-field-detailed" style="display: none;">
                        <label class="pe-field-label">Survey Costs</label>
                        <input type="text" class="pe-input" id="${calculatorType}_survey_costs" data-original-id="survey_costs" data-calculator="${calculatorType}" value="£ 500">
                    </div>
                    <div class="pe-field pe-field-detailed" style="display: none;">
                        <label class="pe-field-label">Legal Fees</label>
                        <input type="text" class="pe-input" id="${calculatorType}_legal_fees" data-original-id="legal_fees" data-calculator="${calculatorType}" value="£ 1,500">
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Total Investment Required</label>
                        <div class="pe-field-with-action">
                            <input type="text" class="pe-input" id="${calculatorType}_total_investment" value="£ 0" readonly>
                            <button class="pe-btn-detail">Detail</button>
                        </div>
                    </div>
                    <a href="#" class="pe-link-add pe-field-detailed" style="display: none;">+ Add additional purchase cost</a>
                </div>
                
                <div class="pe-section">
                    <h3 class="pe-section-title">Initial Financing</h3>
                    <div class="pe-field pe-field-inline">
                        <label class="pe-field-label">Type</label>
                        <div class="pe-financing-type-group">
                            <button type="button" class="pe-financing-type" data-type="mortgage" data-calculator="${calculatorType}">
                                <span class="pe-financing-check">✓</span>
                                <span>Mortgage</span>
                            </button>
                            <button type="button" class="pe-financing-type pe-financing-type-active" data-type="bridging" data-calculator="${calculatorType}">
                                <span class="pe-financing-check">✓</span>
                                <span>Bridging Finance</span>
                            </button>
                            <button type="button" class="pe-financing-type" data-type="cash" data-calculator="${calculatorType}">
                                <span class="pe-financing-check">✓</span>
                                <span>Cash</span>
                            </button>
                            <input type="hidden" name="${calculatorType}_financing_type" value="bridging" id="${calculatorType}_financing_type_hidden">
                        </div>
                    </div>
                    <div class="pe-field pe-field-detailed pe-field-bridging" style="display: none;">
                        <label class="pe-field-label">Bridging Set-up Fee</label>
                        <div class="pe-field-with-action">
                            <input type="text" class="pe-input" id="${calculatorType}_bridging_setup_fee" data-original-id="bridging_setup_fee" data-calculator="${calculatorType}" value="1.75 %">
                            <div class="pe-toggle-group">
                                <button type="button" class="pe-toggle-small pe-toggle-small-active" data-fee-type="percent" data-fee-for="bridging" data-calculator="${calculatorType}">%</button>
                                <button type="button" class="pe-toggle-small" data-fee-type="currency" data-fee-for="bridging" data-calculator="${calculatorType}">£</button>
                            </div>
                        </div>
                    </div>
                    <div class="pe-field pe-field-detailed pe-field-bridging" style="display: none;">
                        <label class="pe-field-label">Bridging Loan to Value</label>
                        <input type="text" class="pe-input" id="${calculatorType}_bridging_ltv" data-original-id="bridging_ltv" data-calculator="${calculatorType}" value="75 %">
                    </div>
                    <div class="pe-field" id="${calculatorType}_financing_payment_field">
                        <label class="pe-field-label" id="${calculatorType}_financing_payment_label">Bridging Interest / pcm</label>
                        <input type="text" class="pe-input" id="${calculatorType}_bridging_interest" data-original-id="bridging_interest" data-calculator="${calculatorType}" value="£ 0">
                    </div>
                    <div class="pe-field pe-field-detailed pe-field-bridging" style="display: none;">
                        <label class="pe-field-label">Interest Rate (monthly)</label>
                        <input type="text" class="pe-input" id="${calculatorType}_bridging_interest_rate_monthly" data-original-id="bridging_interest_rate_monthly" data-calculator="${calculatorType}" value="1 %">
                    </div>
                    <div class="pe-field pe-field-detailed pe-field-bridging" style="display: none;">
                        <label class="pe-field-label">Finance Required</label>
                        <input type="text" class="pe-input" id="${calculatorType}_bridging_finance_required" value="£ 0" readonly>
                    </div>
                    <div class="pe-field pe-field-detailed pe-field-mortgage" style="display: none;">
                        <label class="pe-field-label">Mortgage Set-up Fee</label>
                        <div class="pe-field-with-action">
                            <input type="text" class="pe-input" id="${calculatorType}_mortgage_setup_fee" data-original-id="mortgage_setup_fee" data-calculator="${calculatorType}" value="£ 1,000">
                            <div class="pe-toggle-group">
                                <button type="button" class="pe-toggle-small pe-toggle-small-active" data-fee-type="currency" data-fee-for="mortgage" data-calculator="${calculatorType}">£</button>
                                <button type="button" class="pe-toggle-small" data-fee-type="percent" data-fee-for="mortgage" data-calculator="${calculatorType}">%</button>
                            </div>
                        </div>
                    </div>
                    <div class="pe-field pe-field-detailed pe-field-mortgage" style="display: none;">
                        <label class="pe-field-label">Mortgage Loan to Value</label>
                        <input type="text" class="pe-input" id="${calculatorType}_mortgage_ltv" data-original-id="mortgage_ltv" data-calculator="${calculatorType}" value="75 %">
                    </div>
                    <div class="pe-field pe-field-detailed pe-field-mortgage pe-field-inline" style="display: none;">
                        <label class="pe-field-label">Mortgage Type</label>
                        <div class="pe-financing-type-group">
                            <button type="button" class="pe-financing-type pe-financing-type-active" data-mortgage-type="interest_only" data-calculator="${calculatorType}">
                                <span class="pe-financing-check">✓</span>
                                <span>Interest Only</span>
                            </button>
                            <button type="button" class="pe-financing-type" data-mortgage-type="repayment" data-calculator="${calculatorType}">
                                <span class="pe-financing-check">✓</span>
                                <span>Repayment</span>
                            </button>
                        </div>
                    </div>
                    <div class="pe-field pe-field-detailed pe-field-mortgage" style="display: none;">
                        <label class="pe-field-label">Mortgage Payments / pcm</label>
                        <input type="text" class="pe-input" id="${calculatorType}_mortgage_payments_detailed" value="£ 0" readonly>
                    </div>
                    <div class="pe-field pe-field-detailed pe-field-mortgage" style="display: none;">
                        <label class="pe-field-label">Mortgage Interest Rate (APR)</label>
                        <input type="text" class="pe-input" id="${calculatorType}_mortgage_interest_rate" data-original-id="mortgage_interest_rate" data-calculator="${calculatorType}" value="5.5 %">
                    </div>
                    <div class="pe-field pe-field-detailed pe-field-mortgage pe-field-repayment" style="display: none;">
                        <label class="pe-field-label">Mortgage Term (Years)</label>
                        <input type="text" class="pe-input" id="${calculatorType}_mortgage_term_years" data-original-id="mortgage_term_years" data-calculator="${calculatorType}" value="25">
                    </div>
                    <div class="pe-field pe-field-detailed pe-field-mortgage" style="display: none;">
                        <label class="pe-field-label">Mortgage Required</label>
                        <input type="text" class="pe-input" id="${calculatorType}_mortgage_required" value="£ 0" readonly>
                    </div>
                </div>
                
                <div class="pe-section">
                    <div class="pe-section-header">
                        <h3 class="pe-section-title">Refurb</h3>
                        <label class="pe-toggle">
                            <input type="checkbox" class="pe-toggle-input" id="${calculatorType}_refurb_enabled" checked>
                            <span class="pe-toggle-slider"></span>
                        </label>
                    </div>
                    <div class="pe-section-content" id="${calculatorType}_refurb_content">
                        <div class="pe-field">
                            <label class="pe-field-label">Refurb Cost</label>
                            <input type="text" class="pe-input" id="${calculatorType}_refurb_cost" data-original-id="refurb_cost" data-calculator="${calculatorType}" value="£ 30,000">
                        </div>
                        <div class="pe-field">
                            <button type="button" class="pe-toggle-button" id="${calculatorType}_include_in_bridging" data-calculator="${calculatorType}">
                                <span class="pe-toggle-button-check">✓</span>
                                <span>Include in Bridging Finance</span>
                            </button>
                        </div>
                        <div class="pe-field">
                            <label class="pe-field-label">Vacant Period (months)</label>
                            <input type="text" class="pe-input" id="${calculatorType}_vacant_period" data-original-id="vacant_period" data-calculator="${calculatorType}" value="3">
                        </div>
                        <a href="#" class="pe-link-add pe-field-detailed" style="display: none;">+ Add additional refurb cost</a>
                        <div class="pe-subsection pe-field-detailed" style="display: none;">
                            <h4 class="pe-subsection-title">Annual Expenses During Refurb (Prorated)</h4>
                            <div class="pe-field">
                                <label class="pe-field-label">Council Tax</label>
                                <input type="text" class="pe-input" id="${calculatorType}_refurb_council_tax" data-original-id="refurb_council_tax" data-calculator="${calculatorType}" value="£ 1,670">
                            </div>
                            <div class="pe-field">
                                <label class="pe-field-label">Council Tax (monthly)</label>
                                <input type="text" class="pe-input" id="${calculatorType}_refurb_council_tax_monthly" value="£ 139.17" readonly>
                            </div>
                            <a href="#" class="pe-link-add">+ Add additional annual expense (prorated)</a>
                        </div>
                        <div class="pe-subsection pe-field-detailed" style="display: none;">
                            <h4 class="pe-subsection-title">Monthly Expenses During Refurb</h4>
                            <div class="pe-field">
                                <label class="pe-field-label">Electric / Gas</label>
                                <input type="text" class="pe-input" id="${calculatorType}_refurb_electric_gas" data-original-id="refurb_electric_gas" data-calculator="${calculatorType}" value="£ 60">
                            </div>
                            <div class="pe-field">
                                <label class="pe-field-label">Water</label>
                                <input type="text" class="pe-input" id="${calculatorType}_refurb_water" data-original-id="refurb_water" data-calculator="${calculatorType}" value="£ 30">
                            </div>
                            <div class="pe-field">
                                <label class="pe-field-label">Insurance</label>
                                <input type="text" class="pe-input" id="${calculatorType}_refurb_insurance" data-original-id="refurb_insurance" data-calculator="${calculatorType}" value="£ 40">
                            </div>
                            <a href="#" class="pe-link-add">+ Add additional monthly expense</a>
                        </div>
                    </div>
                </div>
            </div>
            
            <div class="pe-calculator-right">
                <div class="pe-section">
                    <h3 class="pe-section-title">Exit Strategy</h3>
                    <!-- Exit Strategy content can be added here if needed -->
                </div>
                
                <div class="pe-section">
                    <h3 class="pe-section-title">Refinance</h3>
                    <div class="pe-field">
                        <label class="pe-field-label required">Estimated Market Value</label>
                        <input type="text" class="pe-input" id="${calculatorType}_estimated_market_value" data-original-id="estimated_market_value" data-calculator="${calculatorType}" placeholder="Required">
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Mortgage Set-up Fee</label>
                        <div class="pe-field-with-action">
                            <input type="text" class="pe-input" id="${calculatorType}_refinance_setup_fee" data-original-id="refinance_setup_fee" data-calculator="${calculatorType}" value="£ 0">
                            <div class="pe-toggle-group">
                                <button type="button" class="pe-toggle-small pe-toggle-small-active" data-fee-type="currency" data-fee-for="refinance" data-calculator="${calculatorType}">£</button>
                                <button type="button" class="pe-toggle-small" data-fee-type="percent" data-fee-for="refinance" data-calculator="${calculatorType}">%</button>
                            </div>
                        </div>
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Mortgage Loan to Value</label>
                        <input type="text" class="pe-input" id="${calculatorType}_refinance_ltv" data-original-id="refinance_ltv" data-calculator="${calculatorType}" value="75 %">
                    </div>
                    <div class="pe-field pe-field-inline">
                        <label class="pe-field-label">Mortgage Type</label>
                        <div class="pe-financing-type-group">
                            <button type="button" class="pe-financing-type pe-financing-type-active" data-refinance-mortgage-type="interest_only" data-calculator="${calculatorType}">
                                <span class="pe-financing-check">✓</span>
                                <span>Interest Only</span>
                            </button>
                            <button type="button" class="pe-financing-type" data-refinance-mortgage-type="repayment" data-calculator="${calculatorType}">
                                <span class="pe-financing-check">✓</span>
                                <span>Repayment</span>
                            </button>
                        </div>
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Mortgage Payments / pcm</label>
                        <input type="text" class="pe-input" id="${calculatorType}_refinance_mortgage_payments" value="£ 0" readonly>
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Mortgage Interest Rate (APR)</label>
                        <input type="text" class="pe-input" id="${calculatorType}_refinance_interest_rate" data-original-id="refinance_interest_rate" data-calculator="${calculatorType}" placeholder="5.5 %">
                    </div>
                    <div class="pe-field pe-field-refinance-repayment" style="display: none;">
                        <label class="pe-field-label">Mortgage Term (Years)</label>
                        <input type="text" class="pe-input" id="${calculatorType}_refinance_mortgage_term_years" data-original-id="refinance_mortgage_term_years" data-calculator="${calculatorType}" value="25">
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Locked In Equity</label>
                        <input type="text" class="pe-input" id="${calculatorType}_locked_in_equity" value="£ 0" readonly>
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Money Left In</label>
                        <input type="text" class="pe-input" id="${calculatorType}_money_left_in" value="£ 0" readonly>
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Ideal Purchase Price (max)</label>
                        <input type="text" class="pe-input" id="${calculatorType}_ideal_purchase_price" value="£ 0" readonly>
                    </div>
                    <div id="${calculatorType}_additional_refinance_costs"></div>
                    <a href="#" class="pe-link-add" id="${calculatorType}_add_refinance_cost">+ Add additional refinance cost</a>
                </div>
                
                <div class="pe-section">
                    <h3 class="pe-section-title">Rental Income</h3>
                    <div class="pe-field">
                        <label class="pe-field-label required">Monthly Rent</label>
                        <input type="text" class="pe-input" id="${calculatorType}_monthly_rent" data-original-id="monthly_rent" data-calculator="${calculatorType}" placeholder="Required">
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Gross Yield</label>
                        <input type="text" class="pe-input" id="${calculatorType}_gross_yield" value="-" readonly>
                    </div>
                </div>
                
                <div class="pe-section pe-field-detailed" style="display: none;">
                    <h3 class="pe-section-title">Ongoing Costs</h3>
                    <div class="pe-field">
                        <label class="pe-field-label">% of Income on Maintenance</label>
                        <input type="text" class="pe-input" id="${calculatorType}_maintenance_percent" data-original-id="maintenance_percent" data-calculator="${calculatorType}" value="10 %">
                    </div>
                    <a href="#" class="pe-link-add">+ Add additional annual expense</a>
                </div>
                
                <div class="pe-section pe-field-detailed" style="display: none;">
                    <h3 class="pe-section-title">Monthly Expenses</h3>
                    <div class="pe-field">
                        <label class="pe-field-label">Mortgage Payments</label>
                        <input type="text" class="pe-input" id="${calculatorType}_ongoing_mortgage_payments" value="£ 0" readonly>
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Insurance</label>
                        <input type="text" class="pe-input" id="${calculatorType}_ongoing_insurance" data-original-id="ongoing_insurance" data-calculator="${calculatorType}" value="£ 40">
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Agent Fees</label>
                        <input type="text" class="pe-input" id="${calculatorType}_agent_fees" data-original-id="agent_fees" data-calculator="${calculatorType}" value="10 %">
                    </div>
                    <a href="#" class="pe-link-add">+ Add additional monthly expense</a>
                </div>
                
                <div class="pe-section">
                    <h3 class="pe-section-title">Summary</h3>
                    <div class="pe-field">
                        <label class="pe-field-label">Total Annual Expenses</label>
                        <input type="text" class="pe-input" id="${calculatorType}_total_annual_expenses" value="£ 0" readonly>
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Annual Profit</label>
                        <input type="text" class="pe-input" id="${calculatorType}_annual_profit" value="£ 0" readonly>
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Monthly Profit</label>
                        <input type="text" class="pe-input" id="${calculatorType}_monthly_profit" value="£ 0" readonly>
                    </div>
                    <a href="#" class="pe-link-add" id="${calculatorType}_edit_thresholds">Edit target thresholds</a>
                </div>
                
                <!-- Threshold Settings Modal -->
                <div class="pe-threshold-modal" id="${calculatorType}_threshold_modal" style="display: none;">
                    <div class="pe-threshold-modal-content">
                        <div class="pe-threshold-header">
                            <h3>Set the thresholds at which point figures will appear in green to indicate a good investment.</h3>
                        </div>
                        <div class="pe-threshold-settings">
                            <div class="pe-threshold-option-group">
                                <button type="button" class="pe-threshold-option" id="${calculatorType}_use_account_settings">
                                    <span class="pe-financing-check">✓</span>
                                    <span>Use account settings</span>
                                </button>
                                <button type="button" class="pe-threshold-option pe-threshold-option-active" id="${calculatorType}_use_calculator_settings">
                                    <span class="pe-financing-check">✓</span>
                                    <span>Use calculator settings</span>
                                </button>
                            </div>
                            <div class="pe-threshold-fields">
                                <div class="pe-field">
                                    <label class="pe-field-label">Return on capital employed</label>
                                    <input type="text" class="pe-input" id="${calculatorType}_threshold_roce" value="10 %" data-calculator="${calculatorType}">
                                </div>
                                <div class="pe-field">
                                    <label class="pe-field-label">Gross Yield</label>
                                    <input type="text" class="pe-input" id="${calculatorType}_threshold_gross_yield" value="5 %" data-calculator="${calculatorType}">
                                </div>
                                <div class="pe-field">
                                    <label class="pe-field-label">Net Yield</label>
                                    <input type="text" class="pe-input" id="${calculatorType}_threshold_net_yield" value="5 %" data-calculator="${calculatorType}">
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        
        <div class="pe-metrics">
            <div class="pe-metrics-short">
                <div class="pe-metric-box">
                    <div class="pe-metric-label">Return on capital employed</div>
                    <div class="pe-metric-value" id="${calculatorType}_roce">-%</div>
                    <span class="pe-info-icon" title="Return on capital employed">i</span>
                </div>
                <div class="pe-metric-box">
                    <div class="pe-metric-label">Gross yield</div>
                    <div class="pe-metric-value" id="${calculatorType}_gross_yield_metric">-%</div>
                    <span class="pe-info-icon" title="Gross yield">i</span>
                </div>
                <div class="pe-metric-box">
                    <div class="pe-metric-label">Net yield</div>
                    <div class="pe-metric-value" id="${calculatorType}_net_yield">-%</div>
                    <span class="pe-info-icon" title="Net yield">i</span>
                </div>
            </div>
            <div class="pe-metrics-long">
                <div class="pe-metric-box pe-metric-box-blue">
                    <div class="pe-metric-label">Equity in 10 years</div>
                    <div class="pe-metric-value" id="${calculatorType}_equity_10_years">£0</div>
                    <span class="pe-info-icon" title="Projected equity after 10 years">i</span>
                </div>
            </div>
            <div class="pe-metrics-appreciation">
                <label class="pe-field-label">Annual Property Appreciation</label>
                <input type="text" class="pe-input pe-input-small" id="${calculatorType}_appreciation" value="5.5%" data-calculator="${calculatorType}">
            </div>
        </div>
    `;
    
    // Verify element exists immediately after innerHTML
    const testToggle = section.querySelector(`#${calculatorType}_detailed_view`);
    console.log('[Frontend] Element check immediately after innerHTML:', {
        calculatorType,
        toggleId: `${calculatorType}_detailed_view`,
        found: !!testToggle,
        sectionHasContent: section.innerHTML.length > 0,
        sectionChildren: section.children.length
    });
    
    // Add event listeners
    console.log('[Frontend] Setting up event handlers for calculator:', calculatorType);
    setupBRRCalculatorEvents(section, calculatorType);
    console.log('[Frontend] Event handlers set up for calculator:', calculatorType);
    
    // Ensure critical fields have direct event listeners as a backup
    // Do this after a short delay to ensure DOM is ready
    setTimeout(() => {
        const refinanceInterestRateInput = document.getElementById(`${calculatorType}_refinance_interest_rate`);
        if (refinanceInterestRateInput) {
            // Add direct event listeners (multiple events to catch all cases)
            const triggerCalc = () => calculateBRRValues(calculatorType);
            refinanceInterestRateInput.addEventListener('input', triggerCalc, { passive: true });
            refinanceInterestRateInput.addEventListener('change', triggerCalc);
            refinanceInterestRateInput.addEventListener('keyup', triggerCalc);
            refinanceInterestRateInput.addEventListener('blur', triggerCalc);
        }
    }, 200);
    
    return section;
}

// Create Holiday Let calculator (similar structure to BRR)
function createHolidayLetCalculator(calculatorType) {
    console.log('[Frontend] createHolidayLetCalculator called for:', calculatorType);
    
    // Create section similar to BRR but with Holiday Let-specific fields
    const section = document.createElement('div');
    section.className = 'calculator-fields-section propertyengine-calculator';
    section.dataset.calculator = calculatorType;
    
    // Get the BRR HTML and modify it for Holiday Let
    const brrSection = createBRRCalculator(calculatorType);
    section.innerHTML = brrSection.innerHTML;
    
    // Replace "Rental Income" section with "Income" section for Holiday Let
    // Find all section titles and look for "Rental Income"
    const allSectionTitles = section.querySelectorAll('h3.pe-section-title');
    let rentalIncomeSection = null;
    let rentalIncomeSectionContainer = null;
    
    allSectionTitles.forEach(title => {
        if (title.textContent.trim() === 'Rental Income') {
            rentalIncomeSection = title;
            rentalIncomeSectionContainer = title.closest('.pe-section');
        }
    });
    
    if (rentalIncomeSection && rentalIncomeSectionContainer) {
        console.log('[Frontend] Found Rental Income section, replacing with Income section');
        rentalIncomeSection.textContent = 'Income';
        
        // Find the Monthly Rent field and replace with Nightly Rate and Occupancy Rate
        const monthlyRentField = rentalIncomeSectionContainer.querySelector(`#${calculatorType}_monthly_rent`)?.closest('.pe-field');
        const grossYieldField = rentalIncomeSectionContainer.querySelector(`#${calculatorType}_gross_yield`)?.closest('.pe-field');
        
        // Remove Gross Yield field if it exists
        if (grossYieldField) {
            console.log('[Frontend] Removing Gross Yield field from Income section');
            grossYieldField.remove();
        }
        
        if (monthlyRentField) {
            console.log('[Frontend] Found Monthly Rent field, replacing with Nightly Rate');
            monthlyRentField.innerHTML = `
                <label class="pe-field-label required">Nightly Rate</label>
                <input type="text" class="pe-input" id="${calculatorType}_nightly_rate" data-original-id="nightly_rate" data-calculator="${calculatorType}" placeholder="£ 0">
            `;
            
            // Add Occupancy Rate field after Nightly Rate
            const occupancyField = document.createElement('div');
            occupancyField.className = 'pe-field';
            occupancyField.innerHTML = `
                <label class="pe-field-label">Annualised Occupancy Rate</label>
                <input type="text" class="pe-input" id="${calculatorType}_occupancy_rate" data-original-id="occupancy_rate" data-calculator="${calculatorType}" value="70 %">
            `;
            
            // Insert after Nightly Rate field
            monthlyRentField.parentNode.insertBefore(occupancyField, monthlyRentField.nextSibling);
        } else {
            console.warn('[Frontend] Monthly Rent field not found in Rental Income section');
        }
    } else {
        console.warn('[Frontend] Rental Income section not found. Available sections:', 
            Array.from(section.querySelectorAll('h3.pe-section-title')).map(t => t.textContent.trim()));
    }
    
    // Update Ongoing Costs section for Holiday Let
    // Find all sections and look for Ongoing Costs (might be in detailed view)
    const allSections = section.querySelectorAll('.pe-section');
    let ongoingCostsSection = null;
    let monthlyExpensesSection = null;
    
    allSections.forEach(sec => {
        const title = sec.querySelector('h3.pe-section-title');
        if (title) {
            if (title.textContent === 'Ongoing Costs') {
                ongoingCostsSection = sec;
            } else if (title.textContent === 'Monthly Expenses') {
                monthlyExpensesSection = sec;
            }
        }
    });
    
    // If Ongoing Costs exists, replace it. Otherwise create new section
    if (ongoingCostsSection) {
        // Remove the 'pe-field-detailed' class to make it always visible
        ongoingCostsSection.classList.remove('pe-field-detailed');
        ongoingCostsSection.style.display = 'block';
        
        // Clear existing content but keep the title
        const title = ongoingCostsSection.querySelector('h3.pe-section-title');
        
        // Create Annual Expenses subsection
        const annualExpensesSubsection = document.createElement('div');
        annualExpensesSubsection.className = 'pe-subsection';
        annualExpensesSubsection.innerHTML = `
            <h4 class="pe-subsection-title">Annual Expenses</h4>
            <div class="pe-field">
                <label class="pe-field-label">Council Tax</label>
                <input type="text" class="pe-input" id="${calculatorType}_council_tax" data-original-id="council_tax" data-calculator="${calculatorType}" value="£ 1,670">
            </div>
            <div class="pe-field">
                <label class="pe-field-label">% of Income on Maintenance</label>
                <input type="text" class="pe-input" id="${calculatorType}_maintenance_percent" data-original-id="maintenance_percent" data-calculator="${calculatorType}" value="10 %">
            </div>
            <div class="pe-field">
                <label class="pe-field-label" style="display: flex; align-items: center; gap: 8px; cursor: pointer;">
                    <input type="checkbox" id="${calculatorType}_tv_license" data-original-id="tv_license" data-calculator="${calculatorType}" style="width: auto; margin: 0; cursor: pointer;">
                    <span>TV License</span>
                </label>
            </div>
            <a href="#" class="pe-link-add">+ Add additional annual expense</a>
        `;
        
        // Create Monthly Expenses subsection
        const monthlyExpensesSubsection = document.createElement('div');
        monthlyExpensesSubsection.className = 'pe-subsection';
        monthlyExpensesSubsection.innerHTML = `
            <h4 class="pe-subsection-title">Monthly Expenses</h4>
            <div class="pe-field">
                <label class="pe-field-label">Mortgage Payments</label>
                <input type="text" class="pe-input" id="${calculatorType}_ongoing_mortgage_payments" value="£ 0" readonly>
            </div>
            <div class="pe-field">
                <label class="pe-field-label">Electric / Gas</label>
                <input type="text" class="pe-input" id="${calculatorType}_utilities" data-original-id="utilities" data-calculator="${calculatorType}" value="£ 140">
            </div>
            <div class="pe-field">
                <label class="pe-field-label">Water</label>
                <input type="text" class="pe-input" id="${calculatorType}_water" data-original-id="water" data-calculator="${calculatorType}" value="£ 40">
            </div>
            <div class="pe-field">
                <label class="pe-field-label">Broadband / TV</label>
                <input type="text" class="pe-input" id="${calculatorType}_broadband_tv" data-original-id="broadband_tv" data-calculator="${calculatorType}" value="£ 60">
            </div>
            <div class="pe-field">
                <label class="pe-field-label">Insurance</label>
                <input type="text" class="pe-input" id="${calculatorType}_ongoing_insurance" data-original-id="ongoing_insurance" data-calculator="${calculatorType}" value="£ 40">
            </div>
            <div class="pe-field">
                <label class="pe-field-label">Agent Fees</label>
                <div class="pe-field-with-action">
                    <input type="text" class="pe-input" id="${calculatorType}_agent_fees" data-original-id="agent_fees" data-calculator="${calculatorType}" value="0 %">
                    <div class="pe-toggle-group">
                        <button type="button" class="pe-toggle-small" data-fee-type="currency" data-fee-for="agent" data-calculator="${calculatorType}">£</button>
                        <button type="button" class="pe-toggle-small pe-toggle-small-active" data-fee-type="percent" data-fee-for="agent" data-calculator="${calculatorType}">%</button>
                    </div>
                </div>
            </div>
            <div class="pe-field">
                <label class="pe-field-label">Booking Fees</label>
                <div class="pe-field-with-action">
                    <input type="text" class="pe-input" id="${calculatorType}_booking_fees" data-original-id="booking_fees" data-calculator="${calculatorType}" value="12 %">
                    <div class="pe-toggle-group">
                        <button type="button" class="pe-toggle-small" data-fee-type="currency" data-fee-for="booking" data-calculator="${calculatorType}">£</button>
                        <button type="button" class="pe-toggle-small pe-toggle-small-active" data-fee-type="percent" data-fee-for="booking" data-calculator="${calculatorType}">%</button>
                    </div>
                </div>
            </div>
            <div class="pe-field">
                <label class="pe-field-label">Cleaning Costs</label>
                <input type="text" class="pe-input" id="${calculatorType}_cleaning_costs" data-original-id="cleaning_costs" data-calculator="${calculatorType}" value="£ 80">
            </div>
            <a href="#" class="pe-link-add">+ Add additional monthly expense</a>
        `;
        
        // Replace content - clear all children except title
        Array.from(ongoingCostsSection.children).forEach(child => {
            if (child.tagName !== 'H3') {
                child.remove();
            }
        });
        
        // Add subsections
        ongoingCostsSection.appendChild(annualExpensesSubsection);
        ongoingCostsSection.appendChild(monthlyExpensesSubsection);
    }
    
    // Remove the separate Monthly Expenses section if it exists (we've merged it into Ongoing Costs)
    if (monthlyExpensesSection) {
        monthlyExpensesSection.remove();
    }
    
    // Update Exit Strategy to be blank for Holiday Let
    const exitStrategySection = Array.from(section.querySelectorAll('.pe-section')).find(sec => {
        const title = sec.querySelector('h3.pe-section-title');
        return title && title.textContent === 'Exit Strategy';
    });
    
    if (exitStrategySection) {
        exitStrategySection.innerHTML = `
            <h3 class="pe-section-title">Exit Strategy</h3>
        `;
    }
    
    // Set up event handlers (reuse BRR event handlers)
    setupBRRCalculatorEvents(section, calculatorType);
    
    // Add specific event listeners for Holiday Let-specific fields
    setTimeout(() => {
        // TV License checkbox
        const tvLicenseCheckbox = document.getElementById(`${calculatorType}_tv_license`);
        if (tvLicenseCheckbox) {
            tvLicenseCheckbox.addEventListener('change', () => {
                calculateBRRValues(calculatorType);
            });
        }
        
        // Nightly Rate and Occupancy Rate fields (ensure they trigger calculations)
        const nightlyRateInput = document.getElementById(`${calculatorType}_nightly_rate`);
        const occupancyRateInput = document.getElementById(`${calculatorType}_occupancy_rate`);
        
        if (nightlyRateInput) {
            nightlyRateInput.addEventListener('input', () => calculateBRRValues(calculatorType));
            nightlyRateInput.addEventListener('change', () => calculateBRRValues(calculatorType));
        }
        
        if (occupancyRateInput) {
            occupancyRateInput.addEventListener('input', () => calculateBRRValues(calculatorType));
            occupancyRateInput.addEventListener('change', () => calculateBRRValues(calculatorType));
        }
    }, 100);
    
    console.log('[Frontend] Holiday Let section created with custom fields, section exists:', !!section);
    return section;
}

// Create Rent to HMO calculator
function createRentToHMOCalculator(calculatorType) {
    console.log('[Frontend] createRentToHMOCalculator called for:', calculatorType);
    const section = document.createElement('div');
    section.className = 'calculator-fields-section propertyengine-calculator';
    section.dataset.calculator = calculatorType;
    
    section.innerHTML = `
        <div class="pe-calculator-topbar">
            <h2 class="pe-calculator-title">Rent to HMO</h2>
            <div class="pe-calculator-actions">
                <label class="pe-switch-label">
                    <input type="checkbox" class="pe-switch-detailed" id="${calculatorType}_detailed_view">
                    <span>Switch to detailed view</span>
                </label>
                <button class="pe-btn-save">Save Calculator</button>
                <div class="pe-roi-display">
                    <div class="pe-roi-label">Return on Investment</div>
                    <div class="pe-roi-value" id="${calculatorType}_roi">-%</div>
                </div>
            </div>
        </div>
        
        <div class="pe-calculator-content">
            <div class="pe-calculator-left">
                <div class="pe-section">
                    <h3 class="pe-section-title">Acquisition</h3>
                    <div class="pe-field">
                        <label class="pe-field-label">Deposit</label>
                        <input type="text" class="pe-input" id="${calculatorType}_deposit" data-original-id="deposit" data-calculator="${calculatorType}" value="£ 0">
                    </div>
                    <div class="pe-field pe-field-detailed" style="display: none;">
                        <label class="pe-field-label">Survey Costs</label>
                        <input type="text" class="pe-input" id="${calculatorType}_survey_costs" data-original-id="survey_costs" data-calculator="${calculatorType}" value="£ 0">
                    </div>
                    <div class="pe-field pe-field-detailed" style="display: none;">
                        <label class="pe-field-label">Legal Fees</label>
                        <input type="text" class="pe-input" id="${calculatorType}_legal_fees" data-original-id="legal_fees" data-calculator="${calculatorType}" value="£ 0">
                    </div>
                    <div class="pe-field pe-field-detailed" style="display: none;">
                        <label class="pe-field-label">Reference Fees</label>
                        <input type="text" class="pe-input" id="${calculatorType}_reference_fees" data-original-id="reference_fees" data-calculator="${calculatorType}" value="£ 0">
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Total Investment Required</label>
                        <div class="pe-field-with-action">
                            <input type="text" class="pe-input" id="${calculatorType}_total_investment" value="£ 0" readonly>
                            <button class="pe-btn-detail">Detail</button>
                        </div>
                    </div>
                    <a href="#" class="pe-link-add pe-field-detailed" style="display: none;" id="${calculatorType}_add_acquisition_cost">Add additional acquisition cost</a>
                </div>
                
                <div class="pe-section">
                    <div class="pe-section-header">
                        <h3 class="pe-section-title">Refurb</h3>
                        <label class="pe-toggle">
                            <input type="checkbox" class="pe-toggle-input" id="${calculatorType}_refurb_enabled" checked>
                            <span class="pe-toggle-slider"></span>
                        </label>
                    </div>
                    <div class="pe-section-content" id="${calculatorType}_refurb_content">
                        <div class="pe-field">
                            <label class="pe-field-label">Refurb Cost</label>
                            <input type="text" class="pe-input" id="${calculatorType}_refurb_cost" data-original-id="refurb_cost" data-calculator="${calculatorType}" value="£ 0">
                        </div>
                        <div class="pe-field">
                            <label class="pe-field-label">Vacant Period (months)</label>
                            <input type="text" class="pe-input" id="${calculatorType}_vacant_period" data-original-id="vacant_period" data-calculator="${calculatorType}" value="1">
                        </div>
                        <a href="#" class="pe-link-add" id="${calculatorType}_add_refurb_cost">Add additional refurb cost</a>
                        <div class="pe-subsection">
                            <h4 class="pe-subsection-title">Annual Expenses During Refurb (Prorated)</h4>
                            <div class="pe-field">
                                <label class="pe-field-label">Council Tax</label>
                                <input type="text" class="pe-input" id="${calculatorType}_refurb_council_tax" data-original-id="refurb_council_tax" data-calculator="${calculatorType}" value="£ 1,670">
                            </div>
                            <div class="pe-field">
                                <label class="pe-field-label">Council Tax (monthly)</label>
                                <input type="text" class="pe-input" id="${calculatorType}_refurb_council_tax_monthly" value="£ 139.17" readonly>
                            </div>
                            <a href="#" class="pe-link-add" id="${calculatorType}_add_refurb_annual_expense">Add additional annual expense (prorated)</a>
                        </div>
                        <div class="pe-subsection">
                            <h4 class="pe-subsection-title">Monthly Expenses During Refurb</h4>
                            <div class="pe-field">
                                <label class="pe-field-label">Electric / Gas</label>
                                <input type="text" class="pe-input" id="${calculatorType}_refurb_electric_gas" data-original-id="refurb_electric_gas" data-calculator="${calculatorType}" value="£ 60">
                            </div>
                            <div class="pe-field">
                                <label class="pe-field-label">Water</label>
                                <input type="text" class="pe-input" id="${calculatorType}_refurb_water" data-original-id="refurb_water" data-calculator="${calculatorType}" value="£ 30">
                            </div>
                            <div class="pe-field">
                                <label class="pe-field-label">Insurance</label>
                                <input type="text" class="pe-input" id="${calculatorType}_refurb_insurance" data-original-id="refurb_insurance" data-calculator="${calculatorType}" value="£ 40">
                            </div>
                            <a href="#" class="pe-link-add" id="${calculatorType}_add_refurb_monthly_expense">Add additional monthly expense</a>
                        </div>
                    </div>
                </div>
                
                <div class="pe-section">
                    <h3 class="pe-section-title">Rent to Owner</h3>
                    <div class="pe-field">
                        <label class="pe-field-label required">Monthly Rent to Owner</label>
                        <input type="text" class="pe-input" id="${calculatorType}_monthly_rent_to_owner" data-original-id="monthly_rent_to_owner" data-calculator="${calculatorType}" placeholder="Required">
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Chosen Strategy</label>
                        <select class="pe-select" id="${calculatorType}_chosen_strategy" data-calculator="${calculatorType}">
                            <option value="holiday-let">Holiday Let</option>
                            <option value="hmo" selected>HMO</option>
                        </select>
                    </div>
                </div>
            </div>
            
            <div class="pe-calculator-right">
                <div class="pe-section" id="${calculatorType}_rental_income_section">
                    <h3 class="pe-section-title" id="${calculatorType}_rental_income_title">Rental Income</h3>
                    <div class="pe-field">
                        <label class="pe-field-label">Total</label>
                        <input type="text" class="pe-input" id="${calculatorType}_total_rental_income" value="£0 / pcm" readonly>
                    </div>
                    <div id="${calculatorType}_rooms_container">
                        <div class="pe-field">
                            <label class="pe-field-label">Room 1 Monthly Rent</label>
                            <input type="text" class="pe-input" id="${calculatorType}_room_1_rent" data-original-id="room_1_rent" data-calculator="${calculatorType}" data-room="1" value="£ 450">
                        </div>
                        <div class="pe-field">
                            <label class="pe-field-label">Room 2 Monthly Rent</label>
                            <input type="text" class="pe-input" id="${calculatorType}_room_2_rent" data-original-id="room_2_rent" data-calculator="${calculatorType}" data-room="2" value="£ 450">
                        </div>
                        <div class="pe-field">
                            <label class="pe-field-label">Room 3 Monthly Rent</label>
                            <input type="text" class="pe-input" id="${calculatorType}_room_3_rent" data-original-id="room_3_rent" data-calculator="${calculatorType}" data-room="3" value="£ 400">
                        </div>
                        <div class="pe-field">
                            <label class="pe-field-label">Room 4 Monthly Rent</label>
                            <input type="text" class="pe-input" id="${calculatorType}_room_4_rent" data-original-id="room_4_rent" data-calculator="${calculatorType}" data-room="4" value="£ 400">
                        </div>
                    </div>
                    <a href="#" class="pe-link-add" id="${calculatorType}_add_room">Add room</a>
                </div>
                
                <!-- Holiday Let Income Section (hidden by default) -->
                <div class="pe-section" id="${calculatorType}_holiday_let_income_section" style="display: none;">
                    <h3 class="pe-section-title">Income</h3>
                    <div class="pe-field">
                        <label class="pe-field-label required">Nightly Rate</label>
                        <input type="text" class="pe-input" id="${calculatorType}_nightly_rate" data-original-id="nightly_rate" data-calculator="${calculatorType}" placeholder="Required">
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Annualised Occupancy Rate</label>
                        <input type="text" class="pe-input" id="${calculatorType}_occupancy_rate" data-original-id="occupancy_rate" data-calculator="${calculatorType}" value="70 %">
                    </div>
                </div>
                
                <div class="pe-section pe-field-detailed" style="display: none;">
                    <h3 class="pe-section-title">Ongoing Costs</h3>
                    <div class="pe-subsection">
                        <h4 class="pe-subsection-title">Annual Expenses</h4>
                        <div class="pe-field">
                            <label class="pe-field-label">Council Tax</label>
                            <input type="text" class="pe-input" id="${calculatorType}_council_tax" data-original-id="council_tax" data-calculator="${calculatorType}" value="£ 1,670">
                        </div>
                        <div class="pe-field">
                            <label class="pe-field-label">% of Income on Maintenance</label>
                            <input type="text" class="pe-input" id="${calculatorType}_maintenance_percent" data-original-id="maintenance_percent" data-calculator="${calculatorType}" value="2 %">
                        </div>
                        <div class="pe-field" id="${calculatorType}_tv_license_field">
                            <label class="pe-field-label" style="display: flex; align-items: center; gap: 8px; cursor: pointer;">
                                <input type="checkbox" id="${calculatorType}_communal_tv_license" data-original-id="communal_tv_license" data-calculator="${calculatorType}" style="width: auto; margin: 0; cursor: pointer;">
                                <span id="${calculatorType}_tv_license_label">Communal TV License</span>
                            </label>
                        </div>
                        <a href="#" class="pe-link-add" id="${calculatorType}_add_annual_expense">Add additional annual expense</a>
                    </div>
                    <div class="pe-subsection">
                        <h4 class="pe-subsection-title">Monthly Expenses</h4>
                        <div class="pe-field">
                            <label class="pe-field-label">Electric / Gas</label>
                            <input type="text" class="pe-input" id="${calculatorType}_utilities" data-original-id="utilities" data-calculator="${calculatorType}" value="£ 165">
                        </div>
                        <div class="pe-field">
                            <label class="pe-field-label">Water</label>
                            <input type="text" class="pe-input" id="${calculatorType}_water" data-original-id="water" data-calculator="${calculatorType}" value="£ 40">
                        </div>
                        <div class="pe-field">
                            <label class="pe-field-label">Broadband / TV</label>
                            <input type="text" class="pe-input" id="${calculatorType}_broadband_tv" data-original-id="broadband_tv" data-calculator="${calculatorType}" value="£ 60">
                        </div>
                        <div class="pe-field">
                            <label class="pe-field-label">Insurance</label>
                            <input type="text" class="pe-input" id="${calculatorType}_insurance" data-original-id="insurance" data-calculator="${calculatorType}" value="£ 40">
                        </div>
                        <div class="pe-field" id="${calculatorType}_monthly_rent_to_owner_field">
                            <label class="pe-field-label">Monthly Rent to Owner</label>
                            <input type="text" class="pe-input" id="${calculatorType}_ongoing_monthly_rent_to_owner" value="£ 0" readonly>
                        </div>
                        <div class="pe-field">
                            <label class="pe-field-label">Agent Fees</label>
                            <div class="pe-field-with-action">
                                <input type="text" class="pe-input" id="${calculatorType}_agent_fees" data-original-id="agent_fees" data-calculator="${calculatorType}" data-fee-type="percent" value="15 %">
                                <div class="pe-toggle-group">
                                    <button type="button" class="pe-toggle-small" data-fee-type="currency" data-fee-for="agent" data-calculator="${calculatorType}">£</button>
                                    <button type="button" class="pe-toggle-small pe-toggle-small-active" data-fee-type="percent" data-fee-for="agent" data-calculator="${calculatorType}">%</button>
                                </div>
                            </div>
                        </div>
                        <div class="pe-field" id="${calculatorType}_booking_fees_field" style="display: none;">
                            <label class="pe-field-label">Booking Fees</label>
                            <div class="pe-field-with-action">
                                <input type="text" class="pe-input" id="${calculatorType}_booking_fees" data-original-id="booking_fees" data-calculator="${calculatorType}" data-fee-type="percent" value="12 %">
                                <div class="pe-toggle-group">
                                    <button type="button" class="pe-toggle-small" data-fee-type="currency" data-fee-for="booking" data-calculator="${calculatorType}">£</button>
                                    <button type="button" class="pe-toggle-small pe-toggle-small-active" data-fee-type="percent" data-fee-for="booking" data-calculator="${calculatorType}">%</button>
                                </div>
                            </div>
                        </div>
                        <div class="pe-field">
                            <label class="pe-field-label">Cleaning Costs</label>
                            <input type="text" class="pe-input" id="${calculatorType}_cleaning_costs" data-original-id="cleaning_costs" data-calculator="${calculatorType}" value="£ 80">
                        </div>
                        <a href="#" class="pe-link-add" id="${calculatorType}_add_monthly_expense">Add additional monthly expense</a>
                    </div>
                </div>
                
                <div class="pe-section">
                    <h3 class="pe-section-title">Summary</h3>
                    <div class="pe-field">
                        <label class="pe-field-label">Total Annual Expenses</label>
                        <div class="pe-field-with-action">
                            <input type="text" class="pe-input" id="${calculatorType}_total_annual_expenses" value="£ 0" readonly>
                            <button class="pe-btn-detail">Detail</button>
                        </div>
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Annual Profit</label>
                        <input type="text" class="pe-input" id="${calculatorType}_annual_profit" value="£ 0" readonly>
                    </div>
                    <div class="pe-field">
                        <label class="pe-field-label">Monthly Profit</label>
                        <input type="text" class="pe-input" id="${calculatorType}_monthly_profit" value="£ 0" readonly>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    // Set up event handlers
    setupRentToHMOCalculatorEvents(section, calculatorType);
    
    return section;
}

function setupRentToHMOCalculatorEvents(section, calculatorType) {
    console.log('[Frontend] setupRentToHMOCalculatorEvents called for:', calculatorType);
    
    // Detailed view toggle
    const detailedViewToggle = section.querySelector(`#${calculatorType}_detailed_view`);
    if (detailedViewToggle) {
        detailedViewToggle.addEventListener('change', (e) => {
            const detailedFields = section.querySelectorAll('.pe-field-detailed');
            detailedFields.forEach(field => {
                field.style.display = e.target.checked ? 'block' : 'none';
            });
            
            // Also show/hide the Ongoing Costs section
            const ongoingCostsSection = section.querySelector('.pe-section.pe-field-detailed');
            if (ongoingCostsSection) {
                ongoingCostsSection.style.display = e.target.checked ? 'block' : 'none';
            }
            
            const label = detailedViewToggle.parentElement.querySelector('span');
            if (label) {
                label.textContent = e.target.checked ? 'Switch to simple view' : 'Switch to detailed view';
            }
            
            // Recalculate when toggling view
            calculateRentToHMOValues(calculatorType);
        });
    }
    
    // Refurb toggle
    const refurbToggle = section.querySelector(`#${calculatorType}_refurb_enabled`);
    if (refurbToggle) {
        refurbToggle.addEventListener('change', (e) => {
            const content = section.querySelector(`#${calculatorType}_refurb_content`);
            if (content) {
                content.style.display = e.target.checked ? 'block' : 'none';
            }
            calculateRentToHMOValues(calculatorType);
        });
    }
    
    // Refurb council tax - calculate monthly when annual changes
    const refurbCouncilTaxInput = section.querySelector(`#${calculatorType}_refurb_council_tax`);
    const refurbCouncilTaxMonthlyInput = section.querySelector(`#${calculatorType}_refurb_council_tax_monthly`);
    if (refurbCouncilTaxInput && refurbCouncilTaxMonthlyInput) {
        const updateRefurbCouncilTaxMonthly = () => {
            const annualValue = parseCurrency(refurbCouncilTaxInput.value || '0');
            const monthlyValue = annualValue / 12;
            refurbCouncilTaxMonthlyInput.value = formatCurrency(monthlyValue);
            calculateRentToHMOValues(calculatorType);
        };
        refurbCouncilTaxInput.addEventListener('input', updateRefurbCouncilTaxMonthly);
        refurbCouncilTaxInput.addEventListener('change', updateRefurbCouncilTaxMonthly);
    }
    
    // Vacant period - update calculations when changed
    const vacantPeriodInput = section.querySelector(`#${calculatorType}_vacant_period`);
    if (vacantPeriodInput) {
        vacantPeriodInput.addEventListener('input', () => calculateRentToHMOValues(calculatorType));
        vacantPeriodInput.addEventListener('change', () => calculateRentToHMOValues(calculatorType));
    }
    
    // Refurb expense inputs - update calculations when changed
    const refurbExpenseInputs = section.querySelectorAll(`#${calculatorType}_refurb_electric_gas, #${calculatorType}_refurb_water, #${calculatorType}_refurb_insurance`);
    refurbExpenseInputs.forEach(input => {
        input.addEventListener('input', () => calculateRentToHMOValues(calculatorType));
        input.addEventListener('change', () => calculateRentToHMOValues(calculatorType));
    });
    
    // Add room button
    const addRoomBtn = section.querySelector(`#${calculatorType}_add_room`);
    if (addRoomBtn) {
        addRoomBtn.addEventListener('click', (e) => {
            e.preventDefault();
            const roomsContainer = section.querySelector(`#${calculatorType}_rooms_container`);
            if (roomsContainer) {
                const existingRooms = roomsContainer.querySelectorAll('.pe-field');
                const nextRoomNumber = existingRooms.length + 1;
                const newRoomField = document.createElement('div');
                newRoomField.className = 'pe-field';
                newRoomField.innerHTML = `
                    <label class="pe-field-label">Room ${nextRoomNumber} Monthly Rent</label>
                    <input type="text" class="pe-input" id="${calculatorType}_room_${nextRoomNumber}_rent" data-original-id="room_${nextRoomNumber}_rent" data-calculator="${calculatorType}" data-room="${nextRoomNumber}" value="£ 400">
                `;
                roomsContainer.appendChild(newRoomField);
                
                // Add event listener to new room input
                const newRoomInput = newRoomField.querySelector('input');
                if (newRoomInput) {
                    newRoomInput.addEventListener('input', () => calculateRentToHMOValues(calculatorType));
                    newRoomInput.addEventListener('change', () => calculateRentToHMOValues(calculatorType));
                }
            }
        });
    }
    
    // Agent fees toggle buttons
    const agentFeeToggles = section.querySelectorAll(`[data-fee-for="agent"][data-calculator="${calculatorType}"]`);
    agentFeeToggles.forEach(toggle => {
        toggle.addEventListener('click', () => {
            // Toggle active state
            agentFeeToggles.forEach(t => t.classList.remove('pe-toggle-small-active'));
            toggle.classList.add('pe-toggle-small-active');
            
            // Update input data attribute based on selected type
            const agentFeeInput = section.querySelector(`#${calculatorType}_agent_fees`);
            if (agentFeeInput) {
                const feeType = toggle.dataset.feeType;
                // Store the fee type for calculation
                agentFeeInput.dataset.feeType = feeType;
                
                // Update the value format hint if needed
                const currentValue = agentFeeInput.value || '';
                if (feeType === 'percent' && !currentValue.includes('%')) {
                    // If switching to percent and value doesn't have %, add it
                    const numValue = parseFloat(currentValue.replace(/[£,\s]/g, '')) || 0;
                    if (numValue > 0) {
                        agentFeeInput.value = `${numValue} %`;
                    }
                } else if (feeType === 'currency' && currentValue.includes('%')) {
                    // If switching to currency and value has %, remove it
                    const numValue = parseFloat(currentValue.replace(/[%\s]/g, '')) || 0;
                    if (numValue > 0) {
                        agentFeeInput.value = formatCurrency(numValue);
                    }
                }
            }
            
            calculateRentToHMOValues(calculatorType);
        });
    });
    
    // Strategy dropdown - switch between HMO and Holiday Let
    const strategySelect = section.querySelector(`#${calculatorType}_chosen_strategy`);
    if (strategySelect) {
        strategySelect.addEventListener('change', (e) => {
            const strategy = e.target.value;
            switchStrategy(section, calculatorType, strategy);
            calculateRentToHMOValues(calculatorType);
        });
    }
    
    // TV License checkbox
    const tvLicenseCheckbox = section.querySelector(`#${calculatorType}_communal_tv_license`);
    if (tvLicenseCheckbox) {
        tvLicenseCheckbox.addEventListener('change', () => {
            calculateRentToHMOValues(calculatorType);
        });
    }
    
    // Add acquisition cost button
    const addAcquisitionCostBtn = section.querySelector(`#${calculatorType}_add_acquisition_cost`);
    if (addAcquisitionCostBtn) {
        addAcquisitionCostBtn.addEventListener('click', (e) => {
            e.preventDefault();
            // Implementation for adding additional acquisition costs
            // Similar to BRR calculator pattern
        });
    }
    
    // Add annual expense button
    const addAnnualExpenseBtn = section.querySelector(`#${calculatorType}_add_annual_expense`);
    if (addAnnualExpenseBtn) {
        addAnnualExpenseBtn.addEventListener('click', (e) => {
            e.preventDefault();
            // Implementation for adding additional annual expenses
        });
    }
    
    // Add monthly expense button
    const addMonthlyExpenseBtn = section.querySelector(`#${calculatorType}_add_monthly_expense`);
    if (addMonthlyExpenseBtn) {
        addMonthlyExpenseBtn.addEventListener('click', (e) => {
            e.preventDefault();
            // Implementation for adding additional monthly expenses
        });
    }
    
    // Add event listeners to all inputs
    const allInputs = section.querySelectorAll('input, select');
    allInputs.forEach(input => {
        if (input.id && !input.readOnly && input.type !== 'checkbox') {
            input.addEventListener('input', () => calculateRentToHMOValues(calculatorType));
            input.addEventListener('change', () => calculateRentToHMOValues(calculatorType));
        }
    });
}

function switchStrategy(section, calculatorType, strategy) {
    console.log('[Frontend] Switching strategy to:', strategy);
    
    const rentalIncomeSection = section.querySelector(`#${calculatorType}_rental_income_section`);
    const holidayLetIncomeSection = section.querySelector(`#${calculatorType}_holiday_let_income_section`);
    const monthlyRentToOwnerField = section.querySelector(`#${calculatorType}_monthly_rent_to_owner_field`);
    const bookingFeesField = section.querySelector(`#${calculatorType}_booking_fees_field`);
    const tvLicenseField = section.querySelector(`#${calculatorType}_tv_license_field`);
    const tvLicenseLabel = section.querySelector(`#${calculatorType}_tv_license_label`);
    const agentFeesInput = section.querySelector(`#${calculatorType}_agent_fees`);
    const utilitiesInput = section.querySelector(`#${calculatorType}_utilities`);
    const maintenancePercentInput = section.querySelector(`#${calculatorType}_maintenance_percent`);
    
    if (strategy === 'holiday-let') {
        // Show Holiday Let income section, hide HMO rental income
        if (rentalIncomeSection) rentalIncomeSection.style.display = 'none';
        if (holidayLetIncomeSection) holidayLetIncomeSection.style.display = 'block';
        
        // Show Monthly Rent to Owner in Ongoing Costs (it should be visible for both strategies)
        if (monthlyRentToOwnerField) monthlyRentToOwnerField.style.display = 'block';
        
        // Show Booking Fees
        if (bookingFeesField) bookingFeesField.style.display = 'block';
        
        // Change TV License label
        if (tvLicenseLabel) tvLicenseLabel.textContent = 'TV License';
        
        // Update defaults for Holiday Let
        if (agentFeesInput) {
            agentFeesInput.value = '0 %';
            // Set active toggle to percent
            const percentToggle = section.querySelector('[data-fee-for="agent"][data-fee-type="percent"]');
            const currencyToggle = section.querySelector('[data-fee-for="agent"][data-fee-type="currency"]');
            if (percentToggle && currencyToggle) {
                currencyToggle.classList.remove('pe-toggle-small-active');
                percentToggle.classList.add('pe-toggle-small-active');
                agentFeesInput.dataset.feeType = 'percent';
            }
        }
        
        if (utilitiesInput) utilitiesInput.value = '£ 140';
        if (maintenancePercentInput) maintenancePercentInput.value = '10 %';
        
        // Set up booking fees toggle (only if not already set up)
        const bookingFeeToggles = section.querySelectorAll(`[data-fee-for="booking"][data-calculator="${calculatorType}"]`);
        bookingFeeToggles.forEach(toggle => {
            // Remove existing listeners to avoid duplicates
            const newToggle = toggle.cloneNode(true);
            toggle.parentNode.replaceChild(newToggle, toggle);
            
            newToggle.addEventListener('click', () => {
                bookingFeeToggles.forEach(t => {
                    if (t !== newToggle) t.classList.remove('pe-toggle-small-active');
                });
                newToggle.classList.add('pe-toggle-small-active');
                
                const bookingFeeInput = section.querySelector(`#${calculatorType}_booking_fees`);
                if (bookingFeeInput) {
                    const feeType = newToggle.dataset.feeType;
                    bookingFeeInput.dataset.feeType = feeType;
                }
                
                calculateRentToHMOValues(calculatorType);
            });
        });
        
    } else {
        // Show HMO rental income section, hide Holiday Let income
        if (rentalIncomeSection) rentalIncomeSection.style.display = 'block';
        if (holidayLetIncomeSection) holidayLetIncomeSection.style.display = 'none';
        
        // Show Monthly Rent to Owner in Ongoing Costs
        if (monthlyRentToOwnerField) monthlyRentToOwnerField.style.display = 'block';
        
        // Hide Booking Fees
        if (bookingFeesField) bookingFeesField.style.display = 'none';
        
        // Change TV License label back
        if (tvLicenseLabel) tvLicenseLabel.textContent = 'Communal TV License';
        
        // Update defaults for HMO
        if (agentFeesInput) {
            agentFeesInput.value = '15 %';
            // Set active toggle to percent
            const percentToggle = section.querySelector('[data-fee-for="agent"][data-fee-type="percent"]');
            const currencyToggle = section.querySelector('[data-fee-for="agent"][data-fee-type="currency"]');
            if (percentToggle && currencyToggle) {
                currencyToggle.classList.remove('pe-toggle-small-active');
                percentToggle.classList.add('pe-toggle-small-active');
                agentFeesInput.dataset.feeType = 'percent';
            }
        }
        
        if (utilitiesInput) utilitiesInput.value = '£ 165';
        if (maintenancePercentInput) maintenancePercentInput.value = '2 %';
    }
}

function calculateRentToHMOValues(calculatorType) {
    console.log('[Frontend] calculateRentToHMOValues called for:', calculatorType);
    
    // Get input values - Acquisition
    const deposit = parseCurrency(document.getElementById(`${calculatorType}_deposit`)?.value || '0');
    const surveyCosts = parseCurrency(document.getElementById(`${calculatorType}_survey_costs`)?.value || '0');
    const legalFees = parseCurrency(document.getElementById(`${calculatorType}_legal_fees`)?.value || '0');
    const referenceFees = parseCurrency(document.getElementById(`${calculatorType}_reference_fees`)?.value || '0');
    
    // Additional acquisition costs
    let additionalAcquisitionCosts = 0;
    const additionalAcquisitionCostInputs = document.querySelectorAll(`[id^="${calculatorType}_additional_acquisition_cost_"]`);
    additionalAcquisitionCostInputs.forEach(input => {
        additionalAcquisitionCosts += parseCurrency(input.value || '0');
    });
    
    // Refurb
    const refurbCostInput = document.getElementById(`${calculatorType}_refurb_cost`);
    const refurbEnabledCheckbox = document.getElementById(`${calculatorType}_refurb_enabled`);
    const isRefurbEnabled = refurbEnabledCheckbox ? refurbEnabledCheckbox.checked : true;
    const refurbCost = isRefurbEnabled ? parseCurrency(refurbCostInput?.value || '0') : 0;
    
    // Rent to Owner
    const monthlyRentToOwner = parseCurrency(document.getElementById(`${calculatorType}_monthly_rent_to_owner`)?.value || '0');
    
    // Get strategy
    const strategySelect = document.getElementById(`${calculatorType}_chosen_strategy`);
    const strategy = strategySelect?.value || 'hmo';
    
    // Calculate total rental income based on strategy
    let totalRentalIncome = 0;
    let annualRentalIncome = 0;
    
    if (strategy === 'holiday-let') {
        // Holiday Let: Calculate from nightly rate and occupancy
        const nightlyRate = parseCurrency(document.getElementById(`${calculatorType}_nightly_rate`)?.value || '0');
        const occupancyRateValue = document.getElementById(`${calculatorType}_occupancy_rate`)?.value || '70 %';
        const occupancyRate = parseFloat(occupancyRateValue.replace(/[%\s]/g, '')) || 70;
        
        // Annual rent = nightly rate × 365 days × (occupancy rate / 100)
        if (nightlyRate > 0 && occupancyRate > 0) {
            annualRentalIncome = nightlyRate * 365 * (occupancyRate / 100);
            totalRentalIncome = annualRentalIncome / 12; // Monthly equivalent
        }
    } else {
        // HMO: Calculate from room rents
        const roomInputs = document.querySelectorAll(`[id^="${calculatorType}_room_"][id$="_rent"]`);
        roomInputs.forEach(input => {
            const roomRent = parseCurrency(input.value || '0');
            totalRentalIncome += roomRent;
        });
        annualRentalIncome = totalRentalIncome * 12;
    }
    
    // Calculate expenses during refurb period (if refurb is enabled)
    let expensesDuringRefurb = 0;
    if (isRefurbEnabled) {
        const vacantPeriod = parseFloat(document.getElementById(`${calculatorType}_vacant_period`)?.value || '0');
        
        // Annual expenses during refurb (prorated)
        const refurbCouncilTaxAnnual = parseCurrency(document.getElementById(`${calculatorType}_refurb_council_tax`)?.value || '0');
        const refurbCouncilTaxDuringRefurb = (refurbCouncilTaxAnnual / 12) * vacantPeriod;
        
        // Additional annual expenses during refurb (prorated)
        let additionalAnnualExpensesDuringRefurb = 0;
        const additionalAnnualExpensesRefurbInputs = document.querySelectorAll(`[id^="${calculatorType}_additional_annual_expense_refurb_"]`);
        additionalAnnualExpensesRefurbInputs.forEach(input => {
            const value = parseCurrency(input.value || '0');
            additionalAnnualExpensesDuringRefurb += (value / 12) * vacantPeriod;
        });
        
        // Monthly expenses during refurb (multiply by vacant period)
        const refurbElectricGas = parseCurrency(document.getElementById(`${calculatorType}_refurb_electric_gas`)?.value || '0') * vacantPeriod;
        const refurbWater = parseCurrency(document.getElementById(`${calculatorType}_refurb_water`)?.value || '0') * vacantPeriod;
        const refurbInsurance = parseCurrency(document.getElementById(`${calculatorType}_refurb_insurance`)?.value || '0') * vacantPeriod;
        
        // Additional monthly expenses during refurb
        let additionalMonthlyExpensesDuringRefurb = 0;
        const additionalMonthlyExpensesRefurbInputs = document.querySelectorAll(`[id^="${calculatorType}_additional_monthly_expense_refurb_"]`);
        additionalMonthlyExpensesRefurbInputs.forEach(input => {
            const value = parseCurrency(input.value || '0');
            additionalMonthlyExpensesDuringRefurb += value * vacantPeriod;
        });
        
        expensesDuringRefurb = refurbCouncilTaxDuringRefurb + refurbElectricGas + refurbWater + refurbInsurance + 
                               additionalAnnualExpensesDuringRefurb + additionalMonthlyExpensesDuringRefurb;
        
        // Add monthly rent to owner during vacant period (this is part of the initial investment)
        const rentToOwnerDuringRefurb = monthlyRentToOwner * vacantPeriod;
        expensesDuringRefurb += rentToOwnerDuringRefurb;
    }
    
    // Calculate total investment
    const totalInvestment = deposit + surveyCosts + legalFees + referenceFees + additionalAcquisitionCosts + refurbCost + expensesDuringRefurb;
    
    // Calculate annual expenses
    const annualRentToOwner = monthlyRentToOwner * 12;
    
    // Annual expenses (from detailed view)
    const councilTax = parseCurrency(document.getElementById(`${calculatorType}_council_tax`)?.value || '0');
    const maintenancePercentValue = document.getElementById(`${calculatorType}_maintenance_percent`)?.value || (strategy === 'holiday-let' ? '10 %' : '2 %');
    const maintenancePercent = parseFloat(maintenancePercentValue.replace(/[%\s]/g, '')) || (strategy === 'holiday-let' ? 10 : 2);
    const annualMaintenance = annualRentalIncome * (maintenancePercent / 100);
    
    // TV License
    // Check if checkbox exists and is checked
    const tvLicenseCheckbox = document.getElementById(`${calculatorType}_communal_tv_license`);
    let annualTVLicense = 0;
    if (tvLicenseCheckbox && tvLicenseCheckbox.checked) {
        // TV License value is 159 (standard UK TV license fee)
        annualTVLicense = 159;
    }
    
    // Additional annual expenses
    let additionalAnnualExpenses = 0;
    const additionalAnnualExpenseInputs = document.querySelectorAll(`[id^="${calculatorType}_additional_annual_expense_"]`);
    additionalAnnualExpenseInputs.forEach(input => {
        additionalAnnualExpenses += parseCurrency(input.value || '0');
    });
    
    // Monthly expenses (convert to annual)
    const utilities = parseCurrency(document.getElementById(`${calculatorType}_utilities`)?.value || '0') * 12;
    const water = parseCurrency(document.getElementById(`${calculatorType}_water`)?.value || '0') * 12;
    const broadbandTV = parseCurrency(document.getElementById(`${calculatorType}_broadband_tv`)?.value || '0') * 12;
    const insurance = parseCurrency(document.getElementById(`${calculatorType}_insurance`)?.value || '0') * 12;
    const cleaningCosts = parseCurrency(document.getElementById(`${calculatorType}_cleaning_costs`)?.value || '0') * 12;
    
    // Agent fees
    let annualAgentFees = 0;
    const agentFeesInput = document.getElementById(`${calculatorType}_agent_fees`);
    if (agentFeesInput) {
        const agentFeesValue = agentFeesInput.value || '0';
        // Check if fee type is set, otherwise check which toggle button is active
        let feeType = agentFeesInput.dataset.feeType;
        if (!feeType) {
            const activeToggle = document.querySelector(`[data-fee-for="agent"][data-calculator="${calculatorType}"].pe-toggle-small-active`);
            feeType = activeToggle ? activeToggle.dataset.feeType : 'percent';
            // Set it on the input for future reference
            agentFeesInput.dataset.feeType = feeType;
        }
        
        if (feeType === 'percent') {
            const agentFeePercent = parseFloat(agentFeesValue.replace(/[%\s]/g, '')) || 0;
            annualAgentFees = annualRentalIncome * (agentFeePercent / 100);
        } else {
            annualAgentFees = parseCurrency(agentFeesValue) * 12;
        }
    }
    
    // Booking fees (Holiday Let only)
    let annualBookingFees = 0;
    if (strategy === 'holiday-let') {
        const bookingFeesInput = document.getElementById(`${calculatorType}_booking_fees`);
        if (bookingFeesInput) {
            const bookingFeesValue = bookingFeesInput.value || '0';
            let feeType = bookingFeesInput.dataset.feeType;
            if (!feeType) {
                const activeToggle = document.querySelector(`[data-fee-for="booking"][data-calculator="${calculatorType}"].pe-toggle-small-active`);
                feeType = activeToggle ? activeToggle.dataset.feeType : 'percent';
                bookingFeesInput.dataset.feeType = feeType;
            }
            
            if (feeType === 'percent') {
                const bookingFeePercent = parseFloat(bookingFeesValue.replace(/[%\s]/g, '')) || 0;
                annualBookingFees = annualRentalIncome * (bookingFeePercent / 100);
            } else {
                annualBookingFees = parseCurrency(bookingFeesValue) * 12;
            }
        }
    }
    
    // Additional monthly expenses
    let additionalMonthlyExpenses = 0;
    const additionalMonthlyExpenseInputs = document.querySelectorAll(`[id^="${calculatorType}_additional_monthly_expense_"]`);
    additionalMonthlyExpenseInputs.forEach(input => {
        additionalMonthlyExpenses += parseCurrency(input.value || '0') * 12;
    });
    
    // Calculate total annual expenses
    // Base expenses (common to both strategies)
    // Use exact decimal values for maintenance and booking fees (don't round individually)
    let totalAnnualExpenses = councilTax + annualMaintenance + annualTVLicense + 
                              utilities + water + broadbandTV + insurance + annualAgentFees + 
                              cleaningCosts + additionalAnnualExpenses + additionalMonthlyExpenses;
    
    // Add rent to owner for both strategies
    totalAnnualExpenses += annualRentToOwner;
    
    // Add booking fees for Holiday Let
    if (strategy === 'holiday-let') {
        totalAnnualExpenses += annualBookingFees;
    }
    
    // Round total expenses to nearest pound at the end
    totalAnnualExpenses = Math.round(totalAnnualExpenses);
    
    // Debug logging for Rent to HMO calculations
    if (calculatorType === 'rent-to-hmo') {
        console.log('💰 Rent to HMO Calculation Breakdown:', {
            strategy,
            annualRentalIncome,
            annualRentToOwner,
            councilTax,
            annualMaintenance,
            annualTVLicense,
            utilities,
            water,
            broadbandTV,
            insurance,
            annualAgentFees,
            annualBookingFees,
            cleaningCosts,
            additionalAnnualExpenses,
            additionalMonthlyExpenses,
            totalAnnualExpenses,
            totalBeforeRounding: councilTax + annualMaintenance + annualTVLicense + 
                                utilities + water + broadbandTV + insurance + annualAgentFees + 
                                cleaningCosts + additionalAnnualExpenses + additionalMonthlyExpenses +
                                annualRentToOwner + (strategy === 'holiday-let' ? annualBookingFees : 0),
            annualProfit: annualRentalIncome - totalAnnualExpenses
        });
    }
    
    // Calculate profit
    const annualProfit = annualRentalIncome - totalAnnualExpenses;
    const monthlyProfit = annualProfit / 12;
    
    // Calculate ROI
    const roi = totalInvestment > 0 ? (annualProfit / totalInvestment) * 100 : 0;
    
    // Update display values
    const totalInvestmentEl = document.getElementById(`${calculatorType}_total_investment`);
    if (totalInvestmentEl) {
        totalInvestmentEl.value = formatCurrency(totalInvestment);
    }
    
    const totalRentalIncomeEl = document.getElementById(`${calculatorType}_total_rental_income`);
    if (totalRentalIncomeEl) {
        totalRentalIncomeEl.value = `${formatCurrency(totalRentalIncome)} / pcm`;
    }
    
    // Update ongoing monthly rent to owner (readonly field in detailed view)
    const ongoingMonthlyRentToOwnerEl = document.getElementById(`${calculatorType}_ongoing_monthly_rent_to_owner`);
    if (ongoingMonthlyRentToOwnerEl) {
        ongoingMonthlyRentToOwnerEl.value = formatCurrency(monthlyRentToOwner);
    }
    
    const totalAnnualExpensesEl = document.getElementById(`${calculatorType}_total_annual_expenses`);
    if (totalAnnualExpensesEl) {
        totalAnnualExpensesEl.value = formatCurrency(totalAnnualExpenses);
    }
    
    const annualProfitEl = document.getElementById(`${calculatorType}_annual_profit`);
    if (annualProfitEl) {
        annualProfitEl.value = formatCurrency(annualProfit);
    }
    
    const monthlyProfitEl = document.getElementById(`${calculatorType}_monthly_profit`);
    if (monthlyProfitEl) {
        monthlyProfitEl.value = formatCurrency(monthlyProfit);
    }
    
    const roiEl = document.getElementById(`${calculatorType}_roi`);
    if (roiEl) {
        roiEl.textContent = roi >= 0 ? `${roi.toFixed(1)}%` : `-${Math.abs(roi).toFixed(1)}%`;
    }
}

function setupBRRCalculatorEvents(section, calculatorType) {
    console.log('[Frontend] setupBRRCalculatorEvents called for:', calculatorType, '| Section exists:', !!section);
    // Get financing payment field reference (used in multiple places)
    const financingPaymentField = section.querySelector(`#${calculatorType}_financing_payment_field`);
    
    // Detailed view toggle
    const detailedViewToggle = section.querySelector(`#${calculatorType}_detailed_view`);
    console.log('[Frontend] Detailed view toggle found:', !!detailedViewToggle, '| Calculator type:', calculatorType);
    if (detailedViewToggle) {
        detailedViewToggle.addEventListener('change', (e) => {
            const detailedFields = section.querySelectorAll('.pe-field-detailed');
            const hiddenInput = section.querySelector(`#${calculatorType}_financing_type_hidden`);
            const financingType = hiddenInput?.value || 'bridging';
            
            detailedFields.forEach(field => {
                if (e.target.checked) {
                    // Show fields based on financing type
                    if (field.classList.contains('pe-field-bridging') && financingType === 'bridging') {
                        field.style.display = 'block';
                    } else if (field.classList.contains('pe-field-mortgage') && financingType === 'mortgage') {
                        // For mortgage fields, check if it's a repayment field
                        if (field.classList.contains('pe-field-repayment')) {
                            const repaymentBtn = section.querySelector('[data-mortgage-type="repayment"]');
                            const isRepayment = repaymentBtn?.classList.contains('pe-financing-type-active') || false;
                            field.style.display = isRepayment ? 'block' : 'none';
                        } else {
                            field.style.display = 'block';
                        }
                    } else if (!field.classList.contains('pe-field-bridging') && !field.classList.contains('pe-field-mortgage')) {
                        field.style.display = 'block';
                    } else {
                        field.style.display = 'none';
                    }
                } else {
                    field.style.display = 'none';
                }
            });
            
            // Hide the top financing payment field when detailed view is enabled and mortgage is selected
            if (e.target.checked && financingType === 'mortgage') {
                if (financingPaymentField) {
                    financingPaymentField.style.display = 'none';
                }
            } else {
                // Show it again when detailed view is disabled or financing type changes
                if (financingPaymentField) {
                    if (financingType === 'cash') {
                        financingPaymentField.style.display = 'none';
                    } else {
                        financingPaymentField.style.display = 'block';
                    }
                }
            }
            // Update label text
            const label = detailedViewToggle.parentElement.querySelector('span');
            if (label) {
                label.textContent = e.target.checked ? 'Switch to simple view' : 'Switch to detailed view';
            }
        });
    }
    
    // Toggle switches
    const refurbToggle = section.querySelector(`#${calculatorType}_refurb_enabled`);
    
    if (refurbToggle) {
        refurbToggle.addEventListener('change', (e) => {
            const content = section.querySelector(`#${calculatorType}_refurb_content`);
            if (content) {
                content.style.display = e.target.checked ? 'block' : 'none';
            }
            calculateBRRValues(calculatorType);
        });
    }
    
    // Helper function to update "Include in Bridging Finance" button state
    const updateIncludeInBridgingButton = (financingType) => {
        const includeInBridgingBtn = section.querySelector(`#${calculatorType}_include_in_bridging`);
        if (includeInBridgingBtn) {
            if (financingType === 'bridging') {
                // Enable the button when bridging finance is selected
                includeInBridgingBtn.disabled = false;
            } else {
                // Disable the button when mortgage or cash is selected
                includeInBridgingBtn.disabled = true;
                // Also uncheck it if it was checked
                if (includeInBridgingBtn.classList.contains('pe-toggle-button-active')) {
                    includeInBridgingBtn.classList.remove('pe-toggle-button-active');
                    calculateBRRValues(calculatorType);
                }
            }
        }
    };
    
    // Financing type buttons (Mortgage/Bridging/Cash - not the mortgage type buttons)
    const financingTypeButtons = section.querySelectorAll('[data-type]');
    const financingPaymentLabel = section.querySelector(`#${calculatorType}_financing_payment_label`);
    
    // Initialize field state based on default selected financing type
    const defaultFinancingBtn = section.querySelector('[data-type].pe-financing-type-active');
    if (defaultFinancingBtn && financingPaymentField && financingPaymentLabel) {
        const defaultType = defaultFinancingBtn.dataset.type;
        if (defaultType === 'cash') {
            financingPaymentField.style.display = 'none';
        } else {
            financingPaymentField.style.display = 'block';
            if (defaultType === 'mortgage') {
                financingPaymentLabel.textContent = 'Mortgage Payments / pcm';
            } else if (defaultType === 'bridging') {
                financingPaymentLabel.textContent = 'Bridging Interest / pcm';
            }
        }
        // Initialize "Include in Bridging Finance" button state
        updateIncludeInBridgingButton(defaultType);
    }
    
    financingTypeButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            financingTypeButtons.forEach(b => b.classList.remove('pe-financing-type-active'));
            btn.classList.add('pe-financing-type-active');
            const hiddenInput = section.querySelector(`#${calculatorType}_financing_type_hidden`);
            const selectedType = btn.dataset.type;
            
            if (hiddenInput) {
                hiddenInput.value = selectedType;
            }
            
            // Update the payment field label and visibility based on financing type
            if (financingPaymentField && financingPaymentLabel) {
                if (selectedType === 'cash') {
                    // Hide the field for cash
                    financingPaymentField.style.display = 'none';
                } else {
                    // Show the field for mortgage or bridging
                    financingPaymentField.style.display = 'block';
                    if (selectedType === 'mortgage') {
                        financingPaymentLabel.textContent = 'Mortgage Payments / pcm';
                        // Set Interest Only as default when switching to Mortgage
                        const interestOnlyBtn = section.querySelector('[data-mortgage-type="interest_only"]');
                        const repaymentBtn = section.querySelector('[data-mortgage-type="repayment"]');
                        if (interestOnlyBtn && repaymentBtn) {
                            repaymentBtn.classList.remove('pe-financing-type-active');
                            interestOnlyBtn.classList.add('pe-financing-type-active');
                            // Hide Mortgage Term field
                            const repaymentFields = section.querySelectorAll('.pe-field-repayment');
                            repaymentFields.forEach(field => {
                                field.style.display = 'none';
                            });
                        }
                    } else if (selectedType === 'bridging') {
                        financingPaymentLabel.textContent = 'Bridging Interest / pcm';
                    }
                }
            }
            
            // Update detailed view fields visibility based on financing type
            const detailedViewToggle = section.querySelector(`#${calculatorType}_detailed_view`);
            if (detailedViewToggle && detailedViewToggle.checked) {
                const bridgingFields = section.querySelectorAll('.pe-field-bridging');
                const mortgageFields = section.querySelectorAll('.pe-field-mortgage');
                const repaymentFields = section.querySelectorAll('.pe-field-repayment');
                
                bridgingFields.forEach(field => {
                    field.style.display = selectedType === 'bridging' ? 'block' : 'none';
                });
                
                mortgageFields.forEach(field => {
                    if (selectedType === 'mortgage') {
                        // Show all mortgage fields first
                        field.style.display = 'block';
                        // Then hide repayment fields if repayment is not selected
                        if (field.classList.contains('pe-field-repayment')) {
                            const repaymentBtn = section.querySelector('[data-mortgage-type="repayment"]');
                            const isRepayment = repaymentBtn?.classList.contains('pe-financing-type-active') || false;
                            field.style.display = isRepayment ? 'block' : 'none';
                        }
                    } else {
                        field.style.display = 'none';
                    }
                });
                
                // Hide the top financing payment field when mortgage is selected in detailed view
                if (selectedType === 'mortgage') {
                    if (financingPaymentField) {
                        financingPaymentField.style.display = 'none';
                    }
                } else {
                    // Show it for bridging or when switching away from mortgage
                    if (financingPaymentField && selectedType !== 'cash') {
                        financingPaymentField.style.display = 'block';
                    }
                }
            } else {
                // Show the top financing payment field when detailed view is disabled
                if (financingPaymentField && selectedType !== 'cash') {
                    financingPaymentField.style.display = 'block';
                }
            }
            
            // Update "Include in Bridging Finance" button state
            updateIncludeInBridgingButton(selectedType);
            
            calculateBRRValues(calculatorType);
        });
    });
    
    // Include in Bridging Finance toggle button
    const includeInBridgingBtn = section.querySelector(`#${calculatorType}_include_in_bridging`);
    if (includeInBridgingBtn) {
        includeInBridgingBtn.addEventListener('click', () => {
            // Button is disabled when not bridging, so we can safely toggle
            includeInBridgingBtn.classList.toggle('pe-toggle-button-active');
            calculateBRRValues(calculatorType);
        });
    }
    
    // Bridging Set-up Fee toggle (between % and £)
    const bridgingFeeTypeButtons = section.querySelectorAll('[data-fee-type]:not([data-fee-for="mortgage"]):not([data-fee-for="refinance"])');
    bridgingFeeTypeButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const feeType = btn.dataset.feeType;
            const setupFeeInput = section.querySelector(`#${calculatorType}_bridging_setup_fee`);
            
            bridgingFeeTypeButtons.forEach(b => b.classList.remove('pe-toggle-small-active'));
            btn.classList.add('pe-toggle-small-active');
            
            // Update the input value format based on selected type
            if (setupFeeInput && feeType === 'currency') {
                const currentValue = setupFeeInput.value.replace(/[£%\s,]/g, '');
                setupFeeInput.value = `£ ${currentValue}`;
            } else if (setupFeeInput && feeType === 'percent') {
                const currentValue = setupFeeInput.value.replace(/[£%\s,]/g, '');
                setupFeeInput.value = `${currentValue} %`;
            }
            
            calculateBRRValues(calculatorType);
        });
    });
    
    // Mortgage Set-up Fee toggle (between £ and %)
    const mortgageFeeTypeButtons = section.querySelectorAll('[data-fee-type][data-fee-for="mortgage"]');
    mortgageFeeTypeButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const feeType = btn.dataset.feeType;
            const setupFeeInput = section.querySelector(`#${calculatorType}_mortgage_setup_fee`);
            
            mortgageFeeTypeButtons.forEach(b => b.classList.remove('pe-toggle-small-active'));
            btn.classList.add('pe-toggle-small-active');
            
            // Update the input value format based on selected type
            if (setupFeeInput && feeType === 'currency') {
                const currentValue = setupFeeInput.value.replace(/[£%\s,]/g, '');
                setupFeeInput.value = `£ ${currentValue}`;
            } else if (setupFeeInput && feeType === 'percent') {
                const currentValue = setupFeeInput.value.replace(/[£%\s,]/g, '');
                setupFeeInput.value = `${currentValue} %`;
            }
            
            calculateBRRValues(calculatorType);
        });
    });
    
    // Refinance Set-up Fee toggle (between £ and %)
    const refinanceFeeTypeButtons = section.querySelectorAll('[data-fee-type][data-fee-for="refinance"]');
    refinanceFeeTypeButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const feeType = btn.dataset.feeType;
            const setupFeeInput = section.querySelector(`#${calculatorType}_refinance_setup_fee`);
            
            refinanceFeeTypeButtons.forEach(b => b.classList.remove('pe-toggle-small-active'));
            btn.classList.add('pe-toggle-small-active');
            
            // Update the input value format based on selected type
            if (setupFeeInput && feeType === 'currency') {
                const currentValue = setupFeeInput.value.replace(/[£%\s,]/g, '');
                setupFeeInput.value = `£ ${currentValue}`;
            } else if (setupFeeInput && feeType === 'percent') {
                const currentValue = setupFeeInput.value.replace(/[£%\s,]/g, '');
                setupFeeInput.value = `${currentValue} %`;
            }
            
            calculateBRRValues(calculatorType);
        });
    });
    
    // Mortgage type buttons (Initial Financing)
    const initialMortgageTypeButtons = section.querySelectorAll('[data-mortgage-type]');
    initialMortgageTypeButtons.forEach(btn => {
        btn.addEventListener('click', (e) => {
            e.stopPropagation(); // Prevent event bubbling
            initialMortgageTypeButtons.forEach(b => b.classList.remove('pe-financing-type-active'));
            btn.classList.add('pe-financing-type-active');
            
            const mortgageType = btn.dataset.mortgageType;
            const repaymentFields = section.querySelectorAll('.pe-field-repayment');
            const detailedViewToggle = section.querySelector(`#${calculatorType}_detailed_view`);
            const hiddenInput = section.querySelector(`#${calculatorType}_financing_type_hidden`);
            const financingType = hiddenInput?.value || 'bridging';
            
            // Show/hide Mortgage Term field based on mortgage type
            // Only update if detailed view is enabled and mortgage is selected
            if (detailedViewToggle && detailedViewToggle.checked && financingType === 'mortgage') {
                // Show/hide only the repayment-specific fields
                repaymentFields.forEach(field => {
                    field.style.display = mortgageType === 'repayment' ? 'block' : 'none';
                });
                // Ensure all other mortgage fields remain visible
                const allMortgageFields = section.querySelectorAll('.pe-field-mortgage');
                allMortgageFields.forEach(field => {
                    // Don't hide non-repayment fields
                    if (!field.classList.contains('pe-field-repayment')) {
                        field.style.display = 'block';
                    }
                });
            }
            
            calculateBRRValues(calculatorType);
        });
    });
    
    // Mortgage type buttons (Refinance)
    const refinanceMortgageTypeButtons = section.querySelectorAll('[data-refinance-mortgage-type]');
    refinanceMortgageTypeButtons.forEach(btn => {
        btn.addEventListener('click', (e) => {
            e.stopPropagation(); // Prevent event bubbling
            refinanceMortgageTypeButtons.forEach(b => b.classList.remove('pe-financing-type-active'));
            btn.classList.add('pe-financing-type-active');
            
            const refinanceMortgageType = btn.dataset.refinanceMortgageType;
            const refinanceRepaymentFields = section.querySelectorAll('.pe-field-refinance-repayment');
            
            // Show/hide Mortgage Term field based on refinance mortgage type
            refinanceRepaymentFields.forEach(field => {
                field.style.display = refinanceMortgageType === 'repayment' ? 'block' : 'none';
            });
            
            calculateBRRValues(calculatorType);
        });
    });
    
    // Stamp duty collapsible toggle
    const stampDutyToggle = section.querySelector(`#${calculatorType}_stamp_duty_toggle`);
    if (stampDutyToggle) {
        stampDutyToggle.addEventListener('click', () => {
            const content = document.getElementById(stampDutyToggle.dataset.target);
            if (content) {
                const isHidden = content.style.display === 'none';
                content.style.display = isHidden ? 'block' : 'none';
                const icon = stampDutyToggle.querySelector('.pe-toggle-expand-icon');
                if (icon) {
                    icon.textContent = isHidden ? '▲' : '▼';
                }
            }
        });
    }
    
    // Stamp duty period dropdown and buttons
    const stampDutyPeriod = section.querySelector(`#${calculatorType}_stamp_duty_period`);
    const stampDutyButtons = section.querySelectorAll(`#${calculatorType}_individual_moving, #${calculatorType}_first_time_buyer, #${calculatorType}_overseas_buyer`);
    
    if (stampDutyPeriod) {
        stampDutyPeriod.addEventListener('change', () => {
            calculateBRRValues(calculatorType);
        });
    }
    
    // Individual Moving Home button - show/hide First Time Buyer
    const individualMovingBtn = section.querySelector(`#${calculatorType}_individual_moving`);
    const firstTimeBuyerBtn = section.querySelector(`#${calculatorType}_first_time_buyer`);
    const overseasBuyerBtn = section.querySelector(`#${calculatorType}_overseas_buyer`);
    
    if (individualMovingBtn) {
        individualMovingBtn.addEventListener('click', () => {
            individualMovingBtn.classList.toggle('pe-option-button-active');
            const isChecked = individualMovingBtn.classList.contains('pe-option-button-active');
            
            // Show/hide First Time Buyer option
            if (firstTimeBuyerBtn) {
                firstTimeBuyerBtn.style.display = isChecked ? 'flex' : 'none';
                // If Individual Moving Home is unchecked, also uncheck First Time Buyer
                if (!isChecked && firstTimeBuyerBtn.classList.contains('pe-option-button-active')) {
                    firstTimeBuyerBtn.classList.remove('pe-option-button-active');
                }
            }
            
            calculateBRRValues(calculatorType);
        });
    }
    
    // First Time Buyer button
    if (firstTimeBuyerBtn) {
        firstTimeBuyerBtn.addEventListener('click', () => {
            firstTimeBuyerBtn.classList.toggle('pe-option-button-active');
            calculateBRRValues(calculatorType);
        });
    }
    
    // Overseas Buyer button
    if (overseasBuyerBtn) {
        overseasBuyerBtn.addEventListener('click', () => {
            overseasBuyerBtn.classList.toggle('pe-option-button-active');
            calculateBRRValues(calculatorType);
        });
    }
    
    // Use event delegation on the section to catch all input events
    // This ensures it works even if fields are added dynamically
    let calculationTimeout;
    const triggerCalculation = () => {
        clearTimeout(calculationTimeout);
        // Use requestAnimationFrame to ensure DOM is updated before reading value
        requestAnimationFrame(() => {
            setTimeout(() => {
                calculateBRRValues(calculatorType);
            }, 10); // Small delay to ensure input value is processed
        });
    };
    
    // Event delegation for all inputs in the section
    section.addEventListener('input', (e) => {
        const target = e.target;
        if (target && target.tagName === 'INPUT' && !target.readOnly) {
            const fieldId = target.id || '';
            // SKIP the refinance interest rate field - it has its own handlers
            if (fieldId === `${calculatorType}_refinance_interest_rate`) {
                return; // Don't handle this field through delegation
            }
            if (fieldId.startsWith(calculatorType + '_') || target.hasAttribute('data-calculator')) {
                triggerCalculation();
            }
        }
    }, true); // Use capture phase to catch events early
    
    section.addEventListener('change', (e) => {
        const target = e.target;
        if (target && target.tagName === 'INPUT' && !target.readOnly) {
            const fieldId = target.id || '';
            // SKIP the refinance interest rate field - it has its own handlers
            if (fieldId === `${calculatorType}_refinance_interest_rate`) {
                return; // Don't handle this field through delegation
            }
            if (fieldId.startsWith(calculatorType + '_') || target.hasAttribute('data-calculator')) {
                calculateBRRValues(calculatorType); // Immediate calculation on change
            }
        }
    }, true);
    
    // Also add blur event for fields that might need it (like percentage inputs)
    section.addEventListener('blur', (e) => {
        const target = e.target;
        if (target && target.tagName === 'INPUT' && !target.readOnly) {
            const fieldId = target.id || '';
            if (fieldId === `${calculatorType}_refinance_interest_rate` ||
                fieldId === `${calculatorType}_refinance_ltv` ||
                fieldId === `${calculatorType}_mortgage_interest_rate` ||
                fieldId === `${calculatorType}_mortgage_ltv` ||
                fieldId === `${calculatorType}_bridging_ltv`) {
                calculateBRRValues(calculatorType);
            }
        }
    }, true);
    
    // CRITICAL FIX: Use event delegation on document to handle input events
    // This ensures listeners persist even when fields are recreated by showCalculatorFields
    // Only set up once per calculator type using a flag
    if (!window[`${calculatorType}_refinance_input_delegation_setup`]) {
        window[`${calculatorType}_refinance_input_delegation_setup`] = true;
        
        // DEBUG: Log ALL input events to see if ANY are firing
        document.addEventListener('input', (e) => {
            console.log('🔍 ALL INPUT EVENT - target:', e.target?.id, 'value:', e.target?.value, 'type:', e.target?.type);
            const target = e.target;
            const fieldId = target?.id;
            
            // Handle refinance interest rate
            if (fieldId === `${calculatorType}_refinance_interest_rate`) {
                const typedValue = target.value;
                console.log('📝 DELEGATED INPUT EVENT (REFINANCE) - User typed:', typedValue);
                // Store the value in multiple places
                target.setAttribute('data-user-value', typedValue);
                window[`${calculatorType}_refinance_interest_rate_value`] = typedValue;
                // Trigger calculation
                setTimeout(() => calculateBRRValues(calculatorType), 50);
            }
            
            // Handle mortgage interest rate (initial financing)
            if (fieldId === `${calculatorType}_mortgage_interest_rate`) {
                const typedValue = target.value;
                console.log('📝 DELEGATED INPUT EVENT (MORTGAGE) - User typed:', typedValue);
                // Store the value
                target.setAttribute('data-user-value', typedValue);
                window[`${calculatorType}_mortgage_interest_rate_value`] = typedValue;
                // Trigger calculation
                setTimeout(() => calculateBRRValues(calculatorType), 50);
            }
        }, true); // Use capture phase to catch early
        
        // Also handle beforeinput for earlier detection
        document.addEventListener('beforeinput', (e) => {
            console.log('🔵 ALL BEFOREINPUT - target:', e.target?.id, 'data:', e.data, 'inputType:', e.inputType);
            const target = e.target;
            if (target && target.id === `${calculatorType}_refinance_interest_rate`) {
                console.log('🔵 DELEGATED BEFOREINPUT - data:', e.data, 'inputType:', e.inputType, 'target value:', target.value);
            }
        }, true);
        
        // Also listen for keydown/keyup to see if keys are being pressed
        document.addEventListener('keydown', (e) => {
            const target = e.target;
            if (target && target.id === `${calculatorType}_refinance_interest_rate`) {
                console.log('⌨️ KEYDOWN on refinance field - key:', e.key, 'code:', e.code, 'current value:', target.value);
            }
        }, true);
        
        document.addEventListener('keyup', (e) => {
            const target = e.target;
            if (target && target.id === `${calculatorType}_refinance_interest_rate`) {
                console.log('⌨️ KEYUP on refinance field - key:', e.key, 'code:', e.code, 'current value:', target.value);
            }
        }, true);
        
        // Listen for focus/blur to see if field can receive focus
        document.addEventListener('focus', (e) => {
            const target = e.target;
            if (target && target.id === `${calculatorType}_refinance_interest_rate`) {
                console.log('👁️ FOCUS on refinance field - can receive input:', !target.readOnly && !target.disabled);
            }
        }, true);
        
        document.addEventListener('click', (e) => {
            const target = e.target;
            if (target && target.id === `${calculatorType}_refinance_interest_rate`) {
                console.log('🖱️ CLICK on refinance field - can receive input:', !target.readOnly && !target.disabled, 'value:', target.value);
            }
        }, true);
    }
    
    // Initialize the field value when it's created
    // This runs every time setupBRRCalculatorEvents is called (after field creation)
    setTimeout(() => {
        const refinanceInterestRateInput = document.getElementById(`${calculatorType}_refinance_interest_rate`);
        if (refinanceInterestRateInput) {
            console.log('✅ Found refinance interest rate input:', refinanceInterestRateInput.id);
            
            // CRITICAL: Remove the hardcoded value attribute to allow user input
            refinanceInterestRateInput.removeAttribute('value');
            
            // Set initial value programmatically if empty
            if (!refinanceInterestRateInput.value || refinanceInterestRateInput.value === '') {
                refinanceInterestRateInput.value = '5.5 %';
            }
            
            // Ensure it's editable
            refinanceInterestRateInput.removeAttribute('readonly');
            refinanceInterestRateInput.removeAttribute('disabled');
            refinanceInterestRateInput.readOnly = false;
            refinanceInterestRateInput.disabled = false;
            
            console.log('✅ Field initialized - value:', refinanceInterestRateInput.value, 'readonly:', refinanceInterestRateInput.readOnly, 'disabled:', refinanceInterestRateInput.disabled);
        }
        
        const refinanceLTVInput = document.getElementById(`${calculatorType}_refinance_ltv`);
        if (refinanceLTVInput) {
            refinanceLTVInput.addEventListener('input', () => {
                calculateBRRValues(calculatorType);
            });
            refinanceLTVInput.addEventListener('change', () => {
                calculateBRRValues(calculatorType);
            });
        }
    }, 100);
    
    // Add additional refinance cost functionality
    const addRefinanceCostLink = section.querySelector(`#${calculatorType}_add_refinance_cost`);
    const additionalCostsContainer = section.querySelector(`#${calculatorType}_additional_refinance_costs`);
    let refinanceCostCounter = 0;
    
    if (addRefinanceCostLink && additionalCostsContainer) {
        addRefinanceCostLink.addEventListener('click', (e) => {
            e.preventDefault();
            refinanceCostCounter++;
            const costId = `${calculatorType}_additional_refinance_cost_${refinanceCostCounter}`;
            
            const costField = document.createElement('div');
            costField.className = 'pe-field';
            costField.innerHTML = `
                <label class="pe-field-label">Additional Refinance Cost ${refinanceCostCounter}</label>
                <div class="pe-field-with-action">
                    <input type="text" class="pe-input" id="${costId}" data-original-id="additional_refinance_cost_${refinanceCostCounter}" data-calculator="${calculatorType}" value="£ 0" placeholder="£ 0">
                    <button type="button" class="pe-btn-remove" data-cost-id="${costId}">×</button>
                </div>
            `;
            
            additionalCostsContainer.appendChild(costField);
            
            // Add event listener for the input
            const input = costField.querySelector(`#${costId}`);
            if (input) {
                input.addEventListener('input', () => {
                    calculateBRRValues(calculatorType);
                });
                input.addEventListener('change', () => {
                    calculateBRRValues(calculatorType);
                });
            }
            
            // Add event listener for remove button
            const removeBtn = costField.querySelector('.pe-btn-remove');
            if (removeBtn) {
                removeBtn.addEventListener('click', () => {
                    costField.remove();
                    calculateBRRValues(calculatorType);
                });
            }
        });
    }
    
    // Edit target thresholds modal
    const editThresholdsLink = section.querySelector(`#${calculatorType}_edit_thresholds`);
    const thresholdModal = section.querySelector(`#${calculatorType}_threshold_modal`);
    const useAccountSettingsBtn = section.querySelector(`#${calculatorType}_use_account_settings`);
    const useCalculatorSettingsBtn = section.querySelector(`#${calculatorType}_use_calculator_settings`);
    
    if (editThresholdsLink && thresholdModal) {
        editThresholdsLink.addEventListener('click', (e) => {
            e.preventDefault();
            thresholdModal.style.display = thresholdModal.style.display === 'none' ? 'block' : 'none';
        });
    }
    
    if (useAccountSettingsBtn && useCalculatorSettingsBtn) {
        useAccountSettingsBtn.addEventListener('click', () => {
            useAccountSettingsBtn.classList.add('pe-threshold-option-active');
            useCalculatorSettingsBtn.classList.remove('pe-threshold-option-active');
        });
        
        useCalculatorSettingsBtn.addEventListener('click', () => {
            useCalculatorSettingsBtn.classList.add('pe-threshold-option-active');
            useAccountSettingsBtn.classList.remove('pe-threshold-option-active');
        });
    }
    
    // Close modal when clicking outside
    if (thresholdModal) {
        thresholdModal.addEventListener('click', (e) => {
            if (e.target === thresholdModal) {
                thresholdModal.style.display = 'none';
            }
        });
    }
    
    // Initialize: Hide the top financing payment field if detailed view is checked and mortgage is selected on page load
    setTimeout(() => {
        const hiddenInput = section.querySelector(`#${calculatorType}_financing_type_hidden`);
        const initialFinancingType = hiddenInput?.value || 'bridging';
        if (detailedViewToggle && detailedViewToggle.checked && initialFinancingType === 'mortgage') {
            if (financingPaymentField) {
                financingPaymentField.style.display = 'none';
            }
        }
    }, 50);
    
    // Initial calculation
    setTimeout(() => calculateBRRValues(calculatorType), 100);
}

function calculateStampDuty(price, period, individualMoving, firstTimeBuyer, overseasBuyer, breakdown, isAdditionalProperty = false) {
    let total = 0;
    breakdown.length = 0; // Clear breakdown array
    
    if (period === 'none' || price <= 0) {
        return 0;
    }
    
    // Additional property rates (buy-to-let/second home) - applies to all periods
    if (isAdditionalProperty) {
        if (period === 'current') {
            // Additional property rates from 1 April 2025
            if (price <= 125000) {
                total = price * 0.05;
                breakdown.push({ label: 'Up to £125k @ 5%', value: total });
            } else if (price <= 250000) {
                const band1 = 125000 * 0.05;
                const band2 = (price - 125000) * 0.07;
                total = band1 + band2;
                breakdown.push({ label: 'Up to £125k @ 5%', value: band1 });
                breakdown.push({ label: '£125k - £250k @ 7%', value: band2 });
            } else if (price <= 925000) {
                const band1 = 125000 * 0.05;
                const band2 = 125000 * 0.07;
                const band3 = (price - 250000) * 0.10;
                total = band1 + band2 + band3;
                breakdown.push({ label: 'Up to £125k @ 5%', value: band1 });
                breakdown.push({ label: '£125k - £250k @ 7%', value: band2 });
                breakdown.push({ label: '£250k - £925k @ 10%', value: band3 });
            } else if (price <= 1500000) {
                const band1 = 125000 * 0.05;
                const band2 = 125000 * 0.07;
                const band3 = 675000 * 0.10;
                const band4 = (price - 925000) * 0.15;
                total = band1 + band2 + band3 + band4;
                breakdown.push({ label: 'Up to £125k @ 5%', value: band1 });
                breakdown.push({ label: '£125k - £250k @ 7%', value: band2 });
                breakdown.push({ label: '£250k - £925k @ 10%', value: band3 });
                breakdown.push({ label: '£925k - £1.5M @ 15%', value: band4 });
            } else {
                const band1 = 125000 * 0.05;
                const band2 = 125000 * 0.07;
                const band3 = 675000 * 0.10;
                const band4 = 575000 * 0.15;
                const band5 = (price - 1500000) * 0.17;
                total = band1 + band2 + band3 + band4 + band5;
                breakdown.push({ label: 'Up to £125k @ 5%', value: band1 });
                breakdown.push({ label: '£125k - £250k @ 7%', value: band2 });
                breakdown.push({ label: '£250k - £925k @ 10%', value: band3 });
                breakdown.push({ label: '£925k - £1.5M @ 15%', value: band4 });
                breakdown.push({ label: 'Over £1.5M @ 17%', value: band5 });
            }
        } else {
            // For historical periods, use similar additional property rates
            // (rates may vary, but structure is similar)
            if (price <= 125000) {
                total = price * 0.05;
                breakdown.push({ label: 'Up to £125k @ 5%', value: total });
            } else if (price <= 250000) {
                const band1 = 125000 * 0.05;
                const band2 = (price - 125000) * 0.07;
                total = band1 + band2;
                breakdown.push({ label: 'Up to £125k @ 5%', value: band1 });
                breakdown.push({ label: '£125k - £250k @ 7%', value: band2 });
            } else if (price <= 925000) {
                const band1 = 125000 * 0.05;
                const band2 = 125000 * 0.07;
                const band3 = (price - 250000) * 0.10;
                total = band1 + band2 + band3;
                breakdown.push({ label: 'Up to £125k @ 5%', value: band1 });
                breakdown.push({ label: '£125k - £250k @ 7%', value: band2 });
                breakdown.push({ label: '£250k - £925k @ 10%', value: band3 });
            } else if (price <= 1500000) {
                const band1 = 125000 * 0.05;
                const band2 = 125000 * 0.07;
                const band3 = 675000 * 0.10;
                const band4 = (price - 925000) * 0.15;
                total = band1 + band2 + band3 + band4;
                breakdown.push({ label: 'Up to £125k @ 5%', value: band1 });
                breakdown.push({ label: '£125k - £250k @ 7%', value: band2 });
                breakdown.push({ label: '£250k - £925k @ 10%', value: band3 });
                breakdown.push({ label: '£925k - £1.5M @ 15%', value: band4 });
            } else {
                const band1 = 125000 * 0.05;
                const band2 = 125000 * 0.07;
                const band3 = 675000 * 0.10;
                const band4 = 575000 * 0.15;
                const band5 = (price - 1500000) * 0.17;
                total = band1 + band2 + band3 + band4 + band5;
                breakdown.push({ label: 'Up to £125k @ 5%', value: band1 });
                breakdown.push({ label: '£125k - £250k @ 7%', value: band2 });
                breakdown.push({ label: '£250k - £925k @ 10%', value: band3 });
                breakdown.push({ label: '£925k - £1.5M @ 15%', value: band4 });
                breakdown.push({ label: 'Over £1.5M @ 17%', value: band5 });
            }
        }
        
        // Apply overseas buyer surcharge on top of additional property rates
        if (overseasBuyer) {
            const surcharge = price * 0.02;
            total += surcharge;
            breakdown.push({ label: 'Overseas Buyer Surcharge (2%)', value: surcharge });
        }
        
        return total;
    }
    
    // Current rates (from 1st April 2025) - Standard UK rates
    if (period === 'current') {
        // Calculate base stamp duty (standard rates or first-time buyer rates)
        let baseTotal = 0;
        
        if (firstTimeBuyer && price <= 500000) {
            // First-time buyer relief (from 1 April 2025)
            if (price <= 300000) {
                baseTotal = 0;
                breakdown.push({ label: '£0 - £300k @ 0%', value: 0 });
            } else {
                const band1 = (price - 300000) * 0.05;
                baseTotal = band1;
                breakdown.push({ label: '£0 - £300k @ 0%', value: 0 });
                breakdown.push({ label: '£300k - £500k @ 5%', value: band1 });
            }
        } else {
            // Standard rates (from 1 April 2025) - applies to Individual Moving Home and regular purchases
            if (price <= 125000) {
                baseTotal = 0;
                breakdown.push({ label: '£0 - £125k @ 0%', value: 0 });
            } else if (price <= 250000) {
                const band1 = (price - 125000) * 0.02;
                baseTotal = band1;
                breakdown.push({ label: '£0 - £125k @ 0%', value: 0 });
                breakdown.push({ label: '£125k - £250k @ 2%', value: band1 });
            } else if (price <= 925000) {
                const band1 = 125000 * 0.02;
                const band2 = (price - 250000) * 0.05;
                baseTotal = band1 + band2;
                breakdown.push({ label: '£0 - £125k @ 0%', value: 0 });
                breakdown.push({ label: '£125k - £250k @ 2%', value: band1 });
                breakdown.push({ label: '£250k - £925k @ 5%', value: band2 });
            } else if (price <= 1500000) {
                const band1 = 125000 * 0.02;
                const band2 = 675000 * 0.05;
                const band3 = (price - 925000) * 0.10;
                baseTotal = band1 + band2 + band3;
                breakdown.push({ label: '£0 - £125k @ 0%', value: 0 });
                breakdown.push({ label: '£125k - £250k @ 2%', value: band1 });
                breakdown.push({ label: '£250k - £925k @ 5%', value: band2 });
                breakdown.push({ label: '£925k - £1.5M @ 10%', value: band3 });
            } else {
                const band1 = 125000 * 0.02;
                const band2 = 675000 * 0.05;
                const band3 = 575000 * 0.10;
                const band4 = (price - 1500000) * 0.12;
                baseTotal = band1 + band2 + band3 + band4;
                breakdown.push({ label: '£0 - £125k @ 0%', value: 0 });
                breakdown.push({ label: '£125k - £250k @ 2%', value: band1 });
                breakdown.push({ label: '£250k - £925k @ 5%', value: band2 });
                breakdown.push({ label: '£925k - £1.5M @ 10%', value: band3 });
                breakdown.push({ label: 'Over £1.5M @ 12%', value: band4 });
            }
        }
        
        total = baseTotal;
        
        // Note: Overseas buyer surcharge is applied at the end of the function for all periods
    } else if (period === '2024-2025') {
        // Rates from 31st Oct 2024 - 31st Mar 2025
        if (price <= 250000) {
            total = 0;
            breakdown.push({ label: '£0 - £250k @ 0%', value: 0 });
        } else if (price <= 925000) {
            const band1 = (price - 250000) * 0.05;
            total = band1;
            breakdown.push({ label: '£0 - £250k @ 0%', value: 0 });
            breakdown.push({ label: '£250k - £925k @ 5%', value: band1 });
        } else if (price <= 1500000) {
            const band1 = 675000 * 0.05;
            const band2 = (price - 925000) * 0.10;
            total = band1 + band2;
            breakdown.push({ label: '£0 - £250k @ 0%', value: 0 });
            breakdown.push({ label: '£250k - £925k @ 5%', value: band1 });
            breakdown.push({ label: '£925k - £1.5M @ 10%', value: band2 });
        } else {
            const band1 = 675000 * 0.05;
            const band2 = 575000 * 0.10;
            const band3 = (price - 1500000) * 0.12;
            total = band1 + band2 + band3;
            breakdown.push({ label: '£0 - £250k @ 0%', value: 0 });
            breakdown.push({ label: '£250k - £925k @ 5%', value: band1 });
            breakdown.push({ label: '£925k - £1.5M @ 10%', value: band2 });
            breakdown.push({ label: 'Over £1.5M @ 12%', value: band3 });
        }
    } else if (period === '2022-2024') {
        // Rates from 23rd Sep 2022 - 31st Oct 2024
        // Standard rates (Individual Moving Home uses these)
        if (firstTimeBuyer && price <= 500000) {
            // First-time buyer relief
            if (price <= 300000) {
                total = 0;
                breakdown.push({ label: '£0 - £300k @ 0%', value: 0 });
            } else {
                const band1 = (price - 300000) * 0.05;
                total = band1;
                breakdown.push({ label: '£0 - £300k @ 0%', value: 0 });
                breakdown.push({ label: '£300k - £500k @ 5%', value: band1 });
            }
        } else {
            // Standard rates
            if (price <= 250000) {
                total = 0;
                breakdown.push({ label: '£0 - £250k @ 0%', value: 0 });
            } else if (price <= 925000) {
                const band1 = (price - 250000) * 0.05;
                total = band1;
                breakdown.push({ label: '£0 - £250k @ 0%', value: 0 });
                breakdown.push({ label: '£250k - £925k @ 5%', value: band1 });
            } else if (price <= 1500000) {
                const band1 = 675000 * 0.05;
                const band2 = (price - 925000) * 0.10;
                total = band1 + band2;
                breakdown.push({ label: '£0 - £250k @ 0%', value: 0 });
                breakdown.push({ label: '£250k - £925k @ 5%', value: band1 });
                breakdown.push({ label: '£925k - £1.5M @ 10%', value: band2 });
            } else {
                const band1 = 675000 * 0.05;
                const band2 = 575000 * 0.10;
                const band3 = (price - 1500000) * 0.12;
                total = band1 + band2 + band3;
                breakdown.push({ label: '£0 - £250k @ 0%', value: 0 });
                breakdown.push({ label: '£250k - £925k @ 5%', value: band1 });
                breakdown.push({ label: '£925k - £1.5M @ 10%', value: band2 });
                breakdown.push({ label: 'Over £1.5M @ 12%', value: band3 });
            }
        }
    } else if (period === '2021-2022') {
        // Rates from 1st Oct 2021 - 23rd Sep 2022
        // Standard rates (Individual Moving Home uses these)
        if (firstTimeBuyer && price <= 500000) {
            // First-time buyer relief
            if (price <= 300000) {
                total = 0;
                breakdown.push({ label: '£0 - £300k @ 0%', value: 0 });
            } else {
                const band1 = (price - 300000) * 0.05;
                total = band1;
                breakdown.push({ label: '£0 - £300k @ 0%', value: 0 });
                breakdown.push({ label: '£300k - £500k @ 5%', value: band1 });
            }
        } else {
            // Standard rates
            if (price <= 125000) {
                total = 0;
                breakdown.push({ label: '£0 - £125k @ 0%', value: 0 });
            } else if (price <= 250000) {
                const band1 = (price - 125000) * 0.02;
                total = band1;
                breakdown.push({ label: '£0 - £125k @ 0%', value: 0 });
                breakdown.push({ label: '£125k - £250k @ 2%', value: band1 });
            } else if (price <= 925000) {
                const band1 = 125000 * 0.02;
                const band2 = (price - 250000) * 0.05;
                total = band1 + band2;
                breakdown.push({ label: '£0 - £125k @ 0%', value: 0 });
                breakdown.push({ label: '£125k - £250k @ 2%', value: band1 });
                breakdown.push({ label: '£250k - £925k @ 5%', value: band2 });
            } else if (price <= 1500000) {
                const band1 = 125000 * 0.02;
                const band2 = 675000 * 0.05;
                const band3 = (price - 925000) * 0.10;
                total = band1 + band2 + band3;
                breakdown.push({ label: '£0 - £125k @ 0%', value: 0 });
                breakdown.push({ label: '£125k - £250k @ 2%', value: band1 });
                breakdown.push({ label: '£250k - £925k @ 5%', value: band2 });
                breakdown.push({ label: '£925k - £1.5M @ 10%', value: band3 });
            } else {
                const band1 = 125000 * 0.02;
                const band2 = 675000 * 0.05;
                const band3 = 575000 * 0.10;
                const band4 = (price - 1500000) * 0.12;
                total = band1 + band2 + band3 + band4;
                breakdown.push({ label: '£0 - £125k @ 0%', value: 0 });
                breakdown.push({ label: '£125k - £250k @ 2%', value: band1 });
                breakdown.push({ label: '£250k - £925k @ 5%', value: band2 });
                breakdown.push({ label: '£925k - £1.5M @ 10%', value: band3 });
                breakdown.push({ label: 'Over £1.5M @ 12%', value: band4 });
            }
        }
    } else if (period === '2021-mid') {
        // Rates from 1st Jul 2021 - 30th Sep 2021
        // Standard rates (Individual Moving Home uses these)
        if (firstTimeBuyer && price <= 500000) {
            // First-time buyer relief
            if (price <= 300000) {
                total = 0;
                breakdown.push({ label: '£0 - £300k @ 0%', value: 0 });
            } else {
                const band1 = (price - 300000) * 0.05;
                total = band1;
                breakdown.push({ label: '£0 - £300k @ 0%', value: 0 });
                breakdown.push({ label: '£300k - £500k @ 5%', value: band1 });
            }
        } else {
            // Standard rates
            if (price <= 250000) {
                total = 0;
                breakdown.push({ label: '£0 - £250k @ 0%', value: 0 });
            } else if (price <= 925000) {
                const band1 = (price - 250000) * 0.05;
                total = band1;
                breakdown.push({ label: '£0 - £250k @ 0%', value: 0 });
                breakdown.push({ label: '£250k - £925k @ 5%', value: band1 });
            } else if (price <= 1500000) {
                const band1 = 675000 * 0.05;
                const band2 = (price - 925000) * 0.10;
                total = band1 + band2;
                breakdown.push({ label: '£0 - £250k @ 0%', value: 0 });
                breakdown.push({ label: '£250k - £925k @ 5%', value: band1 });
                breakdown.push({ label: '£925k - £1.5M @ 10%', value: band2 });
            } else {
                const band1 = 675000 * 0.05;
                const band2 = 575000 * 0.10;
                const band3 = (price - 1500000) * 0.12;
                total = band1 + band2 + band3;
                breakdown.push({ label: '£0 - £250k @ 0%', value: 0 });
                breakdown.push({ label: '£250k - £925k @ 5%', value: band1 });
                breakdown.push({ label: '£925k - £1.5M @ 10%', value: band2 });
                breakdown.push({ label: 'Over £1.5M @ 12%', value: band3 });
            }
        }
    }
    
    // Apply overseas buyer surcharge (2% on entire purchase price) - applies to all periods
    if (overseasBuyer) {
        const surcharge = price * 0.02;
        total += surcharge;
        breakdown.push({ label: 'Overseas Buyer Surcharge (2%)', value: surcharge });
    }
    
    // Note: Individual Moving Home uses standard rates (already calculated above)
    // No additional calculation needed for individualMoving - it's just a flag
    
    return total;
}

function calculateBRRValues(calculatorType) {
    // Debug: Log that calculation is being called
    // console.log('calculateBRRValues called for:', calculatorType);
    
    // Get all input values
    const purchasePrice = parseCurrency(document.getElementById(`${calculatorType}_purchase_price`)?.value || '0');
    const refurbCostInput = document.getElementById(`${calculatorType}_refurb_cost`);
    const refurbEnabledCheckbox = document.getElementById(`${calculatorType}_refurb_enabled`);
    const isRefurbEnabled = refurbEnabledCheckbox ? refurbEnabledCheckbox.checked : true; // Default to true if checkbox doesn't exist
    const refurbCost = isRefurbEnabled ? parseCurrency(refurbCostInput?.value || '0') : 0;
    const estimatedMarketValue = parseCurrency(document.getElementById(`${calculatorType}_estimated_market_value`)?.value || '0');
    const monthlyRent = parseCurrency(document.getElementById(`${calculatorType}_monthly_rent`)?.value || '0');
    const bridgingInterest = parseCurrency(document.getElementById(`${calculatorType}_bridging_interest`)?.value || '0');
    const vacantPeriod = isRefurbEnabled ? parseFloat(document.getElementById(`${calculatorType}_vacant_period`)?.value || '0') : 0;
    const appreciation = parseFloat(document.getElementById(`${calculatorType}_appreciation`)?.value?.replace('%', '') || '5.5');
    
    // Get financing type
    const hiddenInput = document.getElementById(`${calculatorType}_financing_type_hidden`);
    const financingType = hiddenInput?.value || 'bridging';
    const includeInBridgingBtn = document.getElementById(`${calculatorType}_include_in_bridging`);
    const includeInBridging = includeInBridgingBtn?.classList.contains('pe-toggle-button-active') || false;
    
    // Calculate stamp duty based on selected period and buyer type
    const stampDutyPeriod = document.getElementById(`${calculatorType}_stamp_duty_period`)?.value || 'current';
    const individualMovingBtn = document.getElementById(`${calculatorType}_individual_moving`);
    const firstTimeBuyerBtn = document.getElementById(`${calculatorType}_first_time_buyer`);
    const overseasBuyerBtn = document.getElementById(`${calculatorType}_overseas_buyer`);
    const individualMoving = individualMovingBtn?.classList.contains('pe-option-button-active') || false;
    const firstTimeBuyer = firstTimeBuyerBtn?.classList.contains('pe-option-button-active') || false;
    const overseasBuyer = overseasBuyerBtn?.classList.contains('pe-option-button-active') || false;
    
    let stampDuty = 0;
    let breakdown = [];
    
    // Show/hide stamp duty field based on selection
    const stampDutyField = document.getElementById(`${calculatorType}_stamp_duty`)?.closest('.pe-field');
    if (stampDutyPeriod === 'none') {
        if (stampDutyField) stampDutyField.style.display = 'none';
        const breakdownEl = document.getElementById(`${calculatorType}_stamp_duty_breakdown`);
        if (breakdownEl) breakdownEl.style.display = 'none';
    } else {
        if (stampDutyField) stampDutyField.style.display = 'block';
        // Only calculate if purchase price is valid
        // For investment calculators (BRR, BTL, etc.), assume it's an additional property (buy-to-let/second home)
        // which uses higher stamp duty rates
        // BUT: If "Individual Moving Home" is checked, use standard residential rates instead
        let isAdditionalProperty = calculatorType !== 'purchase'; // Investment properties are typically additional properties
        if (individualMoving) {
            // Individual Moving Home uses standard residential rates (not additional property rates)
            isAdditionalProperty = false;
        }
        if (purchasePrice > 0) {
            stampDuty = calculateStampDuty(purchasePrice, stampDutyPeriod, individualMoving, firstTimeBuyer, overseasBuyer, breakdown, isAdditionalProperty);
        } else {
            stampDuty = 0;
            breakdown.length = 0;
        }
        
        // Update breakdown display
        const breakdownEl = document.getElementById(`${calculatorType}_stamp_duty_breakdown`);
        if (breakdownEl && breakdown.length > 0) {
            breakdownEl.style.display = 'block';
            breakdownEl.innerHTML = breakdown.map(item => 
                `<div class="pe-breakdown-row">
                    <span class="pe-breakdown-label">${item.label}</span>
                    <span class="pe-breakdown-value">${formatCurrency(item.value)}</span>
                </div>`
            ).join('') + `<div class="pe-breakdown-row pe-breakdown-total">
                    <span class="pe-breakdown-label">Total</span>
                    <span class="pe-breakdown-value">${formatCurrency(stampDuty)}</span>
                </div>`;
        } else if (breakdownEl) {
            breakdownEl.style.display = 'none';
        }
    }
    
    // Calculate initial mortgage based on financing type
    let initialMortgage = 0;
    if (financingType === 'mortgage') {
        const mortgageLTV = parseFloat(document.getElementById(`${calculatorType}_mortgage_ltv`)?.value?.replace(/[%\s]/g, '') || '75');
        initialMortgage = purchasePrice * (mortgageLTV / 100);
    } else if (financingType === 'bridging') {
        const bridgingLTV = parseFloat(document.getElementById(`${calculatorType}_bridging_ltv`)?.value?.replace(/[%\s]/g, '') || '75');
        initialMortgage = purchasePrice * (bridgingLTV / 100);
    }
    // Cash = 0 mortgage
    
    // Read additional purchase costs
    const surveyCostsValue = document.getElementById(`${calculatorType}_survey_costs`)?.value || '£ 500';
    const surveyCosts = parseCurrency(surveyCostsValue);
    
    const legalFeesValue = document.getElementById(`${calculatorType}_legal_fees`)?.value || '£ 1,500';
    const legalFees = parseCurrency(legalFeesValue);
    
    // Additional purchase costs (dynamically added)
    let additionalPurchaseCosts = 0;
    const additionalPurchaseCostInputs = document.querySelectorAll(`[id^="${calculatorType}_additional_purchase_cost_"]`);
    additionalPurchaseCostInputs.forEach(input => {
        const value = parseCurrency(input.value || '£ 0');
        additionalPurchaseCosts += value;
    });
    
    // Read mortgage set-up fee (for initial financing)
    let mortgageSetupFee = 0;
    if (financingType === 'mortgage') {
        const mortgageSetupFeeValue = document.getElementById(`${calculatorType}_mortgage_setup_fee`)?.value || '£ 1,000';
        // Check if it's a percentage or currency
        if (mortgageSetupFeeValue.includes('%')) {
            const percent = parseFloat(mortgageSetupFeeValue.replace(/[%\s]/g, '')) || 0;
            mortgageSetupFee = initialMortgage * (percent / 100);
        } else {
            mortgageSetupFee = parseCurrency(mortgageSetupFeeValue);
        }
    } else if (financingType === 'bridging') {
        const bridgingSetupFeeValue = document.getElementById(`${calculatorType}_bridging_setup_fee`)?.value || '1.75 %';
        // Check if it's a percentage or currency
        if (bridgingSetupFeeValue.includes('%')) {
            const percent = parseFloat(bridgingSetupFeeValue.replace(/[%\s]/g, '')) || 0;
            mortgageSetupFee = initialMortgage * (percent / 100);
        } else {
            mortgageSetupFee = parseCurrency(bridgingSetupFeeValue);
        }
    }
    
    // Refinance calculations - read from input fields (MUST be done before calculating refinance setup fee)
    const refinanceLTVValue = document.getElementById(`${calculatorType}_refinance_ltv`)?.value || '75 %';
    const refinanceLTV = parseFloat(refinanceLTVValue.replace(/[%\s]/g, '')) || 75;
    const refinanceAmount = estimatedMarketValue * (refinanceLTV / 100);
    
    // Read refinance set-up fee (AFTER refinanceAmount is calculated, as it may be a percentage of refinanceAmount)
    let refinanceSetupFee = 0;
    const refinanceSetupFeeValue = document.getElementById(`${calculatorType}_refinance_setup_fee`)?.value || '£ 0';
    // Check if it's a percentage or currency
    if (refinanceSetupFeeValue.includes('%')) {
        const percent = parseFloat(refinanceSetupFeeValue.replace(/[%\s]/g, '')) || 0;
        refinanceSetupFee = refinanceAmount * (percent / 100);
    } else {
        refinanceSetupFee = parseCurrency(refinanceSetupFeeValue);
    }
    
    // Read additional refinance costs (these are costs paid at refinance and should be included in total investment)
    let additionalRefinanceCosts = 0;
    const additionalRefinanceCostInputs = document.querySelectorAll(`[id^="${calculatorType}_additional_refinance_cost_"]`);
    additionalRefinanceCostInputs.forEach(input => {
        const value = parseCurrency(input.value || '£ 0');
        additionalRefinanceCosts += value;
    });
    
    // Total investment (deposit + stamp duty + refurb cost + survey costs + legal fees + mortgage set-up fee + refinance set-up fee + additional refinance costs + additional purchase costs)
    // Note: Expenses during refurb will be added after mortgage payments are calculated
    const deposit = purchasePrice - initialMortgage;
    let totalInvestment = deposit + stampDuty + refurbCost + surveyCosts + legalFees + mortgageSetupFee + refinanceSetupFee + additionalRefinanceCosts + additionalPurchaseCosts;
    // lockedInEquity, moneyBack, and moneyLeftIn will be calculated after expenses during refurb are added
    
    // Debug: Log key values for Money Left In calculation
    console.log('💰 Money Left In Calculation:');
    console.log('  purchasePrice:', purchasePrice);
    console.log('  initialMortgage:', initialMortgage);
    console.log('  deposit:', deposit);
    console.log('  stampDuty:', stampDuty);
    console.log('  refurbCost:', refurbCost);
    console.log('  surveyCosts:', surveyCosts);
    console.log('  legalFees:', legalFees);
    console.log('  mortgageSetupFee:', mortgageSetupFee);
    console.log('  refinanceSetupFee:', refinanceSetupFee);
    console.log('  additionalRefinanceCosts:', additionalRefinanceCosts);
    console.log('  additionalPurchaseCosts:', additionalPurchaseCosts);
    // expensesDuringRefurb, totalInvestment, moneyBack, and moneyLeftIn will be logged after they're calculated
    console.log('  estimatedMarketValue:', estimatedMarketValue);
    console.log('  refinanceLTV:', refinanceLTV);
    console.log('  refinanceAmount:', refinanceAmount);
    
    // Ideal purchase price - calculate based on target ROI and interest rates
    // This will be recalculated after we have the interest rates
    let idealPurchasePrice = estimatedMarketValue * 0.7; // Initial estimate
    
    // Rental calculations
    // For Holiday Let, calculate annual rent from nightly rate and occupancy rate
    let annualRent = 0;
    if (calculatorType === 'holiday-let') {
        const nightlyRate = parseCurrency(document.getElementById(`${calculatorType}_nightly_rate`)?.value || '0');
        const occupancyRateValue = document.getElementById(`${calculatorType}_occupancy_rate`)?.value || '0';
        const occupancyRate = parseFloat(occupancyRateValue.replace(/[%\s]/g, '')) || 0;
        
        // Annual rent = nightly rate × 365 days × (occupancy rate / 100)
        if (nightlyRate > 0 && occupancyRate > 0) {
            annualRent = nightlyRate * 365 * (occupancyRate / 100);
        }
    } else {
        // For BRR and other calculators, use monthly rent
        annualRent = monthlyRent * 12;
    }
    const grossYield = estimatedMarketValue > 0 ? (annualRent / estimatedMarketValue) * 100 : 0;
    
    // Expenses - read mortgage interest rates from inputs
    // IMPORTANT: Always read the current value directly from the DOM element
    const initialInterestRateInput = document.getElementById(`${calculatorType}_mortgage_interest_rate`);
    // CRITICAL: Read the ACTUAL current value from the input field
    // ALWAYS prefer the live .value property as it reflects what the user is currently typing
    let initialInterestRateValue = '5.5 %';
    if (initialInterestRateInput) {
        // Priority 1: Read the LIVE .value property (most reliable)
        const liveValue = initialInterestRateInput.value || '';
        
        // Priority 2: Check stored user value (set by input event)
        const storedUserValue = window[`${calculatorType}_mortgage_interest_rate_value`];
        const dataUserValue = initialInterestRateInput.getAttribute('data-user-value');
        
        // Always use live value if available, otherwise fall back to stored values
        if (liveValue && liveValue.trim() !== '') {
            initialInterestRateValue = liveValue;
        } else if (storedUserValue && storedUserValue !== '') {
            initialInterestRateValue = storedUserValue;
        } else if (dataUserValue && dataUserValue !== '') {
            initialInterestRateValue = dataUserValue;
        } else {
            initialInterestRateValue = initialInterestRateInput.textContent || '5.5 %';
        }
        
        // Debug logging
        if (liveValue && liveValue !== '5.5 %') {
            console.log('📊 Reading initial mortgage rate - liveValue:', liveValue, 'stored:', storedUserValue, 'data-attr:', dataUserValue, 'final:', initialInterestRateValue);
        }
    }
    initialInterestRateValue = String(initialInterestRateValue).trim();
    
    const refinanceInterestRateInput = document.getElementById(`${calculatorType}_refinance_interest_rate`);
    // CRITICAL: Read the ACTUAL current value from the input field
    // ALWAYS prefer the live .value property as it reflects what the user is currently typing
    let refinanceInterestRateValue = '';
    if (refinanceInterestRateInput) {
        // For input elements, ALWAYS use .value property (not .getAttribute('value'))
        // .value is the LIVE value that reflects user input
        if (refinanceInterestRateInput.tagName === 'INPUT' || refinanceInterestRateInput.tagName === 'TEXTAREA') {
            // Priority 1: Read the LIVE .value property (most reliable)
            const liveValue = refinanceInterestRateInput.value || '';
            
            // Priority 2: Check stored user value (set by input event)
            const storedUserValue = window[`${calculatorType}_refinance_interest_rate_value`];
            const dataUserValue = refinanceInterestRateInput.getAttribute('data-user-value');
            
            // Always use live value if available, otherwise fall back to stored values
            if (liveValue && liveValue.trim() !== '') {
                refinanceInterestRateValue = liveValue;
            } else if (storedUserValue && storedUserValue !== '') {
                refinanceInterestRateValue = storedUserValue;
            } else if (dataUserValue && dataUserValue !== '') {
                refinanceInterestRateValue = dataUserValue;
            }
            
            // Debug logging
            if (liveValue && liveValue !== '5.5 %') {
                console.log('📊 Reading refinance rate - liveValue:', liveValue, 'stored:', storedUserValue, 'data-attr:', dataUserValue, 'final:', refinanceInterestRateValue);
            }
        } else {
            refinanceInterestRateValue = refinanceInterestRateInput.textContent || '';
        }
    }
    
    // Clean and trim the value - preserve what user typed
    refinanceInterestRateValue = String(refinanceInterestRateValue).trim();
    
    // Debug: Log what we read for refinance rate
    if (refinanceInterestRateInput) {
        console.log('📊 Reading refinance rate - liveValue:', refinanceInterestRateInput.value, 'stored:', window[`${calculatorType}_refinance_interest_rate_value`], 'data-attr:', refinanceInterestRateInput.getAttribute('data-user-value'), 'final:', refinanceInterestRateValue);
    }
    
    // If empty, use initial rate as fallback
    // This ensures that when user changes initial mortgage rate, it affects refinance calculations
    // BUT: Don't use fallback if the field has focus (user is typing)
    const isFocused = refinanceInterestRateInput && document.activeElement === refinanceInterestRateInput;
    if ((!refinanceInterestRateValue || refinanceInterestRateValue === '' || refinanceInterestRateValue === '5.5 %') && !isFocused) {
        // Use initial mortgage rate as refinance rate if refinance rate is not explicitly set
        refinanceInterestRateValue = initialInterestRateValue;
        console.log('⚠️ Refinance rate not set, using initial mortgage rate as fallback:', initialInterestRateValue);
    }
    
    // Parse the interest rate - handle various formats: "5.5 %", "5.5%", "5.5", etc.
    // Remove all non-numeric characters except decimal point
    const initialMortgageRate = parseFloat(String(initialInterestRateValue).replace(/[£,%\s]/g, '')) || 5.5;
    // Extract just the numeric part (handle cases like "533333333333333.333333333333333335 %")
    const cleanedValue = String(refinanceInterestRateValue).replace(/[£,%\s]/g, '');
    let refinanceMortgageRate = parseFloat(cleanedValue);
    
    // Debug: Log what we're parsing
    console.log('🔢 Parsing rates - initialInterestRateValue:', initialInterestRateValue, 'parsed:', initialMortgageRate);
    console.log('🔢 Parsing rates - refinanceInterestRateValue:', refinanceInterestRateValue, 'cleaned:', cleanedValue, 'parsed:', refinanceMortgageRate);
    
    // If parsing failed or resulted in NaN, use initial rate as fallback
    if (isNaN(refinanceMortgageRate) || refinanceMortgageRate === 0) {
        console.log('⚠️ Refinance rate parsing failed, using initial rate:', initialMortgageRate);
        refinanceMortgageRate = initialMortgageRate;
    }
    
    // Cap the interest rate at a reasonable maximum (e.g., 100%)
    if (refinanceMortgageRate > 100) {
        console.log('⚠️ Refinance rate capped at 100% (was:', refinanceMortgageRate, ')');
        refinanceMortgageRate = 100;
    }
    
    // Debug: Always log the final rates being used
    console.log('✅ Final rates - initialMortgageRate:', initialMortgageRate, 'refinanceMortgageRate:', refinanceMortgageRate);
    
    // Check selected mortgage type (initial financing)
    const repaymentBtnInitial = document.querySelector(`[data-mortgage-type="repayment"][data-calculator="${calculatorType}"]`);
    const isInitialRepayment = repaymentBtnInitial?.classList.contains('pe-financing-type-active') || false;
    const mortgageTermYearsInitial = parseFloat(document.getElementById(`${calculatorType}_mortgage_term_years`)?.value || '25');
    
    // Check refinance mortgage type
    const refinanceRepaymentBtn = document.querySelector(`[data-refinance-mortgage-type="repayment"][data-calculator="${calculatorType}"]`);
    const isRefinanceRepayment = refinanceRepaymentBtn?.classList.contains('pe-financing-type-active') || false;
    const mortgageTermYearsRefi = parseFloat(document.getElementById(`${calculatorType}_refinance_mortgage_term_years`)?.value || '25');
    
    // Initial mortgage payments (drives on-screen "Mortgage Payments / pcm")
    let monthlyMortgagePaymentInitial = 0;
    if (initialMortgage > 0) {
        const monthlyRateInitial = (initialMortgageRate / 100) / 12;
        const numberOfPaymentsInitial = mortgageTermYearsInitial * 12;
        
        // Debug: Log initial mortgage payment calculation
        console.log('💰 Initial Mortgage Payment Calc - initialMortgage:', initialMortgage, 'rate:', initialMortgageRate, 'monthlyRate:', monthlyRateInitial, 'isRepayment:', isInitialRepayment, 'termYears:', mortgageTermYearsInitial);
        
        if (isInitialRepayment && monthlyRateInitial > 0) {
            monthlyMortgagePaymentInitial = initialMortgage * (monthlyRateInitial * Math.pow(1 + monthlyRateInitial, numberOfPaymentsInitial)) /
                                            (Math.pow(1 + monthlyRateInitial, numberOfPaymentsInitial) - 1);
        } else if (isInitialRepayment && monthlyRateInitial === 0) {
            monthlyMortgagePaymentInitial = initialMortgage / numberOfPaymentsInitial;
        } else {
            // Interest-only
            monthlyMortgagePaymentInitial = initialMortgage * (initialMortgageRate / 100) / 12;
        }
        
        console.log('💰 Initial Mortgage Payment Result:', monthlyMortgagePaymentInitial);
    }
    
    // Refinance mortgage payments (drives refinance payment display)
    let monthlyMortgagePaymentRefi = 0;
    let annualMortgageInterest = 0;
    if (refinanceAmount > 0) {
        const monthlyRateRefi = (refinanceMortgageRate / 100) / 12;
        const numberOfPaymentsRefi = mortgageTermYearsRefi * 12;
        
        if (isRefinanceRepayment && monthlyRateRefi > 0) {
            monthlyMortgagePaymentRefi = refinanceAmount * (monthlyRateRefi * Math.pow(1 + monthlyRateRefi, numberOfPaymentsRefi)) /
                                         (Math.pow(1 + monthlyRateRefi, numberOfPaymentsRefi) - 1);
            // Approximate first-year interest
            annualMortgageInterest = refinanceAmount * (refinanceMortgageRate / 100);
        } else if (isRefinanceRepayment && monthlyRateRefi === 0) {
            monthlyMortgagePaymentRefi = refinanceAmount / numberOfPaymentsRefi;
            annualMortgageInterest = 0;
        } else {
            // Interest-only
            annualMortgageInterest = refinanceAmount * (refinanceMortgageRate / 100);
            monthlyMortgagePaymentRefi = annualMortgageInterest / 12;
        }
    }
    
    // Debug: Log mortgage payment calculations
    console.log('Mortgage Payment Calc - refinanceAmount:', refinanceAmount, 'rate:', refinanceMortgageRate, 'monthlyPayment:', monthlyMortgagePaymentRefi, 'annualInterest:', annualMortgageInterest);
    
    // Calculate expenses during refurb period (these need to be included in total investment)
    // Annual expenses (prorated for vacant period)
    const refurbCouncilTaxValue = document.getElementById(`${calculatorType}_refurb_council_tax`)?.value || '£ 1,670';
    const refurbCouncilTaxAnnual = parseCurrency(refurbCouncilTaxValue);
    const refurbCouncilTaxDuringRefurb = (refurbCouncilTaxAnnual / 12) * vacantPeriod;
    
    // Additional annual expenses during refurb (prorated)
    let additionalAnnualExpensesDuringRefurb = 0;
    const additionalAnnualExpensesRefurbInputs = document.querySelectorAll(`[id^="${calculatorType}_additional_annual_expense_refurb_"]`);
    additionalAnnualExpensesRefurbInputs.forEach(input => {
        const value = parseCurrency(input.value || '£ 0');
        additionalAnnualExpensesDuringRefurb += (value / 12) * vacantPeriod; // Prorate for vacant period
    });
    
    // Monthly expenses during refurb (multiply by vacant period)
    const refurbElectricGasValue = document.getElementById(`${calculatorType}_refurb_electric_gas`)?.value || '£ 60';
    const refurbElectricGas = parseCurrency(refurbElectricGasValue) * vacantPeriod;
    
    const refurbWaterValue = document.getElementById(`${calculatorType}_refurb_water`)?.value || '£ 30';
    const refurbWater = parseCurrency(refurbWaterValue) * vacantPeriod;
    
    const refurbInsuranceValue = document.getElementById(`${calculatorType}_refurb_insurance`)?.value || '£ 40';
    const refurbInsurance = parseCurrency(refurbInsuranceValue) * vacantPeriod;
    
    // Additional monthly expenses during refurb
    let additionalMonthlyExpensesDuringRefurb = 0;
    const additionalMonthlyExpensesRefurbInputs = document.querySelectorAll(`[id^="${calculatorType}_additional_monthly_expense_refurb_"]`);
    additionalMonthlyExpensesRefurbInputs.forEach(input => {
        const value = parseCurrency(input.value || '£ 0');
        additionalMonthlyExpensesDuringRefurb += value * vacantPeriod;
    });
    
    // Mortgage payments during vacant period (full payment, not just interest)
    let mortgagePaymentsDuringRefurb = 0;
    if (financingType === 'mortgage' && vacantPeriod > 0) {
        mortgagePaymentsDuringRefurb = monthlyMortgagePaymentInitial * vacantPeriod;
    }
    
    // Bridging costs during vacant period (only if bridging finance)
    // This is a one-time cost during refurb, should be included in total investment but NOT in annual expenses
    let bridgingCost = 0;
    if (financingType === 'bridging') {
        // Calculate bridging interest per month from the loan amount and monthly rate
        const bridgingInterestRateMonthlyValue = document.getElementById(`${calculatorType}_bridging_interest_rate_monthly`)?.value || '1 %';
        const bridgingInterestRateMonthly = parseFloat(bridgingInterestRateMonthlyValue.replace(/[%\s]/g, '')) || 1;
        const bridgingInterestPerMonth = initialMortgage * (bridgingInterestRateMonthly / 100);
        bridgingCost = bridgingInterestPerMonth * vacantPeriod;
    }
    
    // Add expenses during refurb to total investment (including bridging cost if applicable)
    // Only include if refurb is enabled
    let expensesDuringRefurbFinal = 0;
    if (isRefurbEnabled) {
        const expensesDuringRefurb = mortgagePaymentsDuringRefurb + refurbCouncilTaxDuringRefurb + refurbElectricGas + refurbWater + refurbInsurance + additionalAnnualExpensesDuringRefurb + additionalMonthlyExpensesDuringRefurb;
        // For bridging finance, use bridging cost instead of mortgage payments during refurb
        expensesDuringRefurbFinal = financingType === 'bridging' ? bridgingCost + refurbCouncilTaxDuringRefurb + refurbElectricGas + refurbWater + refurbInsurance + additionalAnnualExpensesDuringRefurb + additionalMonthlyExpensesDuringRefurb : expensesDuringRefurb;
    }
    totalInvestment += expensesDuringRefurbFinal;
    
    // Log expenses during refurb and updated total investment
    console.log('  expensesDuringRefurb:', expensesDuringRefurbFinal);
    console.log('  bridgingCost:', bridgingCost);
    console.log('  totalInvestment (with refurb expenses):', totalInvestment);
    
    // Calculate money left in (now that totalInvestment is complete)
    lockedInEquity = estimatedMarketValue - refinanceAmount;
    moneyBack = refinanceAmount - initialMortgage;
    moneyLeftIn = totalInvestment - moneyBack;
    
    // Log money back and money left in
    console.log('  moneyBack:', moneyBack);
    console.log('  moneyLeftIn:', moneyLeftIn);
    
    // Calculate ongoing costs
    let annualMaintenance = 0;
    let annualAgentFees = 0;
    let annualInsurance = 0;
    let annualCouncilTax = 0;
    let annualTVLicense = 0;
    let annualUtilities = 0;
    let annualWater = 0;
    let annualBroadband = 0;
    let annualBookingFees = 0;
    let annualCleaningCosts = 0;
    
    if (calculatorType === 'holiday-let') {
        // Holiday Let specific calculations
        // Maintenance (as % of income)
        const maintenancePercentValue = document.getElementById(`${calculatorType}_maintenance_percent`)?.value || '10 %';
        const maintenancePercent = parseFloat(maintenancePercentValue.replace(/[%\s]/g, '')) || 10;
        annualMaintenance = annualRent * (maintenancePercent / 100);
        
        // Council Tax (annual)
        const councilTaxValue = document.getElementById(`${calculatorType}_council_tax`)?.value || '£ 1,670';
        annualCouncilTax = parseCurrency(councilTaxValue);
        
        // TV License (if checked)
        const tvLicenseCheckbox = document.getElementById(`${calculatorType}_tv_license`);
        if (tvLicenseCheckbox && tvLicenseCheckbox.checked) {
            annualTVLicense = 159; // Standard UK TV license fee
        }
        
        // Electric / Gas (monthly, convert to annual)
        const utilitiesValue = document.getElementById(`${calculatorType}_utilities`)?.value || '£ 140';
        const monthlyUtilities = parseCurrency(utilitiesValue);
        annualUtilities = monthlyUtilities * 12;
        
        // Water (monthly, convert to annual)
        const waterValue = document.getElementById(`${calculatorType}_water`)?.value || '£ 40';
        const monthlyWater = parseCurrency(waterValue);
        annualWater = monthlyWater * 12;
        
        // Broadband / TV (monthly, convert to annual)
        const broadbandValue = document.getElementById(`${calculatorType}_broadband_tv`)?.value || '£ 60';
        const monthlyBroadband = parseCurrency(broadbandValue);
        annualBroadband = monthlyBroadband * 12;
        
        // Insurance (monthly, convert to annual)
        const ongoingInsuranceValue = document.getElementById(`${calculatorType}_ongoing_insurance`)?.value || '£ 40';
        const monthlyInsurance = parseCurrency(ongoingInsuranceValue);
        annualInsurance = monthlyInsurance * 12;
        
        // Agent Fees (as % of income or fixed amount)
        const agentFeesValue = document.getElementById(`${calculatorType}_agent_fees`)?.value || '0 %';
        const agentFeesPercent = parseFloat(agentFeesValue.replace(/[%\s]/g, '')) || 0;
        // Check if it's a percentage (<= 100) or fixed amount
        if (agentFeesPercent <= 100) {
            annualAgentFees = annualRent * (agentFeesPercent / 100);
        } else {
            // Fixed monthly amount
            annualAgentFees = agentFeesPercent * 12;
        }
        
        // Booking Fees (as % of income)
        const bookingFeesValue = document.getElementById(`${calculatorType}_booking_fees`)?.value || '12 %';
        const bookingFeesPercent = parseFloat(bookingFeesValue.replace(/[%\s]/g, '')) || 12;
        if (bookingFeesPercent <= 100) {
            annualBookingFees = annualRent * (bookingFeesPercent / 100);
        } else {
            // Fixed monthly amount
            annualBookingFees = bookingFeesPercent * 12;
        }
        
        // Cleaning Costs (monthly, convert to annual)
        const cleaningCostsValue = document.getElementById(`${calculatorType}_cleaning_costs`)?.value || '£ 80';
        const monthlyCleaningCosts = parseCurrency(cleaningCostsValue);
        annualCleaningCosts = monthlyCleaningCosts * 12;
    } else {
        // BRR and other calculators
        // Maintenance (as % of income)
        const maintenancePercentValue = document.getElementById(`${calculatorType}_maintenance_percent`)?.value || '10 %';
        const maintenancePercent = parseFloat(maintenancePercentValue.replace(/[%\s]/g, '')) || 10;
        annualMaintenance = annualRent * (maintenancePercent / 100);
        
        // Agent fees (as % of income)
        const agentFeesValue = document.getElementById(`${calculatorType}_agent_fees`)?.value || '10 %';
        const agentFeesPercent = parseFloat(agentFeesValue.replace(/[%\s]/g, '')) || 10;
        annualAgentFees = annualRent * (agentFeesPercent / 100);
        
        // Insurance (monthly, convert to annual)
        const ongoingInsuranceValue = document.getElementById(`${calculatorType}_ongoing_insurance`)?.value || '£ 40';
        const monthlyInsurance = parseCurrency(ongoingInsuranceValue);
        annualInsurance = monthlyInsurance * 12;
    }
    
    // Ongoing mortgage payments (annual)
    const ongoingMortgagePaymentsValue = document.getElementById(`${calculatorType}_ongoing_mortgage_payments`)?.value || '£ 0';
    const monthlyOngoingMortgage = parseCurrency(ongoingMortgagePaymentsValue);
    const annualOngoingMortgage = monthlyOngoingMortgage * 12;
    
    // Additional annual expenses
    let additionalAnnualExpenses = 0;
    const additionalAnnualExpensesInputs = document.querySelectorAll(`[id^="${calculatorType}_additional_annual_expense_"]`);
    additionalAnnualExpensesInputs.forEach(input => {
        const value = parseCurrency(input.value || '£ 0');
        additionalAnnualExpenses += value;
    });
    
    // Additional monthly expenses (convert to annual)
    let additionalMonthlyExpenses = 0;
    const additionalMonthlyExpensesInputs = document.querySelectorAll(`[id^="${calculatorType}_additional_monthly_expense_"]`);
    additionalMonthlyExpensesInputs.forEach(input => {
        const value = parseCurrency(input.value || '£ 0');
        additionalMonthlyExpenses += value;
    });
    const annualAdditionalMonthlyExpenses = additionalMonthlyExpenses * 12;
    
    // Total annual expenses
    // Use actual mortgage payments (monthly * 12) which includes both interest and principal for repayment mortgages
    // For interest-only, this equals the interest
    // NOTE: bridgingCost is NOT included here - it's a one-time cost during vacant period, not an ongoing annual expense
    const annualMortgagePayments = monthlyMortgagePaymentRefi * 12;
    const totalAnnualExpenses = annualMortgagePayments + annualMaintenance + annualAgentFees + annualInsurance + 
                                annualCouncilTax + annualTVLicense + annualUtilities + annualWater + annualBroadband + 
                                annualBookingFees + annualCleaningCosts + 
                                additionalAnnualExpenses + annualAdditionalMonthlyExpenses;
    
    // Debug logging for Holiday Let expenses
    if (calculatorType === 'holiday-let') {
        console.log('💰 Holiday Let Annual Expenses Breakdown:', {
            annualMortgagePayments,
            annualMaintenance,
            annualAgentFees,
            annualInsurance,
            annualCouncilTax,
            annualTVLicense,
            annualUtilities,
            annualWater,
            annualBroadband,
            annualBookingFees,
            annualCleaningCosts,
            additionalAnnualExpenses,
            annualAdditionalMonthlyExpenses,
            totalAnnualExpenses,
            annualRent
        });
    }
    
    // Update ongoing mortgage payments field with refinance mortgage payments
    updateElement(`${calculatorType}_ongoing_mortgage_payments`, formatCurrency(monthlyMortgagePaymentRefi));
    
    const annualProfit = annualRent - totalAnnualExpenses;
    const monthlyProfit = annualProfit / 12;
    
    // ROI calculations
    // Show "Infinite" if moneyLeftIn is negative (you're getting money back)
    let roi = 0;
    if (moneyLeftIn > 0) {
        roi = (annualProfit / moneyLeftIn) * 100;
    } else if (moneyLeftIn < 0 && annualProfit > 0) {
        roi = Infinity; // Infinite ROI when you get money back
    }
    
    // ROCE - Return on Capital Employed
    // ROCE should use the actual capital employed after refinance (money left in)
    // If moneyLeftIn is negative (you get money back), ROCE is infinite
    let roce = 0;
    if (moneyLeftIn > 0) {
        roce = (annualProfit / moneyLeftIn) * 100;
    } else if (moneyLeftIn < 0 && annualProfit > 0) {
        roce = Infinity; // Infinite ROCE when you get money back
    } else if (moneyLeftIn <= 0 && totalInvestment > 0) {
        // Fallback to totalInvestment if moneyLeftIn is 0 or negative
        roce = (annualProfit / totalInvestment) * 100;
    }
    const netYield = estimatedMarketValue > 0 ? (annualProfit / estimatedMarketValue) * 100 : 0;
    
    // Equity in 10 years
    // PropertyEngine shows the mortgage balance as the initial mortgage balance in the table
    // However, for equity calculation, we should use the refinance mortgage balance
    // But the original PropertyEngine appears to use initialMortgage for the table display
    // Equity in 10 years
    // PropertyEngine uses the INITIAL mortgage balance for equity calculation (not refinance amount)
    // This matches their table display which shows the initial mortgage balance for all years
    const futureValue = estimatedMarketValue * Math.pow(1 + appreciation / 100, 10);
    
    // Use initial mortgage balance (not refinance amount) for equity calculation
    // For interest-only: balance stays at initialMortgage
    // For repayment: calculate remaining balance after 10 years based on initial mortgage
    let mortgageBalance10Years = initialMortgage; // Default for interest-only
    
    // Check if initial mortgage is repayment type (reuse variables already calculated earlier)
    if (isInitialRepayment && initialMortgageRate > 0 && initialMortgage > 0) {
        // Calculate remaining balance after 10 years for repayment mortgage
        const monthlyRate = initialMortgageRate / 100 / 12;
        const totalMonths = mortgageTermYearsInitial * 12;
        const monthsPaid = 10 * 12; // 10 years
        
        if (totalMonths > monthsPaid) {
            // Remaining balance = P * [(1+r)^n - (1+r)^p] / [(1+r)^n - 1]
            // Where P = principal, r = monthly rate, n = total months, p = months paid
            const balanceFactor = (Math.pow(1 + monthlyRate, totalMonths) - Math.pow(1 + monthlyRate, monthsPaid)) / (Math.pow(1 + monthlyRate, totalMonths) - 1);
            mortgageBalance10Years = initialMortgage * balanceFactor;
        } else {
            mortgageBalance10Years = 0; // Fully paid off
        }
    }
    
    const equity10Years = futureValue - mortgageBalance10Years;
    
    // Calculate Ideal Purchase Price
    // This is the maximum purchase price where moneyLeftIn would be 0 (break-even)
    // moneyLeftIn = totalInvestment - (refinanceAmount - initialMortgage)
    // For break-even: totalInvestment = refinanceAmount - initialMortgage
    // Simplifying: purchasePrice + stampDuty + refurbCost + surveyCosts + legalFees + mortgageSetupFee = refinanceAmount
    // But stamp duty and mortgageSetupFee depend on purchase price, so we need to iterate
    
    if (refinanceAmount > 0) {
        // Get mortgage LTV for initial financing
        const mortgageLTV = financingType === 'mortgage' ? 
            parseFloat(document.getElementById(`${calculatorType}_mortgage_ltv`)?.value?.replace(/[%\s]/g, '') || '75') :
            (financingType === 'bridging' ? 
                parseFloat(document.getElementById(`${calculatorType}_bridging_ltv`)?.value?.replace(/[%\s]/g, '') || '75') : 0);
        
        // Determine if additional property for stamp duty calculation
        let isAdditionalPropertyForIdeal = calculatorType !== 'purchase';
        if (individualMoving) {
            isAdditionalPropertyForIdeal = false;
        }
        
        // Iterative calculation to find ideal purchase price
        // For break-even: totalInvestment = refinanceAmount - initialMortgage
        // Simplifying: purchasePrice + stampDuty + refurbCost + surveyCosts + legalFees + setupFee = refinanceAmount
        // Start with a reasonable estimate (refinanceAmount minus fixed costs, ignoring stamp duty for now)
        let testPrice = refinanceAmount - refurbCost - surveyCosts - legalFees - mortgageSetupFee - 1000; // Subtract 1000 as rough estimate for stamp duty
        // Ensure it's reasonable
        if (testPrice < purchasePrice * 0.1) testPrice = purchasePrice * 0.5;
        if (testPrice > estimatedMarketValue) testPrice = estimatedMarketValue * 0.8;
        let iterations = 0;
        const maxIterations = 50;
        let lastTestPrice = testPrice;
        let bestPrice = testPrice;
        let bestMoneyLeftIn = Infinity;
        let testMoneyLeftIn = Infinity; // Initialize to avoid ReferenceError
        
        while (iterations < maxIterations) {
            // Calculate stamp duty for test price
            const testStampDutyBreakdown = [];
            const testStampDuty = calculateStampDuty(testPrice, stampDutyPeriod, individualMoving, firstTimeBuyer, overseasBuyer, testStampDutyBreakdown, isAdditionalPropertyForIdeal);
            
            // Calculate initial mortgage for test price
            const testInitialMortgage = testPrice * (mortgageLTV / 100);
            
            // Calculate setup fee for test price
            let testSetupFee = 0;
            if (financingType === 'mortgage') {
                const mortgageSetupFeeValue = document.getElementById(`${calculatorType}_mortgage_setup_fee`)?.value || '£ 1,000';
                if (mortgageSetupFeeValue.includes('%')) {
                    const percent = parseFloat(mortgageSetupFeeValue.replace(/[%\s]/g, '')) || 0;
                    testSetupFee = testInitialMortgage * (percent / 100);
                } else {
                    testSetupFee = parseCurrency(mortgageSetupFeeValue);
                }
            } else if (financingType === 'bridging') {
                const bridgingSetupFeeValue = document.getElementById(`${calculatorType}_bridging_setup_fee`)?.value || '1.75 %';
                if (bridgingSetupFeeValue.includes('%')) {
                    const percent = parseFloat(bridgingSetupFeeValue.replace(/[%\s]/g, '')) || 0;
                    testSetupFee = testInitialMortgage * (percent / 100);
                } else {
                    testSetupFee = parseCurrency(bridgingSetupFeeValue);
                }
            }
            
            // Calculate refinance setup fee for test price (if it's a percentage)
            let testRefinanceSetupFee = 0;
            const refinanceSetupFeeValueForTest = document.getElementById(`${calculatorType}_refinance_setup_fee`)?.value || '£ 0';
            if (refinanceSetupFeeValueForTest.includes('%')) {
                const percent = parseFloat(refinanceSetupFeeValueForTest.replace(/[%\s]/g, '')) || 0;
                testRefinanceSetupFee = refinanceAmount * (percent / 100);
            } else {
                testRefinanceSetupFee = parseCurrency(refinanceSetupFeeValueForTest);
            }
            
            // Calculate mortgage payments during vacant period for test price
            let testMortgagePaymentsDuringRefurb = 0;
            if (financingType === 'mortgage' && vacantPeriod > 0) {
                // Calculate monthly mortgage payment for test price
                let testMonthlyMortgagePayment = 0;
                if (testInitialMortgage > 0) {
                    const monthlyRateTest = (initialMortgageRate / 100) / 12;
                    const numberOfPaymentsTest = mortgageTermYearsInitial * 12;
                    if (isInitialRepayment && monthlyRateTest > 0) {
                        testMonthlyMortgagePayment = testInitialMortgage * (monthlyRateTest * Math.pow(1 + monthlyRateTest, numberOfPaymentsTest)) /
                                                    (Math.pow(1 + monthlyRateTest, numberOfPaymentsTest) - 1);
                    } else if (isInitialRepayment && monthlyRateTest === 0) {
                        testMonthlyMortgagePayment = testInitialMortgage / numberOfPaymentsTest;
                    } else {
                        testMonthlyMortgagePayment = testInitialMortgage * (initialMortgageRate / 100) / 12;
                    }
                }
                testMortgagePaymentsDuringRefurb = testMonthlyMortgagePayment * vacantPeriod;
            }
            
            // Calculate expenses during refurb for test price (same as main calculation)
            const testRefurbCouncilTaxDuringRefurb = (refurbCouncilTaxAnnual / 12) * vacantPeriod;
            const testRefurbElectricGas = parseCurrency(refurbElectricGasValue) * vacantPeriod;
            const testRefurbWater = parseCurrency(refurbWaterValue) * vacantPeriod;
            const testRefurbInsurance = parseCurrency(refurbInsuranceValue) * vacantPeriod;
            const testExpensesDuringRefurb = testMortgagePaymentsDuringRefurb + testRefurbCouncilTaxDuringRefurb + testRefurbElectricGas + testRefurbWater + testRefurbInsurance;
            
            // Calculate total investment for test price (must match the same calculation as main totalInvestment)
            const testDeposit = testPrice - testInitialMortgage;
            const testTotalInvestment = testDeposit + testStampDuty + refurbCost + surveyCosts + legalFees + testSetupFee + testRefinanceSetupFee + testExpensesDuringRefurb;
            
            // Calculate money left in for test price
            testMoneyLeftIn = testTotalInvestment - (refinanceAmount - testInitialMortgage);
            
            // Track the best (closest to zero) result
            if (Math.abs(testMoneyLeftIn) < Math.abs(bestMoneyLeftIn)) {
                bestPrice = testPrice;
                bestMoneyLeftIn = testMoneyLeftIn;
            }
            
            // If we're very close to zero, we found it
            if (Math.abs(testMoneyLeftIn) < 0.01) {
                idealPurchasePrice = testPrice;
                break;
            }
            
            // Adjust test price based on money left in
            // If moneyLeftIn is positive, totalInvestment > moneyBack, so we need a LOWER purchase price
            // If moneyLeftIn is negative, totalInvestment < moneyBack, so we need a HIGHER purchase price
            // Use binary search approach for more stable convergence
            if (testMoneyLeftIn > 0) {
                // Too high, need to reduce
                testPrice = testPrice * 0.9;
            } else {
                // Too low, need to increase
                testPrice = testPrice * 1.1;
            }
            
            // Alternative: use linear adjustment with damping
            // const adjustmentFactor = 0.2;
            // const adjustment = testMoneyLeftIn * adjustmentFactor;
            // testPrice = testPrice - adjustment;
            
            // Check for convergence
            if (Math.abs(testPrice - lastTestPrice) < 0.01 && iterations > 10) {
                idealPurchasePrice = bestPrice; // Use the best price we found
                break;
            }
            
            // Ensure test price is reasonable
            if (testPrice < 0) testPrice = purchasePrice * 0.1;
            if (testPrice > estimatedMarketValue * 1.2) testPrice = estimatedMarketValue * 0.9;
            
            lastTestPrice = testPrice;
            iterations++;
        }
        
        // Use the best price found, or the final test price if we didn't converge
        idealPurchasePrice = Math.abs(bestMoneyLeftIn) < Math.abs(testMoneyLeftIn) ? bestPrice : testPrice;
        
        // Ensure it's positive and reasonable
        if (idealPurchasePrice < 0 || isNaN(idealPurchasePrice) || !isFinite(idealPurchasePrice)) {
            idealPurchasePrice = purchasePrice; // Fallback to current purchase price
        }
        if (idealPurchasePrice > estimatedMarketValue * 1.5) {
            idealPurchasePrice = estimatedMarketValue * 0.9; // Cap at 90% of market value
        }
    }
    
    // Update UI
    updateElement(`${calculatorType}_stamp_duty`, formatCurrency(stampDuty));
    updateElement(`${calculatorType}_total_investment`, formatCurrency(totalInvestment));
    updateElement(`${calculatorType}_estimated_market_value`, formatCurrency(estimatedMarketValue));
    
    // Update mortgage payments - update both the general field and the detailed field
    updateElement(`${calculatorType}_mortgage_payments`, formatCurrency(monthlyMortgagePaymentInitial));
    updateElement(`${calculatorType}_mortgage_payments_detailed`, formatCurrency(monthlyMortgagePaymentInitial));
    
    // Also update the financing payment field based on financing type
    const financingTypeHidden = document.getElementById(`${calculatorType}_financing_type_hidden`);
    const currentFinancingType = financingTypeHidden?.value || 'bridging';
    if (currentFinancingType === 'mortgage') {
        // When showing mortgage payments, update the bridging_interest field (which is reused for mortgage payments)
        updateElement(`${calculatorType}_bridging_interest`, formatCurrency(monthlyMortgagePaymentInitial));
    } else if (currentFinancingType === 'bridging') {
        // Calculate bridging interest per month: bridging loan amount * monthly interest rate
        const bridgingInterestRateMonthlyValue = document.getElementById(`${calculatorType}_bridging_interest_rate_monthly`)?.value || '1 %';
        const bridgingInterestRateMonthly = parseFloat(bridgingInterestRateMonthlyValue.replace(/[%\s]/g, '')) || 1;
        const bridgingInterestPerMonth = initialMortgage * (bridgingInterestRateMonthly / 100);
        updateElement(`${calculatorType}_bridging_interest`, formatCurrency(bridgingInterestPerMonth));
    }
    
    // Update Mortgage Required / Finance Required field based on financing type
    updateElement(`${calculatorType}_mortgage_required`, formatCurrency(initialMortgage));
    updateElement(`${calculatorType}_bridging_finance_required`, formatCurrency(initialMortgage));
    
    updateElement(`${calculatorType}_refinance_mortgage_payments`, formatCurrency(monthlyMortgagePaymentRefi));
    updateElement(`${calculatorType}_locked_in_equity`, formatCurrency(lockedInEquity));
    updateElement(`${calculatorType}_money_left_in`, formatCurrency(moneyLeftIn));
    updateElement(`${calculatorType}_ideal_purchase_price`, formatCurrency(idealPurchasePrice));
    updateElement(`${calculatorType}_gross_yield`, grossYield > 0 ? grossYield.toFixed(2) + '%' : '-');
    updateElement(`${calculatorType}_total_annual_expenses`, formatCurrency(totalAnnualExpenses));
    updateElement(`${calculatorType}_annual_profit`, formatCurrency(annualProfit));
    updateElement(`${calculatorType}_monthly_profit`, formatCurrency(monthlyProfit));
    updateElement(`${calculatorType}_roi`, roi === Infinity ? 'Infinite' : (roi > 0 ? roi.toFixed(2) + '%' : '-%'));
    updateElement(`${calculatorType}_roce`, roce === Infinity ? 'Infinite' : (roce > 0 ? roce.toFixed(2) + '%' : '-%'));
    updateElement(`${calculatorType}_gross_yield_metric`, grossYield > 0 ? grossYield.toFixed(2) + '%' : '-%');
    updateElement(`${calculatorType}_net_yield`, netYield > 0 ? netYield.toFixed(2) + '%' : '-%');
    updateElement(`${calculatorType}_equity_10_years`, formatCurrency(equity10Years));
}

function updateElement(id, value) {
    const el = document.getElementById(id);
    if (!el) return;
    
    // Don't update input/textarea fields that are currently being edited (have focus)
    // This prevents overwriting user input while they're typing
    if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') {
        // NEVER update the refinance interest rate field - it's user-editable
        // Also check for mortgage interest rate fields
        if (id.includes('refinance_interest_rate') || 
            id.includes('mortgage_interest_rate') ||
            id.endsWith('_interest_rate')) {
            // These are user-editable fields - NEVER overwrite them
            return;
        }
        // Only update if the element doesn't have focus (user isn't currently editing it)
        if (document.activeElement !== el) {
            el.value = value;
        }
    } else {
        el.textContent = value;
    }
}

function parseCurrency(value) {
    if (!value) return 0;
    const cleaned = String(value).replace(/[£,\s%]/g, '');
    return parseFloat(cleaned) || 0;
}

function formatCurrency(value) {
    return `£${value.toLocaleString('en-GB', { minimumFractionDigits: 0, maximumFractionDigits: 0 })}`;
}

function showCalculatorFields(selectedCalculators) {
    console.log('[Frontend] showCalculatorFields called with:', selectedCalculators);
    const fieldsContainer = document.getElementById('calculator-fields-list');
    const calculatorSection = document.getElementById('calculator-fields-container');
    const standardFields = document.getElementById('standard-fields');
    const rentalFields = document.getElementById('rental-fields');
    
    // Hide old static fields
    if (standardFields) standardFields.style.display = 'none';
    if (rentalFields) rentalFields.style.display = 'none';
    
    // Clear container
    fieldsContainer.innerHTML = '';
    
    if (selectedCalculators.length === 0) {
        calculatorSection.style.display = 'none';
        return;
    }
    
    // Show fields for each selected calculator
    selectedCalculators.forEach(calculatorType => {
        console.log('[Frontend] Processing calculator:', calculatorType);
        const config = calculatorConfigs[calculatorType];
        if (!config) {
            console.warn('[Frontend] No config found for calculator:', calculatorType);
            return;
        }
        
        // Special handling for Standard Buy to Let calculator - use simple view
        if (calculatorType === 'standard-btl') {
            console.log('[Frontend] Creating Simple Buy to Let calculator section');
            const btlSection = createSimpleBuyToLetCalculator(calculatorType);
            fieldsContainer.appendChild(btlSection);
            
            // Initial calculation
            setTimeout(() => {
                calculateSimpleBuyToLetValues(calculatorType);
            }, 100);
        }
        // Special handling for BRR calculator - use PropertyEngine style
        else if (calculatorType === 'brr') {
            console.log('[Frontend] Creating BRR calculator section');
            const brrSection = createBRRCalculator(calculatorType);
            fieldsContainer.appendChild(brrSection);
            
            // Populate with mock data if available
            if (window.mockDataStore) {
                setTimeout(() => {
                    const mockData = window.mockDataStore;
                    const setValue = (id, value) => {
                        const el = document.getElementById(id);
                        if (el && value) el.value = value;
                    };
                    
                    setValue(`${calculatorType}_purchase_price`, mockData.purchase_price);
                    setValue(`${calculatorType}_refurb_cost`, mockData.refurb_cost);
                    setValue(`${calculatorType}_estimated_market_value`, mockData.after_refurb_value || '£350,000');
                    setValue(`${calculatorType}_monthly_rent`, mockData.monthly_rent);
                    setValue(`${calculatorType}_bridging_interest`, mockData.bridging_interest);
                    setValue(`${calculatorType}_vacant_period`, mockData.vacant_period);
                    
                    calculateBRRValues(calculatorType);
                }, 200);
            }
        } else if (calculatorType === 'holiday-let') {
            // Special handling for Holiday Let calculator - use PropertyEngine style (similar to BRR)
            console.log('[Frontend] Creating Holiday Let calculator section');
            const holidayLetSection = createHolidayLetCalculator(calculatorType);
            console.log('[Frontend] Holiday Let section created, appending to container');
            fieldsContainer.appendChild(holidayLetSection);
            console.log('[Frontend] Holiday Let section appended, checking event handlers...');
            
            // Verify event handlers are set up
            setTimeout(() => {
                // Try multiple ways to find the section
                const sectionById = document.getElementById(`${calculatorType}_purchase_price`)?.closest('[data-calculator]');
                const sectionByQuery = document.querySelector(`[data-calculator="${calculatorType}"]`);
                const section = sectionById || sectionByQuery || holidayLetSection;
                
                const detailedViewToggleById = document.getElementById(`${calculatorType}_detailed_view`);
                const detailedViewToggle = detailedViewToggleById || section?.querySelector(`#${calculatorType}_detailed_view`);
                
                console.log('[Frontend] Holiday Let event handler check:', {
                    sectionExists: !!section,
                    sectionById: !!sectionById,
                    sectionByQuery: !!sectionByQuery,
                    sectionDirect: !!holidayLetSection,
                    detailedViewToggleExists: !!detailedViewToggle,
                    detailedViewToggleById: !!detailedViewToggleById,
                    toggleId: `${calculatorType}_detailed_view`,
                    hasEventListeners: detailedViewToggle ? (detailedViewToggle.onchange !== null || detailedViewToggle.getAttribute('listener') === 'true') : false,
                    allInputsInSection: section ? section.querySelectorAll('input').length : 0
                });
            }, 200);
            
            // Populate with mock data if available
            if (window.mockDataStore) {
                setTimeout(() => {
                    const mockData = window.mockDataStore;
                    const setValue = (id, value) => {
                        const el = document.getElementById(id);
                        if (el && value) el.value = value;
                    };
                    
                    setValue(`${calculatorType}_purchase_price`, mockData.purchase_price);
                    setValue(`${calculatorType}_refurb_cost`, mockData.refurb_cost);
                    setValue(`${calculatorType}_estimated_market_value`, mockData.after_refurb_value || '£350,000');
                    // Holiday Let specific: nightly rate instead of monthly rent
                    // setValue(`${calculatorType}_nightly_rate`, mockData.nightly_rate);
                    // setValue(`${calculatorType}_occupancy_rate`, mockData.occupancy_rate || '70');
                    
                    // Use BRR calculation function for now (can create separate one later)
                    if (typeof calculateBRRValues === 'function') {
                        calculateBRRValues(calculatorType);
                    }
                }, 200);
            }
        } else if (calculatorType === 'rent-to-hmo') {
            // Special handling for Rent to HMO calculator
            console.log('[Frontend] Creating Rent to HMO calculator section');
            const rentToHMOSection = createRentToHMOCalculator(calculatorType);
            fieldsContainer.appendChild(rentToHMOSection);
            
            // Initial calculation
            setTimeout(() => {
                calculateRentToHMOValues(calculatorType);
            }, 100);
        } else {
            // Standard calculator rendering for other types
            const section = document.createElement('div');
            section.className = 'calculator-fields-section';
            section.dataset.calculator = calculatorType;
            
            const title = document.createElement('h3');
            title.textContent = config.title;
            section.appendChild(title);
            
            const fieldsGrid = document.createElement('div');
            fieldsGrid.className = 'form-grid';
            
            config.fields.forEach(field => {
                // Create unique ID for this calculator's field
                const uniqueId = `${calculatorType}_${field.id}`;
                
                const label = document.createElement('label');
                label.textContent = field.label + (field.required ? ' *' : '');
                label.setAttribute('for', uniqueId);
                
                let input;
                if (field.type === 'currency') {
                    input = document.createElement('input');
                    input.type = 'text';
                    input.id = uniqueId;
                    input.dataset.originalId = field.id;
                    input.dataset.calculator = calculatorType;
                    input.placeholder = 'e.g., £1,000';
                    input.pattern = '[£0-9,.]*';
                } else if (field.type === 'number') {
                    input = document.createElement('input');
                    input.type = 'number';
                    input.id = uniqueId;
                    input.dataset.originalId = field.id;
                    input.dataset.calculator = calculatorType;
                    input.step = '0.01';
                    input.placeholder = 'e.g., 20';
                } else {
                    input = document.createElement('input');
                    input.type = 'text';
                    input.id = uniqueId;
                    input.dataset.originalId = field.id;
                    input.dataset.calculator = calculatorType;
                }
                
                if (field.required) {
                    input.required = true;
                }
                
                fieldsGrid.appendChild(label);
                fieldsGrid.appendChild(input);
            });
            
            section.appendChild(fieldsGrid);
            fieldsContainer.appendChild(section);
            
            // Populate with mock data
            if (window.mockDataStore) {
                const calculatorOverrides = {
                    'holiday-let': {
                        'occupancy_rate': '65',
                        'management_fee': '20'
                    },
                    'rent-to-hmo': {
                        'occupancy_rate': '80',
                        'management_fee': ''
                    },
                    'rent-to-serviced': {
                        'occupancy_rate': '60',
                        'management_fee': '18'
                    },
                    'flip': {
                        // Flip-specific overrides if needed
                    }
                };
                
                const overrides = calculatorOverrides[calculatorType] || {};
                
                config.fields.forEach(field => {
                    const uniqueId = `${calculatorType}_${field.id}`;
                    const input = document.getElementById(uniqueId);
                    if (input) {
                        // First check calculator-specific overrides
                        let value = overrides[field.id];
                        
                        // If no override, get from mock data store
                        if (value === undefined || value === '') {
                            value = window.mockDataStore[field.id];
                        }
                        
                        // Set the value if it exists
                        if (value !== undefined && value !== null && value !== '') {
                            input.value = value;
                        }
                    }
                });
            }
        }
    });
    
    // Show calculator fields section
    calculatorSection.style.display = 'block';
}

// Initialize calculator selection on page load
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializeCalculatorSelection);
} else {
    initializeCalculatorSelection();
}

// Make functions available globally for onclick handlers
window.generatePDF = generatePDFFile;
window.generatePDFFile = generatePDFFile; // Also expose with full name
window.fetchLocationData = fetchLocationData;

