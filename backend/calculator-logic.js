// Calculator logic for different investment types

function parseCurrency(value) {
    if (!value) return 0;
    const cleaned = String(value).replace(/[£,\s]/g, '');
    return parseFloat(cleaned) || 0;
}

function formatCurrency(value) {
    return `£${value.toLocaleString('en-GB', { minimumFractionDigits: 0, maximumFractionDigits: 0 })}`;
}

// Standard Buy to Let Calculator
function calculateStandardBTL(data) {
    const purchasePrice = parseCurrency(data.purchase_price);
    const depositPercent = parseFloat(data.deposit_percent) || 20;
    const monthlyRent = parseCurrency(data.monthly_rent);
    const mortgageRate = parseFloat(data.mortgage_rate) || 5.8;
    
    const depositAmount = purchasePrice * (depositPercent / 100);
    const annualRent = monthlyRent * 12;
    const rentalYield = purchasePrice > 0 ? (annualRent / purchasePrice) * 100 : 0;
    
    const stampDuty = parseCurrency(data.stamp_duty);
    const surveyCost = parseCurrency(data.survey_cost);
    const legalFees = parseCurrency(data.legal_fees);
    const loanSetup = parseCurrency(data.loan_setup);
    
    const totalPurchaseCosts = stampDuty + surveyCost + legalFees + loanSetup;
    const totalInvestment = depositAmount + totalPurchaseCosts;
    
    const mortgageAmount = purchasePrice - depositAmount;
    const annualMortgageInterest = mortgageAmount * (mortgageRate / 100);
    
    const councilTax = parseCurrency(data.council_tax);
    const repairs = parseCurrency(data.repairs_maintenance);
    const utilities = parseCurrency(data.utilities);
    const water = parseCurrency(data.water);
    const broadband = parseCurrency(data.broadband_tv);
    const insurance = parseCurrency(data.insurance);
    
    const totalAnnualExpenses = annualMortgageInterest + councilTax + repairs + utilities + water + broadband + insurance;
    const annualProfit = annualRent - totalAnnualExpenses;
    const monthlyProfit = annualProfit / 12;
    const roi = totalInvestment > 0 ? (annualProfit / totalInvestment) * 100 : 0;
    
    return {
        purchasePrice,
        depositAmount,
        mortgageAmount,
        totalInvestment,
        annualRent,
        monthlyRent,
        rentalYield,
        totalAnnualExpenses,
        annualProfit,
        monthlyProfit,
        roi
    };
}

// Buy Refurbish Refinance Calculator
function calculateBRR(data) {
    const purchasePrice = parseCurrency(data.purchase_price);
    const refurbCost = parseCurrency(data.refurb_cost);
    const afterRefurbValue = parseCurrency(data.after_refurb_value);
    const depositPercent = parseFloat(data.deposit_percent) || 20;
    const mortgageRate = parseFloat(data.mortgage_rate) || 5.8;
    const refinanceLTV = parseFloat(data.refinance_ltv) || 75;
    const monthlyRent = parseCurrency(data.monthly_rent);
    
    const depositAmount = purchasePrice * (depositPercent / 100);
    const initialMortgage = purchasePrice - depositAmount;
    const totalInitialInvestment = depositAmount + refurbCost;
    
    const refinanceAmount = afterRefurbValue * (refinanceLTV / 100);
    const moneyBack = refinanceAmount - initialMortgage;
    const netInvestment = totalInitialInvestment - moneyBack;
    
    const annualRent = monthlyRent * 12;
    const rentalYield = afterRefurbValue > 0 ? (annualRent / afterRefurbValue) * 100 : 0;
    
    const annualMortgageInterest = refinanceAmount * (mortgageRate / 100);
    const councilTax = parseCurrency(data.council_tax);
    const repairs = parseCurrency(data.repairs_maintenance);
    const insurance = parseCurrency(data.insurance);
    
    const totalAnnualExpenses = annualMortgageInterest + councilTax + repairs + insurance;
    const annualProfit = annualRent - totalAnnualExpenses;
    const monthlyProfit = annualProfit / 12;
    const roi = netInvestment > 0 ? (annualProfit / netInvestment) * 100 : 0;
    
    return {
        purchasePrice,
        refurbCost,
        afterRefurbValue,
        depositAmount,
        totalInitialInvestment,
        refinanceAmount,
        moneyBack,
        netInvestment,
        annualRent,
        monthlyRent,
        rentalYield,
        totalAnnualExpenses,
        annualProfit,
        monthlyProfit,
        roi
    };
}

// Flip Calculator
function calculateFlip(data) {
    const purchasePrice = parseCurrency(data.purchase_price);
    const refurbCost = parseCurrency(data.refurb_cost);
    const salePrice = parseCurrency(data.sale_price);
    const holdingPeriod = parseFloat(data.holding_period) || 6;
    
    const stampDuty = parseCurrency(data.stamp_duty);
    const surveyCost = parseCurrency(data.survey_cost);
    const legalFeesPurchase = parseCurrency(data.legal_fees);
    const legalFeesSale = parseCurrency(data.legal_fees_sale);
    const estateAgentFees = parseCurrency(data.estate_agent_fees);
    const financeCost = parseCurrency(data.finance_cost);
    
    const totalPurchaseCosts = stampDuty + surveyCost + legalFeesPurchase;
    const totalSaleCosts = legalFeesSale + estateAgentFees;
    const totalInvestment = purchasePrice + refurbCost + totalPurchaseCosts + financeCost;
    
    const grossProfit = salePrice - totalInvestment - totalSaleCosts;
    const netProfit = grossProfit;
    const roi = totalInvestment > 0 ? (netProfit / totalInvestment) * 100 : 0;
    const monthlyROI = holdingPeriod > 0 ? roi / holdingPeriod : 0;
    
    return {
        purchasePrice,
        refurbCost,
        salePrice,
        totalInvestment,
        grossProfit,
        netProfit,
        roi,
        monthlyROI,
        holdingPeriod
    };
}

// Holiday Let Calculator
function calculateHolidayLet(data) {
    const purchasePrice = parseCurrency(data.purchase_price);
    const depositPercent = parseFloat(data.deposit_percent) || 20;
    const mortgageRate = parseFloat(data.mortgage_rate) || 5.8;
    const weeklyRent = parseCurrency(data.weekly_rent);
    const occupancyRate = parseFloat(data.occupancy_rate) || 50;
    const managementFeePercent = parseFloat(data.management_fee) || 20;
    const cleaningFee = parseCurrency(data.cleaning_fee);
    
    const depositAmount = purchasePrice * (depositPercent / 100);
    const mortgageAmount = purchasePrice - depositAmount;
    
    const weeksPerYear = 52;
    const occupiedWeeks = weeksPerYear * (occupancyRate / 100);
    const annualRent = weeklyRent * occupiedWeeks;
    const managementFee = annualRent * (managementFeePercent / 100);
    const totalCleaningFees = cleaningFee * occupiedWeeks;
    
    const annualMortgageInterest = mortgageAmount * (mortgageRate / 100);
    const councilTax = parseCurrency(data.council_tax);
    const utilities = parseCurrency(data.utilities);
    const insurance = parseCurrency(data.insurance);
    
    const totalAnnualExpenses = annualMortgageInterest + managementFee + totalCleaningFees + councilTax + utilities + insurance;
    const annualProfit = annualRent - totalAnnualExpenses;
    const monthlyProfit = annualProfit / 12;
    const rentalYield = purchasePrice > 0 ? (annualRent / purchasePrice) * 100 : 0;
    
    const totalInvestment = depositAmount;
    const roi = totalInvestment > 0 ? (annualProfit / totalInvestment) * 100 : 0;
    
    return {
        purchasePrice,
        depositAmount,
        mortgageAmount,
        annualRent,
        weeklyRent,
        occupancyRate,
        occupiedWeeks,
        totalAnnualExpenses,
        annualProfit,
        monthlyProfit,
        rentalYield,
        roi
    };
}

// Rent to HMO Calculator
function calculateRentToHMO(data) {
    const monthlyRentPaid = parseCurrency(data.monthly_rent_paid);
    const numberOfRooms = parseFloat(data.number_of_rooms) || 1;
    const rentPerRoom = parseCurrency(data.rent_per_room);
    const occupancyRate = parseFloat(data.occupancy_rate) || 80;
    
    const annualRentPaid = monthlyRentPaid * 12;
    const occupiedRooms = Math.floor(numberOfRooms * (occupancyRate / 100));
    const monthlyIncome = rentPerRoom * occupiedRooms;
    const annualIncome = monthlyIncome * 12;
    
    const councilTax = parseCurrency(data.council_tax);
    const utilities = parseCurrency(data.utilities);
    const insurance = parseCurrency(data.insurance);
    const managementFee = parseCurrency(data.management_fee);
    
    const totalAnnualExpenses = annualRentPaid + councilTax + utilities + insurance + managementFee;
    const annualProfit = annualIncome - totalAnnualExpenses;
    const monthlyProfit = annualProfit / 12;
    const roi = annualRentPaid > 0 ? (annualProfit / annualRentPaid) * 100 : 0;
    
    return {
        monthlyRentPaid,
        annualRentPaid,
        numberOfRooms,
        rentPerRoom,
        occupiedRooms,
        monthlyIncome,
        annualIncome,
        totalAnnualExpenses,
        annualProfit,
        monthlyProfit,
        roi
    };
}

// Rent to Serviced Accommodation Calculator
function calculateRentToServiced(data) {
    const monthlyRentPaid = parseCurrency(data.monthly_rent_paid);
    const dailyRate = parseCurrency(data.daily_rate);
    const occupancyRate = parseFloat(data.occupancy_rate) || 60;
    const cleaningFee = parseCurrency(data.cleaning_fee);
    const managementFeePercent = parseFloat(data.management_fee) || 20;
    
    const annualRentPaid = monthlyRentPaid * 12;
    const daysPerYear = 365;
    const occupiedDays = daysPerYear * (occupancyRate / 100);
    const annualIncome = dailyRate * occupiedDays;
    const managementFee = annualIncome * (managementFeePercent / 100);
    const totalCleaningFees = cleaningFee * occupiedDays;
    
    const councilTax = parseCurrency(data.council_tax);
    const utilities = parseCurrency(data.utilities);
    const insurance = parseCurrency(data.insurance);
    
    const totalAnnualExpenses = annualRentPaid + managementFee + totalCleaningFees + councilTax + utilities + insurance;
    const annualProfit = annualIncome - totalAnnualExpenses;
    const monthlyProfit = annualProfit / 12;
    const roi = annualRentPaid > 0 ? (annualProfit / annualRentPaid) * 100 : 0;
    
    return {
        monthlyRentPaid,
        annualRentPaid,
        dailyRate,
        occupancyRate,
        occupiedDays,
        annualIncome,
        totalAnnualExpenses,
        annualProfit,
        monthlyProfit,
        roi
    };
}

// Main calculator function
function calculateInvestment(data) {
    const calculatorType = data.calculator_type || 'standard-btl';
    
    switch (calculatorType) {
        case 'standard-btl':
        case 'purchase':
            return calculateStandardBTL(data);
        case 'brr':
            return calculateBRR(data);
        case 'flip':
            return calculateFlip(data);
        case 'holiday-let':
            return calculateHolidayLet(data);
        case 'rent-to-hmo':
            return calculateRentToHMO(data);
        case 'rent-to-serviced':
            return calculateRentToServiced(data);
        default:
            return calculateStandardBTL(data);
    }
}

// Export for use in other files
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        calculateInvestment,
        calculateStandardBTL,
        calculateBRR,
        calculateFlip,
        calculateHolidayLet,
        calculateRentToHMO,
        calculateRentToServiced,
        parseCurrency,
        formatCurrency
    };
}

