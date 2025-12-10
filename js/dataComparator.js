/**
 * Data Comparator Module
 * Handles comparison between invoice data and accounting data
 */

/**
 * Compares invoice data with accounting data
 * Matches invoices by invoice number and compares amounts
 * @param {Array} invoiceData - Array of invoice objects from ADE
 * @param {Array} accountingData - Array of accounting objects from gestionale
 * @returns {Array} Array of comparison result objects
 */
export function compareData(invoiceData, accountingData) {
    // Create a map of accounting data for efficient lookup
    const accountingMap = new Map();
    accountingData.forEach(item => {
        accountingMap.set(item.invoiceNumber, item);
    });
    
    const results = [];
    
    // Compare each invoice with accounting data
    invoiceData.forEach(invoice => {
        const accounting = accountingMap.get(invoice.invoiceNumber);
        
        if (accounting) {
            // Invoice found in accounting data - compare amounts
            const invoiceAmount = parseFloat(invoice.amount) || 0;
            const accountingAmount = parseFloat(accounting.amount) || 0;
            const difference = Math.abs(invoiceAmount - accountingAmount);
            
            // Consider amounts matched if difference is less than 0.01 (rounding tolerance)
            const isMatched = difference < 0.01;
            
            results.push({
                invoiceNumber: invoice.invoiceNumber,
                date: invoice.date,
                customer: invoice.customer,
                invoiceAmount: invoiceAmount,
                accountingAmount: accountingAmount,
                difference: invoiceAmount - accountingAmount,
                status: isMatched ? 'matched' : 'mismatched',
                found: true
            });
        } else {
            // Invoice not found in accounting data
            results.push({
                invoiceNumber: invoice.invoiceNumber,
                date: invoice.date,
                customer: invoice.customer,
                invoiceAmount: parseFloat(invoice.amount) || 0,
                accountingAmount: 0,
                difference: parseFloat(invoice.amount) || 0,
                status: 'mismatched',
                found: false
            });
        }
    });
    
    // Check for accounting entries not in invoice data
    accountingData.forEach(accounting => {
        const invoiceExists = invoiceData.some(
            inv => inv.invoiceNumber === accounting.invoiceNumber
        );
        
        if (!invoiceExists) {
            results.push({
                invoiceNumber: accounting.invoiceNumber,
                date: accounting.date,
                customer: accounting.customer,
                invoiceAmount: 0,
                accountingAmount: parseFloat(accounting.amount) || 0,
                difference: -(parseFloat(accounting.amount) || 0),
                status: 'mismatched',
                found: false
            });
        }
    });
    
    return results;
}

/**
 * Calculates KPI metrics from comparison results
 * @param {Array} comparisonResults - Array of comparison result objects
 * @returns {Object} Object containing KPI metrics
 */
export function calculateKPIs(comparisonResults) {
    const total = comparisonResults.length;
    const matched = comparisonResults.filter(r => r.status === 'matched').length;
    const mismatched = total - matched;
    const percentage = total > 0 ? Math.round((matched / total) * 100) : 0;
    
    return {
        total,
        matched,
        mismatched,
        percentage
    };
}
