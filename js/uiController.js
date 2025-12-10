/**
 * UI Controller Module
 * Handles all UI updates and user interactions
 */

/**
 * Updates the KPI dashboard with calculated metrics
 * @param {Object} kpis - Object containing KPI metrics (total, matched, mismatched, percentage)
 */
export function updateDashboard(kpis) {
    document.getElementById('kpi-total').textContent = kpis.total;
    document.getElementById('kpi-matched').textContent = kpis.matched;
    document.getElementById('kpi-mismatched').textContent = kpis.mismatched;
    document.getElementById('kpi-percentage').textContent = kpis.percentage + '%';
}

/**
 * Renders the comparison results table
 * @param {Array} results - Array of comparison result objects
 * @param {string} filter - Filter type: 'all', 'matched', or 'mismatched'
 */
export function renderTable(results, filter = 'all') {
    const tableBody = document.getElementById('table-body');
    tableBody.innerHTML = '';
    
    // Filter results based on selected filter
    let filteredResults = results;
    if (filter === 'matched') {
        filteredResults = results.filter(r => r.status === 'matched');
    } else if (filter === 'mismatched') {
        filteredResults = results.filter(r => r.status === 'mismatched');
    }
    
    // Create table rows
    filteredResults.forEach(result => {
        const row = document.createElement('tr');
        
        // Add row class based on status
        row.className = result.status === 'matched' ? 'row-matched' : 'row-mismatched';
        
        // Invoice Number
        const cellNumber = document.createElement('td');
        cellNumber.textContent = result.invoiceNumber;
        row.appendChild(cellNumber);
        
        // Date
        const cellDate = document.createElement('td');
        cellDate.textContent = result.date;
        row.appendChild(cellDate);
        
        // Customer
        const cellCustomer = document.createElement('td');
        cellCustomer.textContent = result.customer;
        row.appendChild(cellCustomer);
        
        // Invoice Amount
        const cellInvoiceAmount = document.createElement('td');
        cellInvoiceAmount.textContent = formatCurrency(result.invoiceAmount);
        // Highlight if mismatched
        if (result.status === 'mismatched' && result.invoiceAmount !== 0) {
            cellInvoiceAmount.innerHTML = `<span class="highlight-mismatch">${formatCurrency(result.invoiceAmount)}</span>`;
        }
        row.appendChild(cellInvoiceAmount);
        
        // Accounting Amount
        const cellAccountingAmount = document.createElement('td');
        cellAccountingAmount.textContent = formatCurrency(result.accountingAmount);
        // Highlight if mismatched
        if (result.status === 'mismatched' && result.accountingAmount !== 0) {
            cellAccountingAmount.innerHTML = `<span class="highlight-mismatch">${formatCurrency(result.accountingAmount)}</span>`;
        }
        row.appendChild(cellAccountingAmount);
        
        // Difference
        const cellDifference = document.createElement('td');
        const diffText = formatCurrency(Math.abs(result.difference));
        if (result.difference > 0) {
            cellDifference.textContent = '+' + diffText;
        } else if (result.difference < 0) {
            cellDifference.textContent = '-' + diffText;
        } else {
            cellDifference.textContent = diffText;
        }
        if (result.status === 'mismatched') {
            cellDifference.innerHTML = `<span class="highlight-mismatch">${cellDifference.textContent}</span>`;
        }
        row.appendChild(cellDifference);
        
        // Status
        const cellStatus = document.createElement('td');
        const statusBadge = document.createElement('span');
        statusBadge.className = `status-badge ${result.status === 'matched' ? 'status-matched' : 'status-mismatched'}`;
        statusBadge.textContent = result.status === 'matched' ? 'OK' : 'Discrepanza';
        cellStatus.appendChild(statusBadge);
        row.appendChild(cellStatus);
        
        tableBody.appendChild(row);
    });
    
    // Show "no results" message if table is empty
    if (filteredResults.length === 0) {
        const row = document.createElement('tr');
        const cell = document.createElement('td');
        cell.colSpan = 7;
        cell.textContent = 'Nessun risultato da visualizzare';
        cell.style.textAlign = 'center';
        cell.style.padding = '2rem';
        cell.style.color = 'var(--text-secondary)';
        row.appendChild(cell);
        tableBody.appendChild(row);
    }
}

/**
 * Formats a number as currency (EUR)
 * @param {number} amount - Amount to format
 * @returns {string} Formatted currency string
 */
function formatCurrency(amount) {
    return new Intl.NumberFormat('it-IT', {
        style: 'currency',
        currency: 'EUR',
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    }).format(amount);
}

/**
 * Shows the results section
 */
export function showResults() {
    document.getElementById('results-section').style.display = 'block';
    // Scroll to results section
    document.getElementById('results-section').scrollIntoView({ 
        behavior: 'smooth',
        block: 'start'
    });
}

/**
 * Updates file info display when a file is selected
 * @param {string} elementId - ID of the file info element
 * @param {File} file - Selected file object
 */
export function updateFileInfo(elementId, file) {
    const fileInfoElement = document.getElementById(elementId);
    if (file) {
        const size = (file.size / 1024).toFixed(2);
        fileInfoElement.textContent = `✓ ${file.name} (${size} KB)`;
        fileInfoElement.style.color = 'var(--success-color)';
    } else {
        fileInfoElement.textContent = '';
    }
}

/**
 * Shows an error message to the user
 * @param {string} message - Error message to display
 */
export function showError(message) {
    alert('Errore: ' + message);
}

/**
 * Shows a success message to the user
 * @param {string} message - Success message to display
 */
export function showSuccess(message) {
    alert('Successo: ' + message);
}
