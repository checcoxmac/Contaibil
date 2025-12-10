/**
 * Main Application Module
 * Entry point for the Contaibil application
 * Coordinates file uploads, data parsing, comparison, and UI updates
 */

import { parseFile } from './fileParser.js';
import { compareData, calculateKPIs } from './dataComparator.js';
import { updateDashboard, renderTable, showResults, updateFileInfo, showError } from './uiController.js';

// Application state
const appState = {
    invoiceData: null,
    accountingData: null,
    comparisonResults: null,
    currentFilter: 'all'
};

/**
 * Initializes the application and sets up event listeners
 */
function initializeApp() {
    // Get DOM elements
    const invoiceFileInput = document.getElementById('invoice-file');
    const accountingFileInput = document.getElementById('accounting-file');
    const compareButton = document.getElementById('compare-btn');
    const filterRadios = document.querySelectorAll('input[name="filter"]');
    
    // File upload event listeners
    invoiceFileInput.addEventListener('change', handleInvoiceFileUpload);
    accountingFileInput.addEventListener('change', handleAccountingFileUpload);
    
    // Compare button event listener
    compareButton.addEventListener('click', handleCompareClick);
    
    // Filter radio buttons event listeners
    filterRadios.forEach(radio => {
        radio.addEventListener('change', handleFilterChange);
    });
    
    console.log('Contaibil application initialized');
}

/**
 * Handles invoice file upload
 * @param {Event} event - File input change event
 */
async function handleInvoiceFileUpload(event) {
    const file = event.target.files[0];
    
    if (!file) {
        appState.invoiceData = null;
        updateFileInfo('invoice-info', null);
        updateCompareButtonState();
        return;
    }
    
    try {
        // Show file info
        updateFileInfo('invoice-info', file);
        
        // Parse the file
        appState.invoiceData = await parseFile(file);
        console.log(`Parsed ${appState.invoiceData.length} invoice records`);
        
        // Update button state
        updateCompareButtonState();
    } catch (error) {
        console.error('Error parsing invoice file:', error);
        showError('Errore nel parsing del file fatture: ' + error.message);
        appState.invoiceData = null;
        updateFileInfo('invoice-info', null);
        updateCompareButtonState();
    }
}

/**
 * Handles accounting file upload
 * @param {Event} event - File input change event
 */
async function handleAccountingFileUpload(event) {
    const file = event.target.files[0];
    
    if (!file) {
        appState.accountingData = null;
        updateFileInfo('accounting-info', null);
        updateCompareButtonState();
        return;
    }
    
    try {
        // Show file info
        updateFileInfo('accounting-info', file);
        
        // Parse the file
        appState.accountingData = await parseFile(file);
        console.log(`Parsed ${appState.accountingData.length} accounting records`);
        
        // Update button state
        updateCompareButtonState();
    } catch (error) {
        console.error('Error parsing accounting file:', error);
        showError('Errore nel parsing del file gestionale: ' + error.message);
        appState.accountingData = null;
        updateFileInfo('accounting-info', null);
        updateCompareButtonState();
    }
}

/**
 * Updates the compare button enabled/disabled state
 * Button is enabled only when both files are loaded
 */
function updateCompareButtonState() {
    const compareButton = document.getElementById('compare-btn');
    const isEnabled = appState.invoiceData && appState.accountingData;
    compareButton.disabled = !isEnabled;
}

/**
 * Handles compare button click
 * Performs data comparison and updates UI
 */
function handleCompareClick() {
    if (!appState.invoiceData || !appState.accountingData) {
        showError('Caricare entrambi i file prima di confrontare');
        return;
    }
    
    try {
        // Perform comparison
        appState.comparisonResults = compareData(appState.invoiceData, appState.accountingData);
        console.log(`Comparison complete: ${appState.comparisonResults.length} results`);
        
        // Calculate KPIs
        const kpis = calculateKPIs(appState.comparisonResults);
        
        // Update UI
        updateDashboard(kpis);
        renderTable(appState.comparisonResults, appState.currentFilter);
        showResults();
    } catch (error) {
        console.error('Error during comparison:', error);
        showError('Errore durante il confronto: ' + error.message);
    }
}

/**
 * Handles filter change event
 * Re-renders table with the selected filter
 * @param {Event} event - Radio button change event
 */
function handleFilterChange(event) {
    appState.currentFilter = event.target.value;
    
    if (appState.comparisonResults) {
        renderTable(appState.comparisonResults, appState.currentFilter);
    }
}

// Initialize the application when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializeApp);
} else {
    initializeApp();
}
