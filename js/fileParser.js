/**
 * File Parser Module
 * Handles parsing of XML and CSV files containing invoice and accounting data
 */

/**
 * Parses CSV file content into an array of objects
 * Expected CSV format: invoice_number,date,customer,amount
 * @param {string} content - CSV file content as text
 * @returns {Array} Array of invoice objects
 */
export function parseCSV(content) {
    const lines = content.trim().split('\n');
    const invoices = [];
    
    // Skip header row (first line)
    for (let i = 1; i < lines.length; i++) {
        const line = lines[i].trim();
        if (!line) continue;
        
        // Split by comma, handling quoted values
        const values = line.match(/(".*?"|[^",\s]+)(?=\s*,|\s*$)/g) || [];
        const cleanValues = values.map(v => v.replace(/^"|"$/g, '').trim());
        
        if (cleanValues.length >= 4) {
            invoices.push({
                invoiceNumber: cleanValues[0],
                date: cleanValues[1],
                customer: cleanValues[2],
                amount: parseFloat(cleanValues[3]) || 0
            });
        }
    }
    
    return invoices;
}

/**
 * Parses XML file content into an array of objects
 * Supports multiple XML formats for invoice data
 * @param {string} content - XML file content as text
 * @returns {Array} Array of invoice objects
 */
export function parseXML(content) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(content, 'text/xml');
    
    // Check for parsing errors
    const parserError = xmlDoc.querySelector('parsererror');
    if (parserError) {
        throw new Error('Errore nel parsing XML: ' + parserError.textContent);
    }
    
    const invoices = [];
    
    // Try different possible XML structures
    // Structure 1: <invoices><invoice>...</invoice></invoices>
    let invoiceElements = xmlDoc.querySelectorAll('invoice, fattura');
    
    // Structure 2: Direct root elements
    if (invoiceElements.length === 0) {
        const root = xmlDoc.documentElement;
        if (root.tagName.toLowerCase() === 'invoice' || root.tagName.toLowerCase() === 'fattura') {
            invoiceElements = [root];
        }
    }
    
    invoiceElements.forEach(element => {
        // Extract data using multiple possible tag names
        const invoiceNumber = getElementText(element, [
            'invoiceNumber', 'invoice_number', 'numero', 'numero_fattura', 'id'
        ]);
        
        const date = getElementText(element, [
            'date', 'data', 'invoice_date', 'data_fattura'
        ]);
        
        const customer = getElementText(element, [
            'customer', 'cliente', 'customer_name', 'nome_cliente'
        ]);
        
        const amountText = getElementText(element, [
            'amount', 'importo', 'total', 'totale', 'total_amount'
        ]);
        
        if (invoiceNumber && amountText) {
            invoices.push({
                invoiceNumber: invoiceNumber,
                date: date || 'N/A',
                customer: customer || 'N/A',
                amount: parseFloat(amountText.replace(',', '.')) || 0
            });
        }
    });
    
    return invoices;
}

/**
 * Helper function to get text content from element using multiple possible tag names
 * @param {Element} element - Parent XML element
 * @param {Array<string>} tagNames - Array of possible tag names to search
 * @returns {string} Text content or empty string
 */
function getElementText(element, tagNames) {
    for (const tagName of tagNames) {
        const child = element.querySelector(tagName);
        if (child && child.textContent) {
            return child.textContent.trim();
        }
        
        // Also check as attribute
        const attr = element.getAttribute(tagName);
        if (attr) {
            return attr.trim();
        }
    }
    return '';
}

/**
 * Main function to parse file based on its type
 * @param {File} file - File object to parse
 * @returns {Promise<Array>} Promise resolving to array of invoice objects
 */
export async function parseFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = (e) => {
            try {
                const content = e.target.result;
                let invoices;
                
                // Determine file type and parse accordingly
                if (file.name.toLowerCase().endsWith('.csv')) {
                    invoices = parseCSV(content);
                } else if (file.name.toLowerCase().endsWith('.xml')) {
                    invoices = parseXML(content);
                } else {
                    throw new Error('Formato file non supportato. Utilizzare CSV o XML.');
                }
                
                resolve(invoices);
            } catch (error) {
                reject(error);
            }
        };
        
        reader.onerror = () => {
            reject(new Error('Errore nella lettura del file'));
        };
        
        reader.readAsText(file);
    });
}
