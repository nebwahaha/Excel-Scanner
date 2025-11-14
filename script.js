let tnfFiles = [];  // Array of TNF OB Excel files
let productFiles = []; // Array of Product ID files

// Initialize drag and drop for left zone (TNF OB)
const leftDropZone = document.getElementById('leftDropZone');
const leftFileInput = document.getElementById('leftFileInput');
const leftFileList = document.getElementById('leftFileList');

// Initialize drag and drop for right zone (Product ID)
const rightDropZone = document.getElementById('rightDropZone');
const rightFileInput = document.getElementById('rightFileInput');
const rightFileList = document.getElementById('rightFileList');

// Generate button
const generateBtn = document.getElementById('generateBtn');

// Results table container
const resultsTable = document.getElementById('resultsTable');

// Check if XLSX library is loaded
function checkXLSXLoaded() {
    if (typeof XLSX === 'undefined') {
        alert('Excel library failed to load. Please refresh the page and try again.');
        return false;
    }
    return true;
}

// Update file list display
function updateFileList(files, listElement, isLeft) {
    listElement.innerHTML = '';
    files.forEach((file, index) => {
        const fileItem = document.createElement('div');
        fileItem.className = 'file-item';
        fileItem.innerHTML = `
            <span class="file-item-name">✓ ${file.name}</span>
            <button class="file-item-remove" onclick="removeFile(${index}, ${isLeft})">Remove</button>
        `;
        listElement.appendChild(fileItem);
    });
}

// Remove file from list
function removeFile(index, isLeft) {
    if (isLeft) {
        tnfFiles.splice(index, 1);
        updateFileList(tnfFiles, leftFileList, true);
    } else {
        productFiles.splice(index, 1);
        updateFileList(productFiles, rightFileList, false);
    }
}

// LEFT DROP ZONE EVENTS (TNF OB)
leftDropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    e.stopPropagation();
    leftDropZone.classList.add('dragover');
});

leftDropZone.addEventListener('dragenter', (e) => {
    e.preventDefault();
    e.stopPropagation();
    leftDropZone.classList.add('dragover');
});

leftDropZone.addEventListener('dragleave', (e) => {
    e.preventDefault();
    e.stopPropagation();
    // Only remove if leaving the drop zone itself, not child elements
    if (e.target === leftDropZone) {
        leftDropZone.classList.remove('dragover');
    }
});

leftDropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    e.stopPropagation();
    leftDropZone.classList.remove('dragover');

    const files = Array.from(e.dataTransfer.files).filter(file =>
        file.name.endsWith('.xlsx') || file.name.endsWith('.xls')
    );

    if (files.length > 0) {
        tnfFiles = [...tnfFiles, ...files];
        updateFileList(tnfFiles, leftFileList, true);
    } else {
        alert('Please drop Excel files (.xlsx or .xls)');
    }
});

leftFileInput.addEventListener('change', (e) => {
    const files = Array.from(e.target.files);
    tnfFiles = [...tnfFiles, ...files];
    updateFileList(tnfFiles, leftFileList, true);
    e.target.value = ''; // Reset input to allow same file again
});

// RIGHT DROP ZONE EVENTS (Product ID)
rightDropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    e.stopPropagation();
    rightDropZone.classList.add('dragover');
});

rightDropZone.addEventListener('dragenter', (e) => {
    e.preventDefault();
    e.stopPropagation();
    rightDropZone.classList.add('dragover');
});

rightDropZone.addEventListener('dragleave', (e) => {
    e.preventDefault();
    e.stopPropagation();
    // Only remove if leaving the drop zone itself, not child elements
    if (e.target === rightDropZone) {
        rightDropZone.classList.remove('dragover');
    }
});

rightDropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    e.stopPropagation();
    rightDropZone.classList.remove('dragover');

    const files = Array.from(e.dataTransfer.files).filter(file =>
        file.name.endsWith('.xlsx') || file.name.endsWith('.xls')
    );

    if (files.length > 0) {
        productFiles = [...productFiles, ...files];
        updateFileList(productFiles, rightFileList, false);
    } else {
        alert('Please drop Excel files (.xlsx or .xls)');
    }
});

rightFileInput.addEventListener('change', (e) => {
    const files = Array.from(e.target.files);
    productFiles = [...productFiles, ...files];
    updateFileList(productFiles, rightFileList, false);
    e.target.value = ''; // Reset input to allow same file again
});

// Prevent default drag behaviors on the whole document
document.addEventListener('dragover', (e) => {
    e.preventDefault();
});

document.addEventListener('drop', (e) => {
    e.preventDefault();
});

// GENERATE BUTTON CLICK
generateBtn.addEventListener('click', async () => {
    if (!checkXLSXLoaded()) {
        return;
    }

    resultsTable.innerHTML = '';

    if (tnfFiles.length === 0 || productFiles.length === 0) {
        alert('Please upload at least one OB file and one Buyer CBD file');
        return;
    }

    try {
        // Show loading animation
        resultsTable.innerHTML = `
            <div class="loading-container">
                <div class="loader"></div>
                <div class="loading-text">Searching... Please wait.</div>
            </div>
        `;

        // Extract all product IDs and cell values from PRODUCT files
        const products = [];
        for (const file of productFiles) {
            const productData = await extractProductData(file);
            if (productData.productID) {
                products.push({
                    id: productData.productID,
                    fileName: file.name,
                    cellValues: productData.cellValues
                });
            }
        }

        if (products.length === 0) {
            resultsTable.innerHTML = '<div class="no-results">Error: Could not find any product IDs in the uploaded files</div>';
            return;
        }

        // Search each product in all TNF files
        const allResults = [];
        for (const product of products) {
            for (const tnfFile of tnfFiles) {
                const searchResults = await searchProductInWorkbook(tnfFile, product.id);
                allResults.push({
                    tnfFileName: tnfFile.name,
                    productID: product.id,
                    productFileName: product.fileName,
                    found: searchResults.foundLocations.length > 0,
                    locations: searchResults.foundLocations,
                    cellValues: product.cellValues  // Use cell values from PRODUCT file
                });
            }
        }

        // Display results in table format
        displayResultsTable(allResults);

    } catch (error) {
        resultsTable.innerHTML = `<div class="no-results">Error: ${error.message}</div>`;
    }
});

// Display results in table format
function displayResultsTable(results) {
    if (results.length === 0) {
        resultsTable.innerHTML = `
            <div class="no-results">
                <p style="font-size: 1.3em; margin-bottom: 10px;">❌ No Results</p>
                <p>No Buyer CBD or OB files to search</p>
            </div>
        `;
        return;
    }

    // Group results by product ID to check if product was found anywhere
    const productFoundStatus = {};
    results.forEach(result => {
        if (!productFoundStatus[result.productID]) {
            productFoundStatus[result.productID] = false;
        }
        if (result.found) {
            productFoundStatus[result.productID] = true;
        }
    });

    // Build table - only show "NOT FOUND" for products that weren't found anywhere
    let tableHTML = `
        <table class="results-table">
            <thead>
                <tr>
                    <th>OB File/s</th>
                    <th>Buyer CBD File/s</th>
                    <th>Match Status</th>
                    <th>Standard Minute Value</th>
                    <th>Average Efficiency %</th>
                    <th>Hourly Wages with Fringes</th>
                    <th>Overhead Cost Ratio to Direct Labor</th>
                    <th>Factory Profit %</th>
                </tr>
            </thead>
            <tbody>
    `;

    // First, show all FOUND results
    results.forEach((result) => {
        if (result.found) {
            // Show each location for found products
            result.locations.forEach((location) => {
                // Helper function to format and validate cell values
                const formatCellValue = (value, expectedValue, type) => {
                    if (value === null || value === undefined) {
                        // Show expected value and that cell is empty/not found
                        if (type === 'percentage') {
                            return `<span style="color: #991b1b; font-weight: 600;">Cell Empty</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${expectedValue}%</span>`;
                        } else {
                            return `<span style="color: #991b1b; font-weight: 600;">Cell Empty</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${expectedValue.toFixed(3)}</span>`;
                        }
                    }

                    let numValue = parseFloat(value);
                    let displayValue;
                    let isValid;

                    if (type === 'percentage') {
                        // Handle percentage values (Excel stores as decimal: 0.5 = 50%)
                        if (typeof value === 'string') {
                            numValue = parseFloat(value.replace('%', ''));
                        }

                        // If value is less than 1, it's stored as decimal (0.5 = 50%)
                        if (numValue < 1) {
                            numValue = numValue * 100;
                        }

                        displayValue = numValue.toFixed(1) + '%';
                        isValid = Math.abs(numValue - expectedValue) < 0.1;
                    } else {
                        // Handle numeric values
                        displayValue = numValue.toFixed(3);
                        isValid = Math.abs(numValue - expectedValue) < 0.01;
                    }

                    const color = isValid ? '#065f46' : '#991b1b';
                    const expectedDisplay = type === 'percentage' ? `${expectedValue}%` : expectedValue.toFixed(3);

                    // If valid, just show the value in green
                    if (isValid) {
                        return `<span style="color: ${color}; font-weight: 600;">${displayValue}</span>`;
                    }

                    // If invalid, show actual value in red and expected value below
                    return `<span style="color: ${color}; font-weight: 600;">${displayValue}</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${expectedDisplay}</span>`;
                };

                // Helper function to compare Standard Minute Values
                const formatSMVComparison = (productSMV, tnfSMV) => {
                    // Helper to truncate to 3 decimal places without rounding
                    const truncateToThreeDecimals = (num) => {
                        return Math.floor(num * 1000) / 1000;
                    };

                    // Helper to format number with exactly 3 decimal places without rounding
                    const formatThreeDecimals = (num) => {
                        const truncated = truncateToThreeDecimals(num);
                        const str = truncated.toString();
                        const parts = str.split('.');
                        if (parts.length === 1) {
                            return str + '.000';
                        } else {
                            const decimals = parts[1].padEnd(3, '0');
                            return parts[0] + '.' + decimals;
                        }
                    };

                    if (productSMV === null || productSMV === undefined) {
                        return `<span style="color: #991b1b; font-weight: 600;">Product: Empty</span>`;
                    }
                    if (tnfSMV === null || tnfSMV === undefined) {
                        const formattedProduct = formatThreeDecimals(productSMV);
                        return `<span style="color: #991b1b; font-weight: 600;">TNF: Empty</span><br><span style="font-size: 0.85em; color: #849bba;">Product: ${formattedProduct}</span>`;
                    }

                    const truncatedProduct = truncateToThreeDecimals(productSMV);
                    const truncatedTNF = truncateToThreeDecimals(tnfSMV);

                    const isMatch = Math.abs(truncatedProduct - truncatedTNF) < 0.001;
                    const color = isMatch ? '#065f46' : '#991b1b';
                    const difference = truncateToThreeDecimals(truncatedProduct - truncatedTNF);
                    const diffSign = difference > 0 ? '+' : '';

                    const formattedProduct = formatThreeDecimals(productSMV);
                    const formattedTNF = formatThreeDecimals(tnfSMV);
                    const formattedDiff = formatThreeDecimals(Math.abs(difference));

                    if (isMatch) {
                        return `<span style="color: ${color}; font-weight: 600;">${formattedProduct}</span>`;
                    } else {
                        return `<span style="color: ${color}; font-weight: 600;">Product: ${formattedProduct}</span><br><span style="font-size: 0.85em; color: #849bba;">OB Total SMV: ${formattedTNF} (${diffSign}${formattedDiff})</span>`;
                    }
                };

                tableHTML += `
                    <tr>
                        <td>
                            <strong>${result.tnfFileName}</strong>
                            <div class="match-details">Sheet: ${location.sheet}</div>
                            <div class="match-details">Cell: ${location.cell} (Row ${location.row}, Col ${location.col})</div>
                        </td>
                        <td>
                            <strong>${result.productID}</strong>
                            <div class="match-details">From: ${result.productFileName}</div>
                        </td>
                        <td>
                            <span class="match-status match-found">✓ FOUND</span>
                        </td>
                        <td>${formatSMVComparison(result.cellValues.standardMinuteValue, location.smv)}</td>
                        <td>${formatCellValue(result.cellValues.averageEfficiency, 50, 'percentage')}</td>
                        <td>${formatCellValue(result.cellValues.hourlyWages, 1.750, 'number')}</td>
                        <td>${formatCellValue(result.cellValues.overheadCost, 70, 'percentage')}</td>
                        <td>${formatCellValue(result.cellValues.factoryProfit, 10, 'percentage')}</td>
                    </tr>
                `;
            });
        }
    });

    // Then, show NOT FOUND only for products that weren't found in ANY file
    const notFoundProducts = new Set();
    results.forEach((result) => {
        if (!result.found && !productFoundStatus[result.productID]) {
            // Only add     once per product (not per CBD file)
            if (!notFoundProducts.has(result.productID)) {
                notFoundProducts.add(result.productID);
                tableHTML += `
                    <tr>
                        <td>
                            <em>Searched in all files</em>
                        </td>
                        <td>
                            <strong>${result.productID}</strong>
                            <div class="match-details">From: ${result.productFileName}</div>
                        </td>
                        <td>
                            <span class="match-status match-not-found">✗ NOT FOUND</span>
                        </td>
                        <td colspan="5" style="text-align: center; color: #849bba;">-</td>
                    </tr>
                `;
            }
        }
    });

    tableHTML += `
            </tbody>
        </table>
    `;

    // Add summary
    const matchedResults = results.filter(r => r.found);
    const totalProducts = new Set(results.map(r => r.productID)).size;
    const matchedProducts = new Set(matchedResults.map(r => r.productID)).size;
    const notFoundCount = notFoundProducts.size;
    const totalTNFFiles = new Set(results.map(r => r.tnfFileName)).size;

    const summaryHTML = `
        <div style="margin-bottom: 20px; padding: 15px; background: #f0f7ff; border-radius: 10px; border-left: 4px solid #3b82f6;">
            <strong>Summary:</strong> Found ${matchedProducts} of ${totalProducts} products across ${totalTNFFiles} TNF OB file(s). 
            ${notFoundCount > 0 ? `<span style="color: #991b1b;">${notFoundCount} product(s) not found in any file.</span>` : ''}
            Total matches: ${matchedResults.reduce((sum, r) => sum + r.locations.length, 0)}
        </div>
    `;

    resultsTable.innerHTML = summaryHTML + tableHTML;
}

// Extract Product ID and cell values from the PRODUCT file
async function extractProductData(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });

                // Get first sheet
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

                let productID = null;

                // First, try to extract from cell E14 (common location for Product ID)
                if (worksheet['E14']) {
                    const e14Value = String(worksheet['E14'].v).trim();
                    console.log('Cell E14 value:', e14Value);

                    // Check if E14 contains a valid product ID pattern
                    // Support both patterns: A8HMM style and VN000Q1F style
                    const e14Match = e14Value.match(/^([A-Z]{1,2}\d[A-Z0-9]{3,8})/);
                    if (e14Match) {
                        productID = e14Match[1];
                        console.log('Product ID found in E14:', productID);
                    }
                }

                // If not found in E14, try to extract from filename
                if (!productID) {
                    const fileName = file.name.replace(/\.[^/.]+$/, ''); // Remove extension
                    const fileNameMatch = fileName.match(/^([A-Z]{1,2}\d[A-Z0-9]{3,8})/);

                    if (fileNameMatch) {
                        productID = fileNameMatch[1];
                        console.log('Product ID found in filename:', productID);
                    }
                }

                // If still not found, search for "Style #" or similar pattern in the sheet
                if (!productID) {
                    for (let i = 0; i < Math.min(50, jsonData.length); i++) {
                        const row = jsonData[i];
                        for (let j = 0; j < row.length; j++) {
                            const cellValue = String(row[j]).trim();

                            // Look for "Style #" or "Style No." label
                            if ((cellValue.toLowerCase().includes('style') && cellValue.includes('#')) ||
                                cellValue.toLowerCase().includes('style no')) {
                                // Check next cell or cells nearby for the actual ID
                                if (j + 1 < row.length && row[j + 1]) {
                                    const nextCell = String(row[j + 1]).trim();
                                    const match = nextCell.match(/^([A-Z]{1,2}\d[A-Z0-9]{3,8})/);
                                    if (match) {
                                        productID = match[1];
                                        console.log('Product ID found near Style label:', productID);
                                        break;
                                    }
                                }
                            }

                            // Also check if cell matches pattern directly
                            // Support both A8HMN and VN000Q1F patterns
                            if (/^[A-Z]{1,2}\d[A-Z0-9]{3,8}/.test(cellValue)) {
                                const match = cellValue.match(/^([A-Z]{1,2}\d[A-Z0-9]{3,8})/);
                                if (match) {
                                    productID = match[1];
                                    console.log('Product ID found in cell:', productID);
                                }
                            }
                        }
                        if (productID) break;
                    }
                }

                // Now extract cell values from specific cells in the PRODUCT file
                const cellValues = {
                    standardMinuteValue: null,
                    averageEfficiency: null,
                    hourlyWages: null,
                    overheadCost: null,
                    factoryProfit: null
                };

                // Helper function to extract numeric value from cell
                const extractValue = (cellRef) => {
                    if (!worksheet[cellRef]) {
                        console.log(`Cell ${cellRef} not found in PRODUCT file`);
                        return null;
                    }

                    let value = worksheet[cellRef].v;
                    console.log(`Cell ${cellRef} in PRODUCT file - raw value:`, value, 'Type:', typeof value);

                    // If value is already a number, return it
                    if (typeof value === 'number') {
                        return value;
                    }

                    // If value is a string, try to extract the number
                    if (typeof value === 'string') {
                        // Remove currency symbols, commas, and extract number
                        let cleaned = value.replace(/[$,\s]/g, '');

                        // Try to extract percentage (e.g., "50.0%" -> 50)
                        let percentMatch = cleaned.match(/([\d.]+)%/);
                        if (percentMatch) {
                            return parseFloat(percentMatch[1]);
                        }

                        // Try to extract plain number
                        let numberMatch = cleaned.match(/([\d.]+)/);
                        if (numberMatch) {
                            return parseFloat(numberMatch[1]);
                        }
                    }

                    return null;
                };

                // Extract values from specific cells in the PRODUCT file
                console.log('=== Extracting cell values from PRODUCT file ===');
                console.log('Product ID:', productID);
                console.log('File name:', file.name);

                // K7 - Standard Minute Value
                cellValues.standardMinuteValue = extractValue('K7');

                // K8 - Average Efficiency %
                cellValues.averageEfficiency = extractValue('K8');

                // K9 - Hourly Wages with Fringes
                cellValues.hourlyWages = extractValue('K9');

                // K11 - Overhead Cost Ratio to Direct Labor
                cellValues.overheadCost = extractValue('K11');

                // R5 - Factory Profit %
                cellValues.factoryProfit = extractValue('R5');

                console.log('=== Final extracted values from PRODUCT file ===');
                console.log('Standard Minute Value (K7):', cellValues.standardMinuteValue);
                console.log('Average Efficiency (K8):', cellValues.averageEfficiency);
                console.log('Hourly Wages (K9):', cellValues.hourlyWages);
                console.log('Overhead Cost (K11):', cellValues.overheadCost);
                console.log('Factory Profit (R5):', cellValues.factoryProfit);

                resolve({ productID, cellValues });
            } catch (error) {
                reject(new Error(`Failed to parse product file: ${error.message}`));
            }
        };

        reader.onerror = () => {
            reject(new Error('Failed to read product file'));
        };

        reader.readAsArrayBuffer(file);
    });
}

// Search for product ID across all sheets in the TNF workbook
async function searchProductInWorkbook(file, productID) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });

                let foundLocations = [];

                // Search through each sheet for the product ID and its Total SMV
                workbook.SheetNames.forEach((sheetName) => {
                    const worksheet = workbook.Sheets[sheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

                    // Search for product ID in this sheet
                    for (let rowIndex = 0; rowIndex < jsonData.length; rowIndex++) {
                        const row = jsonData[rowIndex];
                        for (let colIndex = 0; colIndex < row.length; colIndex++) {
                            const cellValue = String(row[colIndex]).trim();

                            // Check if cell contains the product ID
                            if (cellValue === productID || cellValue.includes(productID)) {
                                const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });

                                // Find the SMV value for THIS specific occurrence
                                let smvForThisOccurrence = null;

                                // Search in the next 20 rows for "Total SMV"
                                for (let searchRow = rowIndex; searchRow < Math.min(rowIndex + 20, jsonData.length); searchRow++) {
                                    const searchRowData = jsonData[searchRow];
                                    for (let searchCol = 0; searchCol < searchRowData.length; searchCol++) {
                                        const searchCellValue = String(searchRowData[searchCol]).trim().toLowerCase();

                                        // Look for "Total SMV" label
                                        if (searchCellValue.includes('total smv')) {
                                            // Check the next few cells in the same row for the numeric value
                                            for (let valueCol = searchCol + 1; valueCol < Math.min(searchCol + 5, searchRowData.length); valueCol++) {
                                                const smvCellRef = XLSX.utils.encode_cell({ r: searchRow, c: valueCol });
                                                if (worksheet[smvCellRef]) {
                                                    let smvValue = worksheet[smvCellRef].v;
                                                    if (typeof smvValue === 'number' && smvValue > 0) {
                                                        smvForThisOccurrence = smvValue;
                                                        console.log(`Found Total SMV for ${productID} on sheet ${sheetName}: ${smvForThisOccurrence} at ${smvCellRef}`);
                                                        break;
                                                    }
                                                }
                                            }
                                            if (smvForThisOccurrence !== null) break;
                                        }
                                    }
                                    if (smvForThisOccurrence !== null) break;
                                }

                                foundLocations.push({
                                    sheet: sheetName,
                                    cell: cellAddress,
                                    row: rowIndex + 1,
                                    col: colIndex + 1,
                                    value: cellValue,
                                    smv: smvForThisOccurrence
                                });
                            }
                        }
                    }
                });

                resolve({ foundLocations });
            } catch (error) {
                reject(new Error(`Failed to search workbook: ${error.message}`));
            }
        };

        reader.onerror = () => {
            reject(new Error('Failed to read TNF OB file'));
        };

        reader.readAsArrayBuffer(file);
    });
}

// Dark Mode Toggle
const darkModeToggle = document.getElementById('darkModeToggle');

// Check for saved dark mode preference
const isDarkMode = localStorage.getItem('darkMode') === 'true';
if (isDarkMode) {
    document.body.classList.add('dark-mode');
}

// Toggle dark mode
darkModeToggle.addEventListener('click', () => {
    document.body.classList.toggle('dark-mode');
    const isDark = document.body.classList.contains('dark-mode');
    localStorage.setItem('darkMode', isDark);
});
