// Export functionality for Costing Validation results

/**
 * Export the results table as a PDF
 */
async function exportToPDF() {
    const table = document.getElementById('mainResultsTable');

    if (!table) {
        alert('No results to export. Please generate results first.');
        return;
    }

    try {
        // Show loading state
        const exportBtn = document.querySelector('.export-btn');
        const originalText = exportBtn.textContent;
        exportBtn.textContent = 'Exporting...';
        exportBtn.disabled = true;

        // Import jsPDF library dynamically if not already loaded
        if (typeof window.jspdf === 'undefined') {
            await loadJsPDF();
        }

        const { jsPDF } = window.jspdf;
        const doc = new jsPDF('l', 'mm', 'a4'); // Landscape orientation for wide table

        // Add title
        doc.setFontSize(18);
        doc.setFont(undefined, 'bold');
        doc.text('Costing Validation Results', 14, 15);

        // Add timestamp
        doc.setFontSize(10);
        doc.setFont(undefined, 'normal');
        const timestamp = new Date().toLocaleString();
        doc.text(`Generated: ${timestamp}`, 14, 22);

        // Get summary information
        const summaryDiv = document.querySelector('div[style*="background: #f0f7ff"]');
        let summaryHeight = 28;

        if (summaryDiv) {
            const summaryText = summaryDiv.textContent.trim();
            doc.setFontSize(9);

            // Split the summary text into lines to calculate proper height
            const lines = doc.splitTextToSize(summaryText, 260);
            doc.text(lines, 14, 28);

            // Calculate the height needed for the summary (each line is ~4mm)
            summaryHeight = 28 + (lines.length * 4) + 5; // Add 5mm padding
        }

        // Prepare table data
        const tableData = extractTableData(table);

        // Add table using autoTable plugin
        doc.autoTable({
            head: tableData.headers,
            body: tableData.rows,
            startY: summaryHeight,
            styles: {
                fontSize: 7,
                cellPadding: 2,
                overflow: 'linebreak',
                cellWidth: 'wrap'
            },
            headStyles: {
                fillColor: [35, 62, 92], // Madison88 primary color
                textColor: [255, 255, 255],
                fontStyle: 'bold',
                halign: 'center'
            },
            columnStyles: {
                0: { cellWidth: 35 }, // OB File/s
                1: { cellWidth: 30 }, // Buyer CBD File/s
                2: { cellWidth: 25 }, // Match Status
                3: { cellWidth: 30 }, // Standard Minute Value
                4: { cellWidth: 25 }, // Average Efficiency
                5: { cellWidth: 30 }, // Hourly Wages
                6: { cellWidth: 35 }, // Overhead Cost
                7: { cellWidth: 25 }  // Factory Profit
            },
            alternateRowStyles: {
                fillColor: [245, 245, 245]
            },
            margin: { top: 10, right: 10, bottom: 10, left: 10 },
            didParseCell: function (data) {
                // Color code the Match Status column (index 2)
                if (data.column.index === 2 && data.section === 'body') {
                    const cellText = data.cell.text[0];
                    if (cellText && cellText.includes('✓ FOUND')) {
                        data.cell.styles.textColor = [6, 95, 70]; // Green
                        data.cell.styles.fontStyle = 'bold';
                    } else if (cellText && cellText.includes('✗ NOT FOUND')) {
                        data.cell.styles.textColor = [153, 27, 27]; // Red
                        data.cell.styles.fontStyle = 'bold';
                    }
                }

                // Color code the Standard Minute Value column (index 3)
                if (data.column.index === 3 && data.section === 'body') {
                    const cellText = data.cell.text[0];
                    // Check if the cell contains "Product:" and "OB Total SMV:" which indicates a mismatch
                    if (cellText && cellText.includes('Product:') && cellText.includes('OB Total SMV:')) {
                        // Extract the difference value to determine color
                        // Format is like: "Product: 1.234 OB Total SMV: 1.230 (+0.004)" or "(-0.004)"
                        const diffMatch = cellText.match(/\([\+\-]([\d.]+)\)/);

                        if (diffMatch) {
                            const difference = parseFloat(diffMatch[1]);

                            // Apply color based on difference magnitude (same logic as website)
                            if (difference <= 0.01) {
                                // Small difference (0.001 to 0.01) - orange
                                data.cell.styles.textColor = [217, 119, 6]; // Orange
                            } else {
                                // Larger difference (> 0.01) - red
                                data.cell.styles.textColor = [153, 27, 27]; // Red
                            }
                            data.cell.styles.fontStyle = 'bold';
                        } else {
                            // Fallback: if we can't parse the difference, use orange
                            data.cell.styles.textColor = [217, 119, 6]; // Orange
                            data.cell.styles.fontStyle = 'bold';
                        }
                    } else if (cellText && (cellText.includes('Empty') || cellText.includes('TNF: Empty'))) {
                        // Empty cell - color it red
                        data.cell.styles.textColor = [153, 27, 27]; // Red
                        data.cell.styles.fontStyle = 'bold';
                    } else if (cellText && cellText !== '-' && !cellText.includes('Product:')) {
                        // If it's just a number without "Product:" prefix, it's a match - color it green
                        data.cell.styles.textColor = [6, 95, 70]; // Green
                        data.cell.styles.fontStyle = 'bold';
                    }
                }

                // Color code the Average Efficiency % column (index 4)
                // Expected value: 50.0%
                if (data.column.index === 4 && data.section === 'body') {
                    const cellText = data.cell.text[0];
                    if (cellText && cellText.includes('Cell Empty')) {
                        data.cell.styles.textColor = [153, 27, 27]; // Red
                        data.cell.styles.fontStyle = 'bold';
                    } else if (cellText && cellText !== '-') {
                        // Extract the percentage value
                        const match = cellText.match(/([\d.]+)%/);
                        if (match) {
                            const value = parseFloat(match[1]);
                            // Check if it's not 50.0% (with small tolerance)
                            if (Math.abs(value - 50.0) >= 0.1) {
                                data.cell.styles.textColor = [217, 119, 6]; // Orange
                                data.cell.styles.fontStyle = 'bold';
                            } else {
                                // Correct value - green
                                data.cell.styles.textColor = [6, 95, 70]; // Green
                                data.cell.styles.fontStyle = 'bold';
                            }
                        }
                    }
                }

                // Color code the Hourly Wages with Fringes column (index 5)
                // Expected value: 1.750
                if (data.column.index === 5 && data.section === 'body') {
                    const cellText = data.cell.text[0];
                    if (cellText && cellText.includes('Cell Empty')) {
                        data.cell.styles.textColor = [153, 27, 27]; // Red
                        data.cell.styles.fontStyle = 'bold';
                    } else if (cellText && cellText !== '-') {
                        // Extract the numeric value
                        const match = cellText.match(/([\d.]+)/);
                        if (match) {
                            const value = parseFloat(match[1]);
                            // Check if it's not 1.750 (with small tolerance)
                            if (Math.abs(value - 1.750) >= 0.01) {
                                data.cell.styles.textColor = [217, 119, 6]; // Orange
                                data.cell.styles.fontStyle = 'bold';
                            } else {
                                // Correct value - green
                                data.cell.styles.textColor = [6, 95, 70]; // Green
                                data.cell.styles.fontStyle = 'bold';
                            }
                        }
                    }
                }

                // Color code the Overhead Cost Ratio to Direct Labor column (index 6)
                // Expected value: 70.0%
                if (data.column.index === 6 && data.section === 'body') {
                    const cellText = data.cell.text[0];
                    if (cellText && cellText.includes('Cell Empty')) {
                        data.cell.styles.textColor = [153, 27, 27]; // Red
                        data.cell.styles.fontStyle = 'bold';
                    } else if (cellText && cellText !== '-') {
                        // Extract the percentage value
                        const match = cellText.match(/([\d.]+)%/);
                        if (match) {
                            const value = parseFloat(match[1]);
                            // Check if it's not 70.0% (with small tolerance)
                            if (Math.abs(value - 70.0) >= 0.1) {
                                data.cell.styles.textColor = [217, 119, 6]; // Orange
                                data.cell.styles.fontStyle = 'bold';
                            } else {
                                // Correct value - green
                                data.cell.styles.textColor = [6, 95, 70]; // Green
                                data.cell.styles.fontStyle = 'bold';
                            }
                        }
                    }
                }

                // Color code the Factory Profit % column (index 7)
                // Expected value: 10.0%
                if (data.column.index === 7 && data.section === 'body') {
                    const cellText = data.cell.text[0];
                    if (cellText && cellText.includes('Cell Empty')) {
                        data.cell.styles.textColor = [153, 27, 27]; // Red
                        data.cell.styles.fontStyle = 'bold';
                    } else if (cellText && cellText !== '-') {
                        // Extract the percentage value
                        const match = cellText.match(/([\d.]+)%/);
                        if (match) {
                            const value = parseFloat(match[1]);
                            // Check if it's not 10.0% (with small tolerance)
                            if (Math.abs(value - 10.0) >= 0.1) {
                                data.cell.styles.textColor = [217, 119, 6]; // Orange
                                data.cell.styles.fontStyle = 'bold';
                            } else {
                                // Correct value - green
                                data.cell.styles.textColor = [6, 95, 70]; // Green
                                data.cell.styles.fontStyle = 'bold';
                            }
                        }
                    }
                }
            }
        });

        // Add page numbers
        const pageCount = doc.internal.getNumberOfPages();
        for (let i = 1; i <= pageCount; i++) {
            doc.setPage(i);
            doc.setFontSize(8);
            doc.text(
                `Page ${i} of ${pageCount}`,
                doc.internal.pageSize.getWidth() / 2,
                doc.internal.pageSize.getHeight() - 10,
                { align: 'center' }
            );
        }

        // Generate filename with date and time
        const now = new Date();
        const date = now.toISOString().slice(0, 10); // YYYY-MM-DD
        const time = now.toTimeString().slice(0, 8).replace(/:/g, '-'); // HH-MM-SS
        const filename = `CostingValidation_${date}.pdf`;

        // Save the PDF
        doc.save(filename);

        // Reset button state
        exportBtn.textContent = originalText;
        exportBtn.disabled = false;

        console.log('PDF exported successfully:', filename);

    } catch (error) {
        console.error('Error exporting PDF:', error);
        alert('Failed to export PDF. Please try again.');

        // Reset button state
        const exportBtn = document.querySelector('.export-btn');
        exportBtn.textContent = 'Export';
        exportBtn.disabled = false;
    }
}

/**
 * Extract table data from the HTML table
 */
function extractTableData(table) {
    const headers = [];
    const rows = [];

    // Extract headers from the first header row (not the filter row)
    const headerRow = table.querySelector('thead tr.header-labels-row');
    if (headerRow) {
        const headerCells = headerRow.querySelectorAll('th');
        headerCells.forEach(cell => {
            headers.push(cell.textContent.trim());
        });
    }

    // Extract visible rows from tbody
    const tbody = table.querySelector('tbody');
    const bodyRows = tbody.querySelectorAll('tr');

    bodyRows.forEach(row => {
        // Skip hidden rows (filtered out)
        if (row.style.display === 'none') {
            return;
        }

        const rowData = [];
        const cells = row.querySelectorAll('td');

        cells.forEach((cell, index) => {
            let cellText = '';

            // Extract text content, handling special formatting
            if (index === 0 || index === 1) {
                // For OB File and Buyer CBD columns, extract main text and details
                const strong = cell.querySelector('strong');
                const details = cell.querySelectorAll('.match-details');

                if (strong) {
                    cellText = strong.textContent.trim();
                    if (details.length > 0) {
                        details.forEach(detail => {
                            cellText += '\n' + detail.textContent.trim();
                        });
                    }
                } else {
                    cellText = cell.textContent.trim();
                }
            } else if (index === 2) {
                // Match Status - extract just the status text
                const statusSpan = cell.querySelector('.match-status');
                cellText = statusSpan ? statusSpan.textContent.trim() : cell.textContent.trim();
            } else {
                // For other columns, extract all text including spans
                // Remove HTML and get clean text
                const spans = cell.querySelectorAll('span');
                if (spans.length > 0) {
                    const textParts = [];
                    spans.forEach(span => {
                        const text = span.textContent.trim();
                        if (text && !text.includes('Expected:')) {
                            textParts.push(text);
                        }
                    });
                    cellText = textParts.join(' ');
                } else {
                    cellText = cell.textContent.trim();
                }

                // Clean up extra whitespace
                cellText = cellText.replace(/\s+/g, ' ').trim();
            }

            rowData.push(cellText);
        });

        rows.push(rowData);
    });

    return {
        headers: [headers],
        rows: rows
    };
}

/**
 * Load jsPDF library dynamically
 */
function loadJsPDF() {
    return new Promise((resolve, reject) => {
        // Check if already loaded
        if (typeof window.jspdf !== 'undefined') {
            resolve();
            return;
        }

        // Load jsPDF
        const jsPDFScript = document.createElement('script');
        jsPDFScript.src = 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
        jsPDFScript.onload = () => {
            // Load autoTable plugin
            const autoTableScript = document.createElement('script');
            autoTableScript.src = 'https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.31/jspdf.plugin.autotable.min.js';
            autoTableScript.onload = () => {
                console.log('jsPDF and autoTable loaded successfully');
                resolve();
            };
            autoTableScript.onerror = () => {
                reject(new Error('Failed to load jsPDF autoTable plugin'));
            };
            document.head.appendChild(autoTableScript);
        };
        jsPDFScript.onerror = () => {
            reject(new Error('Failed to load jsPDF library'));
        };
        document.head.appendChild(jsPDFScript);
    });
}

/**
 * Main export function - can be extended to support multiple formats
 */
function exportResults() {
    // For now, only PDF export is implemented
    exportToPDF();
}
