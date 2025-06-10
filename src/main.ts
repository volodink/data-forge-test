import { DataFrame, IDataFrame } from 'data-forge';
import * as XLSX from 'xlsx';

// DOM elements
const app = document.getElementById('app') as HTMLDivElement;
const fileInput = document.createElement('input');
const tableContainer = document.createElement('div');
const loadingIndicator = document.createElement('p');

// Configure file input
fileInput.type = 'file';
fileInput.accept = '.xlsx,.xls';
fileInput.style.marginBottom = '20px';
fileInput.id = 'excel-upload';

// Configure loading indicator
loadingIndicator.id = 'loading';
loadingIndicator.style.display = 'none';
loadingIndicator.textContent = 'Loading data...';

// Set up UI
app.innerHTML = '<h1>Excel Data Viewer</h1>';
app.appendChild(fileInput);
app.appendChild(loadingIndicator);
app.appendChild(tableContainer);

// Handle file selection
fileInput.addEventListener('change', handleFileUpload);

async function handleFileUpload(event: Event) {
    const input = event.target as HTMLInputElement;
    if (!input.files?.length) return;

    try {
        // Show loading indicator
        loadingIndicator.style.display = 'block';
        tableContainer.innerHTML = '';
        
        const file = input.files[0];
        const data = await parseExcel(file);
        displayData(data);
    } catch (error) {
        console.error('Error processing file:', error);
        tableContainer.innerHTML = `<p class="error">Error: ${(error as Error).message}</p>`;
    } finally {
        // Hide loading indicator
        loadingIndicator.style.display = 'none';
    }
}

async function parseExcel(file: File): Promise<IDataFrame> {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = (e) => {
            try {
                const buffer = e.target?.result as ArrayBuffer;
                if (!buffer) {
                    throw new Error('Failed to read file content');
                }
                
                const workbook = XLSX.read(buffer, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const jsonData: Record<string, unknown>[] = XLSX.utils.sheet_to_json(worksheet);
                
                // Create DataFrame
                let df: IDataFrame = new DataFrame(jsonData);
                
                // Add calculated column if Age exists
                if (df.getColumnNames().includes('Age')) {
                    df = df.generateSeries({
                        AgeGroup: (row: any) => row.Age >= 30 ? '30+' : 'Under 30'
                    });
                }
                
                resolve(df);
            } catch (error) {
                reject(error);
            }
        };
        
        reader.onerror = () => {
            reject(new Error('Failed to read file'));
        };
        
        reader.readAsArrayBuffer(file);
    });
}

function displayData(df: IDataFrame) {
    const columnNames = df.getColumnNames();
    const data = df.toArray();

    // Create table
    const table = document.createElement('table');
    
    // Create header
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    columnNames.forEach(col => {
        const th = document.createElement('th');
        th.textContent = col;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);
    
    // Create body
    const tbody = document.createElement('tbody');
    data.forEach(row => {
        const tr = document.createElement('tr');
        columnNames.forEach(col => {
            const td = document.createElement('td');
            td.textContent = row[col]?.toString() || '';
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    
    // Clear container and append table
    tableContainer.innerHTML = '';
    tableContainer.appendChild(table);
}

// Optional: Add sample file link
const sampleLink = document.createElement('a');
sampleLink.href = '/sample-data.xlsx';
sampleLink.textContent = 'Download Sample Excel File';
sampleLink.style.display = 'block';
sampleLink.style.marginTop = '10px';
app.appendChild(sampleLink);
