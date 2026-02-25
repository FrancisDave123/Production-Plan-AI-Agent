import Papa from 'papaparse';
import ExcelJS from 'exceljs';
import * as mammoth from 'mammoth';
import * as pdfjsLib from 'pdfjs-dist';
import JSZip from 'jszip';
import { ActualDataItem, FileAttachment } from '../types/production';

// Set worker source for PDF.js
// Use unpkg as it mirrors npm versions reliably. Note the .mjs extension for v4+
pdfjsLib.GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjsLib.version}/build/pdf.worker.min.mjs`;

const generateMetadata = (data: any[], fileName: string): string => {
    if (!data || data.length === 0) return `File ${fileName} is empty.`;
    
    // Get all unique headers from all rows (in case of inconsistent CSV)
    const headers = Array.from(new Set(data.flatMap(Object.keys)));
    
    // Limit to 1000 rows for performance/context safety
    const slicedData = data.slice(0, 1000);
    const rowCount = data.length;
    
    const csvRows = slicedData.map(row => {
        return headers.map(header => {
            const val = row[header];
            // Simple CSV escape: replace newlines and commas
            return val === null || val === undefined ? '' : String(val).replace(/[\n\r,]/g, ' ');
        }).join(',');
    });
    
    const csvString = [headers.join(','), ...csvRows].join('\n');
    
    return `\n\n**File Content (${fileName}):**\n\`\`\`csv\n${csvString}\n\`\`\`\n${rowCount > 1000 ? `\n(Showing first 1000 of ${rowCount} rows)` : ''}`;
};

export const parseCSV = (file: File): Promise<{ data: ActualDataItem[]; metadata: string }> => {
    return new Promise((resolve, reject) => {
        Papa.parse(file, {
            header: true,
            skipEmptyLines: true,
            complete: (results) => {
                // Generate metadata from the raw parsed data (all columns)
                const metadata = generateMetadata(results.data as any[], file.name);

                // Extract specific production data if available
                const parsedData: ActualDataItem[] = results.data.map((row: any) => ({
                    date: row.Date || row.date || '',
                    name: row.Name || row.name || '',
                    actual: parseFloat(row.Actual || row.actual || '0')
                })).filter(item => item.date && item.name);

                resolve({ data: parsedData, metadata });
            },
            error: (error) => reject(error)
        });
    });
};

export const parseExcel = async (file: File): Promise<{ data: ActualDataItem[]; metadata: string }> => {
    const reader = new FileReader();
    return new Promise((resolve, reject) => {
        reader.onload = async (e) => {
            try {
                const buffer = e.target?.result as ArrayBuffer;
                const workbook = new ExcelJS.Workbook();
                await workbook.xlsx.load(buffer);
                const worksheet = workbook.getWorksheet(1);
                if (!worksheet) {
                    resolve({ data: [], metadata: "" });
                    return;
                }

                const jsonData: any[] = [];
                const headers: string[] = [];
                worksheet.getRow(1).eachCell((cell, colNumber) => {
                    headers[colNumber] = cell.value?.toString() || '';
                });

                worksheet.eachRow((row, rowNumber) => {
                    if (rowNumber === 1) return;
                    const rowData: any = {};
                    row.eachCell((cell, colNumber) => {
                        const val = cell.value;
                        // Handle ExcelJS cell types (e.g. formulas, rich text) if necessary, 
                        // but usually cell.value is sufficient for basic data.
                        // For dates, ExcelJS returns Date objects.
                        rowData[headers[colNumber]] = val && typeof val === 'object' && 'result' in val ? val.result : val;
                    });
                    jsonData.push(rowData);
                });

                const metadata = generateMetadata(jsonData, file.name);

                const parsedData: ActualDataItem[] = jsonData.map((row: any) => ({
                    date: row.Date || row.date || '',
                    name: row.Name || row.name || '',
                    actual: parseFloat(row.Actual || row.actual || '0')
                })).filter(item => item.date && item.name);

                resolve({ data: parsedData, metadata });
            } catch (err) {
                reject(err);
            }
        };
        reader.readAsArrayBuffer(file);
    });
};

export const parsePDF = async (file: File): Promise<string> => {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    let fullText = '';

    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        const pageText = textContent.items.map((item: any) => item.str).join(' ');
        fullText += `\n--- Page ${i} ---\n${pageText}`;
    }

    return `\n\n**File Content (${file.name}):**\n\`\`\`text\n${fullText}\n\`\`\``;
};

export const parseDOCX = async (file: File): Promise<string> => {
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.extractRawText({ arrayBuffer });
    return `\n\n**File Content (${file.name}):**\n\`\`\`text\n${result.value}\n\`\`\``;
};

export const parsePPTX = async (file: File): Promise<string> => {
    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);
    let fullText = '';
    
    // Iterate through slide XML files
    const slideFiles = Object.keys(zip.files).filter(fileName => fileName.match(/ppt\/slides\/slide\d+\.xml/));
    
    // Sort slides numerically
    slideFiles.sort((a, b) => {
        const numA = parseInt(a.match(/slide(\d+)\.xml/)![1]);
        const numB = parseInt(b.match(/slide(\d+)\.xml/)![1]);
        return numA - numB;
    });

    for (const fileName of slideFiles) {
        const xmlContent = await zip.file(fileName)!.async('string');
        // Simple regex to extract text from XML tags (e.g., <a:t>Text</a:t>)
        const textMatches = xmlContent.match(/<a:t>(.*?)<\/a:t>/g);
        if (textMatches) {
            const slideText = textMatches.map(t => t.replace(/<\/?a:t>/g, '')).join(' ');
            const slideNum = fileName.match(/slide(\d+)\.xml/)![1];
            fullText += `\n--- Slide ${slideNum} ---\n${slideText}`;
        }
    }

    return `\n\n**File Content (${file.name}):**\n\`\`\`text\n${fullText}\n\`\`\``;
};

export const parseText = async (file: File): Promise<string> => {
    const text = await file.text();
    return `\n\n**File Content (${file.name}):**\n\`\`\`text\n${text}\n\`\`\``;
};

export const processImage = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target?.result as string);
        reader.onerror = (e) => reject(e);
        reader.readAsDataURL(file);
    });
};

export const handleFileProcessing = async (file: File): Promise<Partial<FileAttachment>> => {
    const fileType = file.type;
    const fileName = file.name.toLowerCase();
    
    const isCSV = fileType === 'text/csv' || fileName.endsWith('.csv');
    const isExcel = fileType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || fileName.endsWith('.xlsx') || fileName.endsWith('.xls');
    const isPDF = fileType === 'application/pdf' || fileName.endsWith('.pdf');
    const isDOCX = fileType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || fileName.endsWith('.docx');
    const isPPTX = fileType === 'application/vnd.openxmlformats-officedocument.presentationml.presentation' || fileName.endsWith('.pptx');
    const isText = fileType.startsWith('text/') || fileName.endsWith('.txt') || fileName.endsWith('.md') || fileName.endsWith('.json');
    const isImage = fileType.startsWith('image/');

    try {
        // Always read file as Data URL to allow download/preview
        const base64Data = await processImage(file);

        if (isCSV) {
            const { data: parsedData, metadata } = await parseCSV(file);
            return { name: file.name, type: fileType, data: base64Data, file, metadata, parsedData } as any;
        } else if (isExcel) {
            const { data: parsedData, metadata } = await parseExcel(file);
            return { name: file.name, type: fileType, data: base64Data, file, metadata, parsedData } as any;
        } else if (isPDF) {
            const metadata = await parsePDF(file);
            return { name: file.name, type: fileType, data: base64Data, file, metadata } as any;
        } else if (isDOCX) {
            const metadata = await parseDOCX(file);
            return { name: file.name, type: fileType, data: base64Data, file, metadata } as any;
        } else if (isPPTX) {
            const metadata = await parsePPTX(file);
            return { name: file.name, type: fileType, data: base64Data, file, metadata } as any;
        } else if (isText) {
            const metadata = await parseText(file);
            return { name: file.name, type: fileType, data: base64Data, file, metadata } as any;
        } else if (isImage) {
            // base64Data is already computed
            return { name: file.name, type: fileType, data: base64Data, file };
        } else {
            throw new Error(`Unsupported file type: ${fileType || fileName}`);
        }
    } catch (error) {
        console.error("Error processing file:", error);
        throw new Error(`Failed to process file ${file.name}: ${error instanceof Error ? error.message : String(error)}`);
    }
};
