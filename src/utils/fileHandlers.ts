import Papa from 'papaparse';
import ExcelJS from 'exceljs';
import { ActualDataItem, FileAttachment } from '../types/production';

export const parseCSV = (file: File): Promise<{ data: ActualDataItem[]; metadata: string }> => {
    return new Promise((resolve, reject) => {
        Papa.parse(file, {
            header: true,
            skipEmptyLines: true,
            complete: (results) => {
                const parsedData: ActualDataItem[] = results.data.map((row: any) => ({
                    date: row.Date || row.date || '',
                    name: row.Name || row.name || '',
                    actual: parseFloat(row.Actual || row.actual || '0')
                })).filter(item => item.date && item.name);

                const snippet = parsedData.slice(0, 5).map(row =>
                    `| ${row.date} | ${row.name} | ${row.actual} |`
                ).join('\n');

                const metadata = `\n\n**Data Preview (from ${file.name}):**\n| Date | Name | Actual |\n|---|---|---|\n${snippet}\n\nTotal rows: ${parsedData.length}`;

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
                        rowData[headers[colNumber]] = cell.value;
                    });
                    jsonData.push(rowData);
                });

                const parsedData: ActualDataItem[] = jsonData.map((row: any) => ({
                    date: row.Date || row.date || '',
                    name: row.Name || row.name || '',
                    actual: parseFloat(row.Actual || row.actual || '0')
                })).filter(item => item.date && item.name);

                const snippet = parsedData.slice(0, 5).map(row =>
                    `| ${row.date} | ${row.name} | ${row.actual} |`
                ).join('\n');

                const metadata = `\n\n**Data Preview (from ${file.name}):**\n| Date | Name | Actual |\n|---|---|---|\n${snippet}\n\nTotal rows: ${parsedData.length}`;

                resolve({ data: parsedData, metadata });
            } catch (err) {
                reject(err);
            }
        };
        reader.readAsArrayBuffer(file);
    });
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
    const isCSV = fileType === 'text/csv' || file.name.endsWith('.csv');
    const isExcel = fileType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || file.name.endsWith('.xlsx') || file.name.endsWith('.xls');
    const isImage = fileType.startsWith('image/');

    if (isCSV) {
        const { data: parsedData, metadata } = await parseCSV(file);
        return { name: file.name, type: fileType, data: '', file, metadata, parsedData } as any;
    } else if (isExcel) {
        const { data: parsedData, metadata } = await parseExcel(file);
        return { name: file.name, type: fileType, data: '', file, metadata, parsedData } as any;
    } else if (isImage) {
        const base64Data = await processImage(file);
        return { name: file.name, type: fileType, data: base64Data, file };
    } else {
        throw new Error("Unsupported file type.");
    }
};
