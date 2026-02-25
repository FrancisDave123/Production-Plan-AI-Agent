import ExcelJS from 'exceljs';
import { eachDayOfInterval, isValid, isSameDay } from 'date-fns';
import { ProjectData, ActualDataItem } from '../types/production';

export const getColumnLetter = (colIndex: number): string => {
    let letter = '';
    while (colIndex > 0) {
        let temp = (colIndex - 1) % 26;
        letter = String.fromCharCode(65 + temp) + letter;
        colIndex = Math.floor((colIndex - temp) / 26);
    }
    return letter;
};

export const sanitizeSheetName = (name: string) => name.replace(/[\[\]\:\*\?\/\\]/g, '').substring(0, 31);

export const generateExcelFile = async (projectData: ProjectData): Promise<ExcelJS.Buffer> => {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Production Plan Agent';
    workbook.created = new Date();

    const start = new Date(projectData.startDate);
    const end = new Date(projectData.endDate);

    if (!isValid(start) || !isValid(end)) {
        throw new Error("Invalid dates provided");
    }

    const days = eachDayOfInterval({ start, end });
    const scheduleItems: any[] = [];

    days.forEach(day => {
        projectData.resources.forEach(resource => {
            // Find actual data if it exists
            let actualMatch: ActualDataItem | null = null;
            if (projectData.actualData) {
                actualMatch = projectData.actualData.find(item => {
                    const itemDate = new Date(item.date);
                    return isValid(itemDate) && isSameDay(itemDate, day) && item.name.toLowerCase() === resource.toLowerCase();
                }) || null;
            }

            const item: any = {
                date: day,
                name: resource,
                actual: actualMatch ? actualMatch.actual : null
            };

            // Add extra daily data if defined
            if (projectData.dailyColumns && actualMatch) {
                projectData.dailyColumns.forEach(col => {
                    item[col.key] = actualMatch![col.key] || null;
                });
            }

            scheduleItems.push(item);
        });
    });

    const totalItems = scheduleItems.length;

    // --- LPB Target Distribution Logic ---
    const weights = scheduleItems.map((_, index) => {
        const t = index / (totalItems - 1 || 1);
        if (t < 0.25) {
            return 0.3 + (0.6 - 0.3) * (t / 0.25);
        } else if (t < 0.75) {
            return 0.6 + (1.0 - 0.6) * ((t - 0.25) / 0.5);
        } else {
            return 1.0;
        }
    });

    const totalWeight = weights.reduce((sum, w) => sum + w, 0);
    const itemsWithTargets = scheduleItems.map((item, index) => ({
        ...item,
        target: (weights[index] / totalWeight) * projectData.goal
    }));

    // --- Sheet 1: Daily_Production_Key ---
    const sheetKey = workbook.addWorksheet(sanitizeSheetName('Daily_Production_Key'));
    const baseKeyCols = [
        { header: 'Date', key: 'date', width: 15 },
        { header: 'Day', key: 'day', width: 15 },
        { header: 'Week', key: 'week', width: 10 },
        { header: 'Month', key: 'month', width: 15 },
        { header: 'Name', key: 'name', width: 20 },
    ];

    const dynamicKeyCols = projectData.dailyColumns.map(col => ({
        header: col.header,
        key: col.key,
        width: 15,
        formula: col.formula
    }));

    sheetKey.columns = [...baseKeyCols, ...dynamicKeyCols];

    const tableCols = [
        { name: 'Date', filterButton: true },
        { name: 'Day', filterButton: true },
        { name: 'Week', filterButton: true },
        { name: 'Month', filterButton: true },
        { name: 'Name', filterButton: true },
        ...dynamicKeyCols.map(col => ({
            name: col.header,
            filterButton: true,
            totalsRowFunction: col.header.toLowerCase().includes('rate') ? undefined : 'sum'
        }))
    ];

    sheetKey.addTable({
        name: 'DailyProductionTable',
        ref: 'A1',
        headerRow: true,
        totalsRow: true,
        style: { theme: 'TableStyleMedium2', showRowStripes: true },
        columns: tableCols as any,
        rows: itemsWithTargets.map((item, index) => {
            const rowIndex = index + 2;
            const row: any[] = [
                item.date,
                { formula: `TEXT(A${rowIndex}, "dddd")` },
                { formula: `WEEKNUM(A${rowIndex})` },
                { formula: `TEXT(A${rowIndex}, "mmmm")` },
                item.name,
            ];

            dynamicKeyCols.forEach(col => {
                if (col.formula) {
                    row.push({ formula: col.formula.replace(/{rowIndex}/g, rowIndex.toString()) });
                } else {
                    row.push(item[col.key]);
                }
            });
            return row;
        }),
    });

    // Numbers formatting
    const keyRows = scheduleItems.length + 1;
    const lastKeyColLetter = getColumnLetter(sheetKey.columns.length);
    sheetKey.addConditionalFormatting({
        ref: `F2:${lastKeyColLetter}${keyRows}`,
        rules: [
            { priority: 1, type: 'expression', formulae: ['MOD(F2,1)=0'], style: { numFmt: '#,##0' } },
            { priority: 2, type: 'expression', formulae: ['MOD(F2,1)<>0'], style: { numFmt: '#,##0.00' } },
        ],
    });

    // --- Sheet 2: Production Plan ---
    const sheetPlan = workbook.addWorksheet(sanitizeSheetName(`${projectData.name} Plan`));
    const unitLabel = (projectData.unit || 'Units').charAt(0).toUpperCase() + (projectData.unit || 'Units').slice(1);

    const dynamicColumns = [
        { header: 'Date', key: 'date', width: 15, section: 'Target' as const },
        { header: 'Month', key: 'month', width: 15, section: 'Target' as const },
        ...projectData.columns
    ];

    sheetPlan.columns = dynamicColumns.map(col => ({ key: col.key, width: col.width || 18 }));

    // Header Row 1
    const lastColLetter = getColumnLetter(dynamicColumns.length);
    sheetPlan.mergeCells(`A1:${lastColLetter}1`);
    const titleCell = sheetPlan.getCell('A1');
    titleCell.value = `${projectData.name}: Production Plan & Daily Output Tracking`;
    titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF006633' } };
    titleCell.font = { color: { argb: 'FFFFFFFF' }, bold: true, size: 12 };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };

    // Section Headers Rows 2 & 3
    const sections = ['Target', 'Actual', 'Accumulative'] as const;
    let currentColIndex = 1;
    sections.forEach(section => {
        const sectionCols = dynamicColumns.filter(c => c.section === section);
        if (sectionCols.length > 0) {
            const startCol = getColumnLetter(currentColIndex);
            const endCol = getColumnLetter(currentColIndex + sectionCols.length - 1);

            if (section === 'Target') {
                sheetPlan.mergeCells(`${startCol}2:${endCol}3`);
                const cell = sheetPlan.getCell(`${startCol}2`);
                cell.value = `Target ${unitLabel} Output`;
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
            } else {
                sheetPlan.mergeCells(`${startCol}2:${endCol}2`);
                const cell2 = sheetPlan.getCell(`${startCol}2`);
                cell2.value = section === 'Actual' ? `${unitLabel} Output Tracking` : section;
                cell2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF70AD47' } };
                cell2.font = { color: { argb: 'FFFFFFFF' }, bold: true };
                cell2.alignment = { horizontal: 'center', vertical: 'middle' };

                sheetPlan.mergeCells(`${startCol}3:${endCol}3`);
                const cell3 = sheetPlan.getCell(`${startCol}3`);
                cell3.value = section;
                cell3.fill = { type: 'pattern', pattern: 'solid', fgColor: section === 'Actual' ? { argb: 'FFC6E0B4' } : { argb: 'FFD9D9D9' } };
                cell3.font = { bold: true };
                cell3.alignment = { horizontal: 'center', vertical: 'middle' };
            }
            currentColIndex += sectionCols.length;
        }
    });

    // Header Row 4
    const headerRow4 = sheetPlan.getRow(4);
    dynamicColumns.forEach((col, i) => {
        const cell = headerRow4.getCell(i + 1);
        cell.value = col.header;
        cell.font = { bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        if (col.section === 'Target') cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
        else if (col.section === 'Actual') cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6E0B4' } };
        else cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
    });
    headerRow4.height = 40;

    // Data Rows
    const uniqueDates = Array.from(new Set(scheduleItems.map(s => s.date.toISOString()))).sort();
    uniqueDates.forEach((dateIso, index) => {
        const rowIndex = index + 5;
        const dateObj = new Date(dateIso);
        const row = sheetPlan.getRow(rowIndex);
        dynamicColumns.forEach((col, colIdx) => {
            const cell = row.getCell(colIdx + 1);
            if (col.key === 'date') cell.value = dateObj;
            else if (col.key === 'month') cell.value = { formula: `TEXT(A${rowIndex}, "mmmm")` };
            else if (col.formula) cell.value = { formula: col.formula.replace(/{rowIndex}/g, rowIndex.toString()) };
            cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
            if (col.header.toLowerCase().includes('rate') || col.header.toLowerCase().includes('%')) cell.numFmt = '0.00%';
        });
    });

    // Grand Total
    const totalRowIndex = uniqueDates.length + 5;
    const totalRow = sheetPlan.getRow(totalRowIndex);
    totalRow.getCell(1).value = 'Grand Total';
    totalRow.font = { bold: true };
    dynamicColumns.forEach((col, colIdx) => {
        if (colIdx === 0) return;
        const cell = totalRow.getCell(colIdx + 1);
        const colLetter = getColumnLetter(colIdx + 1);
        const isRate = col.header.toLowerCase().includes('rate') || col.header.toLowerCase().includes('%');
        if (col.key !== 'month' && !isRate) {
            cell.value = { formula: `SUM(${colLetter}5:${colLetter}${totalRowIndex - 1})` };
        }
        cell.border = { top: { style: 'double' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
    });

    // Plan Conditional Formatting
    sheetPlan.addConditionalFormatting({
        ref: `C5:${lastColLetter}${totalRowIndex}`,
        rules: [
            { priority: 1, type: 'expression', formulae: ['MOD(C5,1)=0'], style: { numFmt: '#,##0' } },
            { priority: 2, type: 'expression', formulae: ['MOD(C5,1)<>0'], style: { numFmt: '#,##0.00' } },
        ],
    });

    // --- Sheet 3: Pivot Summary ---
    const sheetPivot = workbook.addWorksheet(sanitizeSheetName('Production_Pivot'));
    const pivotColumns = projectData.pivotColumns || [
        { header: 'Total Target', formula: `SUMIFS(DailyProductionTable[Target], DailyProductionTable[Week], A{rowIndex})` },
        { header: 'Total Actual', formula: `SUMIFS(DailyProductionTable[Actual], DailyProductionTable[Week], A{rowIndex})` },
        { header: 'Total Variance', formula: `SUMIFS(DailyProductionTable[Variance], DailyProductionTable[Week], A{rowIndex})` },
        { header: 'Cumulative Actual', formula: `SUM($D$2:D{rowIndex})` },
    ];

    const pivotCols = [
        { header: 'Week', key: 'week', width: 10 },
        { header: 'Month', key: 'month', width: 15 },
        ...pivotColumns.map(col => ({ header: col.header, width: 20, formula: col.formula }))
    ];
    sheetPivot.columns = pivotCols.map(c => ({ key: c.header, width: c.width }));

    const uniqueWeeks = Array.from(new Set(scheduleItems.map(item => {
        // Note: This is an approximation for generating the weekly list in JS
        // In Excel, the table helps. We'll just generate enough rows to cover the period.
        const weekNum = Math.ceil((item.date.getTime() - start.getTime()) / (7 * 24 * 60 * 60 * 1000)) + 1;
        return weekNum;
    }))).sort((a, b) => a - b);

    uniqueWeeks.forEach((week, idx) => {
        const rIdx = idx + 2;
        const row = sheetPivot.getRow(rIdx);
        row.getCell(1).value = week;
        row.getCell(2).value = { formula: `VLOOKUP(A${rIdx}, DailyProductionTable[[Week]:[Month]], 2, FALSE)` };
        pivotColumns.forEach((col, cIdx) => {
            row.getCell(cIdx + 3).value = { formula: col.formula.replace(/{rowIndex}/g, rIdx.toString()) };
        });
    });

    sheetPivot.addTable({
        name: 'PivotTable',
        ref: 'A1',
        headerRow: true,
        style: { theme: 'TableStyleMedium9', showRowStripes: true },
        columns: pivotCols.map(c => ({ name: c.header, filterButton: true })),
        rows: [] // rows already added
    });

    return await workbook.xlsx.writeBuffer();
};
