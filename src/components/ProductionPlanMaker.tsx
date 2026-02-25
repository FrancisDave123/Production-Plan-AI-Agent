import React, { useState, useEffect, useRef } from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { format, eachDayOfInterval, isValid, isSameDay } from 'date-fns';
import { Send, Bot, User as UserIcon, RefreshCw, FileSpreadsheet, Loader2, Download, Paperclip, X } from 'lucide-react';
import { GoogleGenAI, Type } from "@google/genai";
import Papa from 'papaparse';
import ReactMarkdown from 'react-markdown';

interface Message {
  id: string;
  role: 'agent' | 'user';
  content: string;
  type?: 'text' | 'file';
  fileData?: {
    name: string;
    buffer: ExcelJS.Buffer;
  };
}

interface ActualDataItem {
  date: string;
  name: string;
  actual: number;
  [key: string]: any;
}

interface ProjectColumn {
  header: string;
  key: string;
  section: 'Target' | 'Actual' | 'Accumulative';
  formula?: string;
  width?: number;
}

interface DailyColumn {
  header: string;
  key: string;
  formula?: string;
}

interface ProjectData {
  name: string;
  goal: number;
  unit: string;
  startDate: string;
  endDate: string;
  resources: string[];
  actualData?: ActualDataItem[];
  columns: ProjectColumn[];
  dailyColumns: DailyColumn[];
  pivotColumns?: { header: string; formula: string }[];
  dashboardMetrics?: { label: string; formula: string; format?: string }[];
}

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

const getColumnLetter = (colIndex: number): string => {
  let letter = '';
  while (colIndex > 0) {
    let temp = (colIndex - 1) % 26;
    letter = String.fromCharCode(65 + temp) + letter;
    colIndex = Math.floor((colIndex - temp) / 26);
  }
  return letter;
};

const sanitizeSheetName = (name: string) => name.replace(/[\[\]\:\*\?\/\\]/g, '').substring(0, 31);

export default function ProductionPlanMaker() {
  const [messages, setMessages] = useState<Message[]>([
    {
      id: '1',
      role: 'agent',
      content: "Hello! I'm your Production Plan Agent. I can help you create a detailed Excel production plan. \n\nTo get started, please tell me about your project: **What is the project name, your total goal, the start/end dates, and who is working on it?**"
    }
  ]);
  const [inputValue, setInputValue] = useState('');
  const [isTyping, setIsTyping] = useState(false);
  const [isStreaming, setIsStreaming] = useState(false);
  const [uploadedData, setUploadedData] = useState<ActualDataItem[] | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const chatRef = useRef<any>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const streamIntervalRef = useRef<ReturnType<typeof setInterval> | null>(null);

  useEffect(() => {
    if (!chatRef.current) {
      chatRef.current = ai.chats.create({
        model: "gemini-2.0-flash",
        config: {
          systemInstruction: `You are a professional Production Planning Assistant. 
          Your goal is to collect the following information from the user to generate an Excel production plan:
          1. Project Name
          2. Overall Goal (numeric value)
          3. Unit of measurement (e.g., units, hours, revenue)
          4. Start Date (YYYY-MM-DD)
          5. End Date (YYYY-MM-DD)
          6. List of Resources/Teams
          
          The user may also provide 'actual' production data points (Date, Name, Actual value) directly in the chat or via file upload.
          If they provide it in text, extract it into the 'actualData' parameter.
          
          DYNAMIC SCHEMA ANALYSIS:
          Based on the project type and unit, you must suggest a set of columns for the production plan. 
          
          1. DAILY KEY COLUMNS: Define raw data tracked per resource per day. 
             - ALWAYS include 'Target' and 'Actual'.
             - Add others like 'Duration', 'Ops', or 'Variance'.
             - You can provide formulas for calculated fields (e.g., Variance: G{rowIndex}-F{rowIndex}).
          2. PLAN COLUMNS: Define daily summaries in the main plan (e.g., 'Total Target', 'Actual Ops').
          3. PIVOT COLUMNS: Define weekly/monthly aggregations (e.g., 'Total Actual', 'Avg Variance').
          4. DASHBOARD METRICS: Define high-level KPIs (e.g., 'Overall Goal', '% Completion').
          
          TABLE & COLUMN NAMES:
          - The raw data table is named 'DailyProductionTable'.
          - Base columns in 'DailyProductionTable' are: [Date (A), Day (B), Week (C), Month (D), Name (E)].
          - Your 'dailyColumns' start at Column F (Index 6).
          - Use these names and letters EXACTLY in your formulas.
          
          For every PLAN, PIVOT, and DASHBOARD item, you MUST provide an Excel formula that references the 'DailyProductionTable'.
          Use {rowIndex} for relative row references in Plan/Pivot.
          
          You MUST suggest this full architecture to the user and confirm it before generation.
          
          IMPORTANT: You are a specialized Production Plan Agent. You must ONLY respond to queries related to production planning, project scheduling, and Excel generation for these plans. 
          If a user asks about unrelated topics (e.g., weather, general knowledge, jokes, other software), politely decline and redirect them back to production planning.
          
          Once you have the core project details (1-6) and have confirmed the full 4-sheet architecture, call the 'generate_production_plan' tool.
          Be conversational and helpful within your domain. If information is missing, ask for it.`,
          tools: [{
            functionDeclarations: [{
              name: "generate_production_plan",
              description: "Generates the production planning Excel file once core project details and column structure are confirmed.",
              parameters: {
                type: Type.OBJECT,
                properties: {
                  name: { type: Type.STRING, description: "The name of the project" },
                  goal: { type: Type.NUMBER, description: "The total numeric goal" },
                  unit: { type: Type.STRING, description: "The unit of measurement (e.g., 'units', 'hours')" },
                  startDate: { type: Type.STRING, description: "Start date in YYYY-MM-DD format" },
                  endDate: { type: Type.STRING, description: "End date in YYYY-MM-DD format" },
                  resources: {
                    type: Type.ARRAY,
                    items: { type: Type.STRING },
                    description: "List of names of teams or individuals"
                  },
                  columns: {
                    type: Type.ARRAY,
                    items: {
                      type: Type.OBJECT,
                      properties: {
                        header: { type: Type.STRING, description: "The display name of the column" },
                        key: { type: Type.STRING, description: "A unique key for the column" },
                        section: {
                          type: Type.STRING,
                          enum: ["Target", "Actual", "Accumulative"],
                          description: "Which section the column belongs to"
                        },
                        formula: { type: Type.STRING, description: "Excel formula referencing DailyProductionTable. Use {rowIndex} for the current row." }
                      },
                      required: ["header", "key", "section", "formula"]
                    },
                    description: "The dynamic list of columns for the summary plan sheet"
                  },
                  dailyColumns: {
                    type: Type.ARRAY,
                    items: {
                      type: Type.OBJECT,
                      properties: {
                        header: { type: Type.STRING, description: "Display name" },
                        key: { type: Type.STRING, description: "Key (must match keys in actualData)" },
                        formula: { type: Type.STRING, description: "Optional formula referencing other columns in the same row (e.g., G{rowIndex}-F{rowIndex})" }
                      },
                      required: ["header", "key"]
                    },
                    description: "Data columns for the Daily Key sheet. Usually includes 'Target', 'Actual', 'Variance', etc."
                  },
                  pivotColumns: {
                    type: Type.ARRAY,
                    items: {
                      type: Type.OBJECT,
                      properties: {
                        header: { type: Type.STRING, description: "Display name" },
                        formula: { type: Type.STRING, description: "Excel formula referencing DailyProductionTable. Use {rowIndex} for current week/month row." }
                      },
                      required: ["header", "formula"]
                    },
                    description: "Dynamic columns for the Pivot Summary sheet"
                  },
                  dashboardMetrics: {
                    type: Type.ARRAY,
                    items: {
                      type: Type.OBJECT,
                      properties: {
                        label: { type: Type.STRING, description: "KPI Label" },
                        formula: { type: Type.STRING, description: "Excel formula for the metric" },
                        format: { type: Type.STRING, description: "Optional number format (e.g., '0.00%')" }
                      },
                      required: ["label", "formula"]
                    },
                    description: "Dynamic metrics for the Summary Dashboard"
                  },
                  actualData: {
                    type: Type.ARRAY,
                    items: {
                      type: Type.OBJECT,
                      properties: {
                        date: { type: Type.STRING, description: "Date in YYYY-MM-DD format" },
                        name: { type: Type.STRING, description: "Resource name" },
                        actual: { type: Type.NUMBER, description: "Main actual production value" }
                      },
                      required: ["date", "name", "actual"]
                    },
                    description: "Optional list of actual production data points. Include any extra fields mentioned by the user (e.g., duration, ops) as additional properties."
                  }
                },
                required: ["name", "goal", "unit", "startDate", "endDate", "resources", "columns"]
              }
            }]
          }]
        }
      });
    }
  }, []);

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  };

  useEffect(() => {
    scrollToBottom();
  }, [messages, isTyping]);

  useEffect(() => {
    return () => {
      if (streamIntervalRef.current) clearInterval(streamIntervalRef.current);
    };
  }, []);

  const typewriterEffect = (fullText: string, msgId: string) => {
    let i = 0;
    setIsStreaming(true);
    if (streamIntervalRef.current) clearInterval(streamIntervalRef.current);
    streamIntervalRef.current = setInterval(() => {
      i++;
      setMessages(prev =>
        prev.map(m => m.id === msgId ? { ...m, content: fullText.slice(0, i) } : m)
      );
      if (i >= fullText.length) {
        clearInterval(streamIntervalRef.current!);
        streamIntervalRef.current = null;
        setIsStreaming(false);
      }
    }, 15);
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setFileName(file.name);
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: (results) => {
        const parsedData: ActualDataItem[] = results.data.map((row: any) => ({
          date: row.Date || row.date || '',
          name: row.Name || row.name || '',
          actual: parseFloat(row.Actual || row.actual || '0')
        })).filter(item => item.date && item.name);
        setUploadedData(parsedData);
        setMessages(prev => [...prev, {
          id: Date.now().toString(),
          role: 'user',
          content: `Uploaded actual data: ${file.name} (${parsedData.length} rows detected)`
        }]);
      },
      error: (error) => {
        console.error("CSV Parse Error:", error);
        alert("Error parsing CSV file. Please ensure it has Date, Name, and Actual columns.");
      }
    });
  };

  const handleSendMessage = async () => {
    if (!inputValue.trim() || isTyping || isStreaming) return;
    const userMsg: Message = { id: Date.now().toString(), role: 'user', content: inputValue };
    setMessages(prev => [...prev, userMsg]);
    setInputValue('');
    setIsTyping(true);

    try {
      const result = await chatRef.current.sendMessage({ message: userMsg.content });
      const response = result;

      if (response.functionCalls) {
        for (const call of response.functionCalls) {
          if (call.name === 'generate_production_plan') {
            const projectData = call.args as ProjectData;
            const combinedActualData = [...(projectData.actualData || [])];
            if (uploadedData) {
              uploadedData.forEach(upItem => {
                const exists = combinedActualData.some(combItem =>
                  combItem.date === upItem.date && combItem.name === upItem.name
                );
                if (!exists) combinedActualData.push(upItem);
              });
            }
            projectData.actualData = combinedActualData.length > 0 ? combinedActualData : undefined;
            await generateExcelFile(projectData);
          }
        }
      } else {
        const fullText = response.text || "I'm sorry, I didn't quite get that. Could you please provide the project details?";
        const msgId = Date.now().toString();
        setMessages(prev => [...prev, { id: msgId, role: 'agent', content: '' }]);
        typewriterEffect(fullText, msgId);
      }
    } catch (error) {
      console.error("Gemini Error:", error);
      const msgId = Date.now().toString();
      setMessages(prev => [...prev, { id: msgId, role: 'agent', content: '' }]);
      typewriterEffect("I'm having a bit of trouble connecting to my brain. Could you try again?", msgId);
    } finally {
      setIsTyping(false);
    }
  };

  const generateExcelFile = async (projectData: ProjectData) => {
    try {
      const workbook = new ExcelJS.Workbook();
      workbook.creator = 'Production Plan Agent';
      workbook.created = new Date();

      const start = new Date(projectData.startDate);
      const end = new Date(projectData.endDate);
      if (!isValid(start) || !isValid(end)) throw new Error("Invalid dates provided");

      const days = eachDayOfInterval({ start, end });
      const scheduleItems: any[] = [];

      days.forEach(day => {
        projectData.resources.forEach(resource => {
          let actualMatch: ActualDataItem | null = null;
          if (projectData.actualData) {
            actualMatch = projectData.actualData.find(item => {
              const itemDate = new Date(item.date);
              return isValid(itemDate) && isSameDay(itemDate, day) && item.name.toLowerCase() === resource.toLowerCase();
            }) || null;
          }
          const item: any = { date: day, name: resource, actual: actualMatch ? actualMatch.actual : null };
          if (projectData.dailyColumns && actualMatch) {
            projectData.dailyColumns.forEach(col => { item[col.key] = actualMatch![col.key] || null; });
          }
          scheduleItems.push(item);
        });
      });

      const totalItems = scheduleItems.length;
      const weights = scheduleItems.map((_, index) => {
        const t = index / (totalItems - 1 || 1);
        if (t < 0.25) return 0.3 + (0.6 - 0.3) * (t / 0.25);
        else if (t < 0.75) return 0.6 + (1.0 - 0.6) * ((t - 0.25) / 0.5);
        else return 1.0;
      });
      const totalWeight = weights.reduce((sum, w) => sum + w, 0);
      const itemsWithTargets = scheduleItems.map((item, index) => ({
        ...item,
        target: (weights[index] / totalWeight) * projectData.goal
      }));

      // Sheet 1
      const sheetKey = workbook.addWorksheet(sanitizeSheetName('Daily_Production_Key'));
      const baseKeyCols = [
        { header: 'Date', key: 'date', width: 15 },
        { header: 'Day', key: 'day', width: 15 },
        { header: 'Week', key: 'week', width: 10 },
        { header: 'Month', key: 'month', width: 15 },
        { header: 'Name', key: 'name', width: 20 },
      ];
      const dynamicKeyCols = projectData.dailyColumns.map(col => ({ header: col.header, key: col.key, width: 15, formula: col.formula }));
      sheetKey.columns = [...baseKeyCols, ...dynamicKeyCols];
      const tableCols = [
        { name: 'Date', filterButton: true }, { name: 'Day', filterButton: true },
        { name: 'Week', filterButton: true }, { name: 'Month', filterButton: true },
        { name: 'Name', filterButton: true },
        ...dynamicKeyCols.map(col => ({ name: col.header, filterButton: true, totalsRowFunction: col.header.toLowerCase().includes('rate') ? undefined : 'sum' }))
      ];
      sheetKey.addTable({
        name: 'DailyProductionTable', ref: 'A1', headerRow: true, totalsRow: true,
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
            if (col.formula) row.push({ formula: col.formula.replace(/{rowIndex}/g, rowIndex.toString()) });
            else row.push(item[col.key]);
          });
          return row;
        }),
      });
      const keyRows = scheduleItems.length + 1;
      const lastKeyColLetter = getColumnLetter(sheetKey.columns.length);
      sheetKey.addConditionalFormatting({
        ref: `F2:${lastKeyColLetter}${keyRows}`,
        rules: [
          { priority: 1, type: 'expression', formulae: ['MOD(F2,1)=0'], style: { numFmt: '#,##0' } },
          { priority: 2, type: 'expression', formulae: ['MOD(F2,1)<>0'], style: { numFmt: '#,##0.00' } },
        ],
      });

      // Sheet 2
      const sheetPlan = workbook.addWorksheet(sanitizeSheetName(`${projectData.name} Plan`));
      const unit = projectData.unit || 'Units';
      const unitLabel = unit.charAt(0).toUpperCase() + unit.slice(1);
      const dynamicColumns: ProjectColumn[] = [
        { header: 'Date', key: 'date', width: 15, section: 'Target' as const },
        { header: 'Month', key: 'month', width: 15, section: 'Target' as const },
        ...projectData.columns
      ];
      sheetPlan.columns = dynamicColumns.map(col => ({ key: col.key, width: col.width || 18 }));

      const totalCols = dynamicColumns.length;
      const lastColLetter = getColumnLetter(totalCols);
      sheetPlan.mergeCells(`A1:${lastColLetter}1`);
      const titleCell = sheetPlan.getCell('A1');
      titleCell.value = `${projectData.name}: Production Plan & Daily Output Tracking`;
      titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF133020' } };
      titleCell.font = { color: { argb: 'FFFFFFFF' }, bold: true, size: 12 };
      titleCell.alignment = { horizontal: 'center', vertical: 'middle' };

      const sections = ['Target', 'Actual', 'Accumulative'] as const;
      let currentColIndex = 1;
      sections.forEach(section => {
        const sectionCols = dynamicColumns.filter(c => c.section === section);
        if (sectionCols.length > 0) {
          const startCol = getColumnLetter(currentColIndex);
          const endCol = getColumnLetter(currentColIndex + sectionCols.length - 1);
          const ref2 = `${startCol}2:${endCol}2`;
          sheetPlan.mergeCells(ref2);
          const cell2 = sheetPlan.getCell(`${startCol}2`);
          cell2.value = section === 'Target' ? `Target ${unitLabel} Output` : (section === 'Actual' ? `${unitLabel} Output Tracking` : section);
          if (section === 'Target') {
            cell2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
          } else {
            cell2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF046241' } };
            cell2.font = { color: { argb: 'FFFFFFFF' }, bold: true };
          }
          cell2.alignment = { horizontal: 'center', vertical: 'middle' };
          if (section === 'Actual' || section === 'Accumulative') {
            const ref3 = `${startCol}3:${endCol}3`;
            sheetPlan.mergeCells(ref3);
            const cell3 = sheetPlan.getCell(`${startCol}3`);
            cell3.value = section;
            cell3.fill = { type: 'pattern', pattern: 'solid', fgColor: section === 'Actual' ? { argb: 'FFFFC370' } : { argb: 'FFD9D9D9' } };
            cell3.font = { bold: true };
            cell3.alignment = { horizontal: 'center', vertical: 'middle' };
          } else {
            sheetPlan.unMergeCells(ref2);
            sheetPlan.mergeCells(`${startCol}2:${endCol}3`);
          }
          currentColIndex += sectionCols.length;
        }
      });

      const headerRow4 = sheetPlan.getRow(4);
      dynamicColumns.forEach((col, i) => {
        const cell = headerRow4.getCell(i + 1);
        cell.value = col.header;
        cell.font = { bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        if (col.section === 'Target') cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
        else if (col.section === 'Actual') cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC370' } };
        else cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
      });
      headerRow4.height = 40;

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

      const totalRowIndex = uniqueDates.length + 5;
      const totalRow = sheetPlan.getRow(totalRowIndex);
      totalRow.getCell(1).value = 'Grand Total';
      totalRow.font = { bold: true };
      dynamicColumns.forEach((col, colIdx) => {
        if (colIdx === 0) return;
        const cell = totalRow.getCell(colIdx + 1);
        const colLetter = getColumnLetter(colIdx + 1);
        const isRate = col.header.toLowerCase().includes('rate') || col.header.toLowerCase().includes('%');
        const isMonth = col.key === 'month';
        if (!isMonth && !isRate) cell.value = { formula: `SUM(${colLetter}5:${colLetter}${totalRowIndex - 1})` };
        cell.border = { top: { style: 'double' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF9F7F7' } };
      });

      const planRows = uniqueDates.length + 5;
      sheetPlan.addConditionalFormatting({
        ref: `C5:${lastColLetter}${planRows}`,
        rules: [
          { priority: 1, type: 'expression', formulae: ['MOD(C5,1)=0'], style: { numFmt: '#,##0' } },
          { priority: 2, type: 'expression', formulae: ['MOD(C5,1)<>0'], style: { numFmt: '#,##0.00' } },
        ],
      });

      // Sheet 3
      const sheetPivot = workbook.addWorksheet(sanitizeSheetName('Production_Pivot'));
      const basePivotCols = [{ header: 'Week', key: 'week', width: 10 }, { header: 'Month', key: 'month', width: 15 }];
      const dynamicPivotCols = (projectData.pivotColumns || [
        { header: 'Total Target', formula: `SUMIFS(DailyProductionTable[Target], DailyProductionTable[Week], A{rowIndex})` },
        { header: 'Total Actual', formula: `SUMIFS(DailyProductionTable[Actual], DailyProductionTable[Week], A{rowIndex})` },
        { header: 'Total Variance', formula: `SUMIFS(DailyProductionTable[Variance], DailyProductionTable[Week], A{rowIndex})` },
        { header: 'Cumulative Actual', formula: `SUM($D$2:D{rowIndex})` },
      ]).filter(col => !basePivotCols.some(bp => bp.header === col.header)).map(col => ({ header: col.header, formula: col.formula, width: 18 }));

      sheetPivot.columns = [...basePivotCols, ...dynamicPivotCols.map(c => ({ header: c.header, key: c.header.replace(/\s+/g, ''), width: c.width }))];
      const weeks = Array.from(new Set(scheduleItems.map(s => format(s.date, 'w')))).sort((a, b) => Number(a) - Number(b));
      sheetPivot.addTable({
        name: 'PivotSummaryTable', ref: 'A1', headerRow: true, totalsRow: true,
        style: { theme: 'TableStyleLight1', showRowStripes: true },
        columns: [
          { name: 'Week', filterButton: true }, { name: 'Month', filterButton: true },
          ...dynamicPivotCols.map(col => ({
            name: col.header, filterButton: true,
            totalsRowFunction: (col.header.toLowerCase().includes('rate') || col.header.toLowerCase().includes('%') ? 'none' : 'sum') as any
          }))
        ],
        rows: weeks.map((week, index) => {
          const rowIndex = index + 2;
          const row: any[] = [Number(week), { formula: `INDEX(DailyProductionTable[Month], MATCH(${week}, DailyProductionTable[Week], 0))` }];
          dynamicPivotCols.forEach(col => row.push({ formula: col.formula.replace(/{rowIndex}/g, rowIndex.toString()) }));
          return row;
        })
      });

      const pivotRows = weeks.length + 1;
      const lastPivotColLetter = getColumnLetter(sheetPivot.columns.length);
      sheetPivot.addConditionalFormatting({
        ref: `C2:${lastPivotColLetter}${pivotRows}`,
        rules: [
          { priority: 1, type: 'expression', formulae: ['MOD(C2,1)=0'], style: { numFmt: '#,##0' } },
          { priority: 2, type: 'expression', formulae: ['MOD(C2,1)<>0'], style: { numFmt: '#,##0.00' } },
        ],
      });

      // Sheet 4
      const sheetDash = workbook.addWorksheet(sanitizeSheetName('Summary_Dashboard'));
      sheetDash.mergeCells('A1:B1');
      sheetDash.getCell('A1').value = 'Project Summary Dashboard';
      sheetDash.getCell('A1').font = { bold: true, size: 16, color: { argb: 'FF133020' } };
      sheetDash.getCell('A1').alignment = { horizontal: 'center' };

      const dashData = projectData.dashboardMetrics || [
        { label: 'Overall Goal', formula: `${projectData.goal}` },
        { label: 'Total Actual', formula: `SUM(DailyProductionTable[Actual])` },
        { label: 'Total Remaining', formula: `B3-B4` },
        { label: '% Completion', formula: `B4/B3`, format: '0.00%' },
        { label: 'Avg Daily Production', formula: `AVERAGE(DailyProductionTable[Actual])` },
        { label: 'Required Daily Production', formula: `IF(OR(B5<=0, COUNTBLANK(DailyProductionTable[Actual])=0), 0, B5 / COUNTBLANK(DailyProductionTable[Actual]))` },
        { label: 'Status', formula: `IF(B5<=0, "Completed", IF(B7>=AVERAGE(DailyProductionTable[Target]), "On Track", "Behind"))` }
      ];

      dashData.forEach((item, index) => {
        const r = index + 3;
        sheetDash.getCell(`A${r}`).value = item.label;
        sheetDash.getCell(`A${r}`).font = { bold: true, color: { argb: 'FF133020' } };
        const cell = sheetDash.getCell(`B${r}`);
        cell.value = { formula: item.formula };
        if (item.format) {
          cell.numFmt = item.format;
        } else {
          sheetDash.addConditionalFormatting({
            ref: `B${r}`,
            rules: [
              { priority: 1, type: 'expression', formulae: [`MOD(B${r},1)=0`], style: { numFmt: '#,##0' } },
              { priority: 2, type: 'expression', formulae: [`MOD(B${r},1)<>0`], style: { numFmt: '#,##0.00' } },
            ],
          });
        }
      });
      sheetDash.getColumn(1).width = 30;
      sheetDash.getColumn(2).width = 25;

      const buffer = await workbook.xlsx.writeBuffer();
      const msgId = Date.now().toString();
      const successText = `I've generated the production plan for **${projectData.name}**. You can download it below.`;
      setMessages(prev => [...prev, {
        id: msgId, role: 'agent', content: '', type: 'file',
        fileData: {
          name: `${projectData.name.replace(/\s+/g, '_')}_Production_Planning.xlsx`,
          buffer: buffer as ExcelJS.Buffer
        }
      }]);
      typewriterEffect(successText, msgId);

    } catch (error) {
      console.error(error);
      const msgId = Date.now().toString();
      setMessages(prev => [...prev, { id: msgId, role: 'agent', content: '' }]);
      typewriterEffect("I encountered an error generating the Excel file. Please check the details and try again.", msgId);
    }
  };

  const handleDownload = (fileName: string, buffer: ExcelJS.Buffer) => {
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, fileName);
  };

  const resetChat = () => {
    if (streamIntervalRef.current) clearInterval(streamIntervalRef.current);
    setIsStreaming(false);
    setMessages([{
      id: '1',
      role: 'agent',
      content: "Hello! I'm your Production Plan Agent. I can help you create a detailed Excel production plan. \n\nTo get started, please tell me about your project: **What is the project name, your total goal, the start/end dates, and who is working on it?**"
    }]);
    setUploadedData(null);
    setFileName(null);
    chatRef.current = null;
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSendMessage();
    }
  };

  const textareaRef = useRef<HTMLTextAreaElement>(null);
  useEffect(() => {
    if (textareaRef.current) {
      textareaRef.current.style.height = 'auto';
      textareaRef.current.style.height = `${Math.min(textareaRef.current.scrollHeight, 200)}px`;
    }
  }, [inputValue]);

  return (
    <div className="max-w-2xl mx-auto h-[700px] flex flex-col rounded-2xl shadow-xl overflow-hidden" style={{ border: '1px solid #e5e0d5' }}>

      {/* ── Header: Dark Serpent #133020 ── */}
      <div className="p-4 flex justify-between items-center" style={{ backgroundColor: '#133020', borderBottom: '1px solid #046241' }}>
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-full flex items-center justify-center text-white" style={{ backgroundColor: '#046241' }}>
            <Bot className="w-6 h-6" />
          </div>
          <div>
            <h1 className="font-bold text-white">Production Plan Agent</h1>
            <p className="text-xs flex items-center gap-1" style={{ color: '#FFB347' }}>
              <span className="w-2 h-2 rounded-full animate-pulse inline-block" style={{ backgroundColor: '#FFB347' }}></span>
              Powered by Gemini AI
            </p>
          </div>
        </div>
        <button
          onClick={resetChat}
          className="p-2 rounded-full transition-opacity hover:opacity-70"
          style={{ color: '#FFC370' }}
          title="Reset Chat"
        >
          <RefreshCw className="w-5 h-5" />
        </button>
      </div>

      {/* ── Messages: Paper #f5eedb ── */}
      <div className="flex-1 overflow-y-auto p-4 space-y-6" style={{ backgroundColor: '#f5eedb' }}>
        {messages.map((msg) => (
          <div key={msg.id} className={`flex gap-3 ${msg.role === 'user' ? 'flex-row-reverse' : ''}`}>

            {/* Avatar */}
            <div
              className="w-8 h-8 rounded-full flex items-center justify-center flex-shrink-0 text-white"
              style={{ backgroundColor: msg.role === 'agent' ? '#046241' : '#133020' }}
            >
              {msg.role === 'agent' ? <Bot className="w-5 h-5" /> : <UserIcon className="w-5 h-5" />}
            </div>

            <div className="max-w-[80%] space-y-2">
              {/* Bubble */}
              <div
                className="p-4 shadow-sm"
                style={
                  msg.role === 'agent'
                    ? { backgroundColor: '#ffffff', color: '#133020', borderRadius: '0 1rem 1rem 1rem', border: '1px solid #e5e0d5' }
                    : { backgroundColor: '#133020', color: '#ffffff', borderRadius: '1rem 0 1rem 1rem' }
                }
              >
                <div className="leading-relaxed prose prose-sm max-w-none">
                  <ReactMarkdown
                    components={{
                      p: ({ children }) => <p className="mb-2 last:mb-0">{children}</p>,
                      strong: ({ children }) => <strong className="font-semibold">{children}</strong>,
                      ul: ({ children }) => <ul className="list-disc list-inside mb-2 space-y-1">{children}</ul>,
                      ol: ({ children }) => <ol className="list-decimal list-inside mb-2 space-y-1">{children}</ol>,
                      li: ({ children }) => <li className="text-sm">{children}</li>,
                      code: ({ children }) => <code className="px-1 rounded text-xs font-mono" style={{ backgroundColor: '#F9F7F7', color: '#133020' }}>{children}</code>,
                    }}
                  >{msg.content}</ReactMarkdown>
                </div>
              </div>

              {/* Download — Saffron #FFC370 */}
              {msg.type === 'file' && msg.fileData && !isStreaming && (
                <button
                  onClick={() => handleDownload(msg.fileData!.name, msg.fileData!.buffer)}
                  className="flex items-center gap-3 p-4 rounded-xl w-full transition-opacity text-left hover:opacity-90"
                  style={{ backgroundColor: '#FFC370', border: '1px solid #FFB347' }}
                >
                  <div className="w-10 h-10 rounded-lg flex items-center justify-center" style={{ backgroundColor: 'rgba(255,255,255,0.3)' }}>
                    <FileSpreadsheet className="w-6 h-6" style={{ color: '#133020' }} />
                  </div>
                  <div className="flex-1">
                    <p className="font-medium" style={{ color: '#133020' }}>{msg.fileData.name}</p>
                    <p className="text-xs" style={{ color: '#046241' }}>Click to download</p>
                  </div>
                  <Download className="w-5 h-5" style={{ color: '#133020' }} />
                </button>
              )}
            </div>
          </div>
        ))}

        {/* Typing indicator */}
        {isTyping && (
          <div className="flex gap-3">
            <div className="w-8 h-8 rounded-full flex items-center justify-center text-white flex-shrink-0" style={{ backgroundColor: '#046241' }}>
              <Bot className="w-5 h-5" />
            </div>
            <div className="p-4 rounded-2xl shadow-sm flex items-center gap-2" style={{ backgroundColor: '#ffffff', border: '1px solid #e5e0d5' }}>
              <Loader2 className="w-4 h-4 animate-spin" style={{ color: '#046241' }} />
              <span className="text-sm" style={{ color: '#133020' }}>Agent is analyzing...</span>
            </div>
          </div>
        )}
        <div ref={messagesEndRef} />
      </div>

      {/* ── Input Area: White ── */}
      <div className="p-4 space-y-3" style={{ backgroundColor: '#ffffff', borderTop: '1px solid #e5e0d5' }}>
        {fileName && (
          <div className="flex items-center justify-between px-3 py-2 rounded-lg" style={{ backgroundColor: '#f5eedb', border: '1px solid #FFC370' }}>
            <div className="flex items-center gap-2 text-sm" style={{ color: '#046241' }}>
              <Paperclip className="w-4 h-4" />
              <span className="font-medium truncate max-w-[200px]">{fileName}</span>
            </div>
            <button onClick={() => { setFileName(null); setUploadedData(null); }} className="hover:opacity-70" style={{ color: '#FFB347' }}>
              <X className="w-4 h-4" />
            </button>
          </div>
        )}

        <div className="flex items-end gap-2">
          {/* Attach */}
          <button
            onClick={() => fileInputRef.current?.click()}
            className="p-3 mb-0.5 rounded-xl transition-opacity hover:opacity-70"
            style={{ color: '#046241' }}
            title="Upload actual data (CSV)"
          >
            <Paperclip className="w-5 h-5" />
          </button>
          <input type="file" ref={fileInputRef} onChange={handleFileUpload} accept=".csv" className="hidden" />

          {/* Textarea */}
          <textarea
            ref={textareaRef}
            rows={1}
            value={inputValue}
            onChange={(e) => setInputValue(e.target.value)}
            onKeyDown={handleKeyDown}
            placeholder="Describe your project..."
            disabled={isTyping || isStreaming}
            className="flex-1 px-4 py-3 rounded-xl outline-none transition-all disabled:opacity-50 disabled:cursor-not-allowed resize-none overflow-y-auto max-h-[200px]"
            style={{ backgroundColor: '#F9F7F7', border: '1px solid #e5e0d5', color: '#133020' }}
            onFocus={e => (e.currentTarget.style.borderColor = '#046241')}
            onBlur={e => (e.currentTarget.style.borderColor = '#e5e0d5')}
          />

          {/* Send — Castleton Green #046241 */}
          <button
            onClick={handleSendMessage}
            disabled={!inputValue.trim() || isTyping || isStreaming}
            className="p-3 mb-0.5 rounded-xl transition-opacity shadow-sm disabled:opacity-50 disabled:cursor-not-allowed text-white"
            style={{ backgroundColor: '#046241' }}
            onMouseEnter={e => (e.currentTarget.style.backgroundColor = '#133020')}
            onMouseLeave={e => (e.currentTarget.style.backgroundColor = '#046241')}
          >
            <Send className="w-5 h-5" />
          </button>
        </div>
      </div>
    </div>
  );
}