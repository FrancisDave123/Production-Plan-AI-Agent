import React, { useState, useEffect, useRef } from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { format, eachDayOfInterval, isValid, isSameDay } from 'date-fns';
import { Send, Bot, User as UserIcon, RefreshCw, FileSpreadsheet, Loader2, Download, Paperclip, X } from 'lucide-react';
import { GoogleGenAI, Type } from "@google/genai";
import Papa from 'papaparse';

interface Message {
  id: string;
  role: 'agent' | 'user';
  content: string;
  type?: 'text' | 'file';
  fileData?: {
    name: string;
    buffer: ExcelJS.Buffer;
  };
  attachment?: {
    name: string;
    type: string;
    data: string; // Base64 for images, or other data
  };
}

interface ActualDataItem {
  date: string;
  name: string;
  actual: number;
  [key: string]: any; // Allow for extra custom fields like 'duration', 'ops', etc.
}

interface ProjectColumn {
  header: string;
  key: string;
  section: 'Target' | 'Actual' | 'Accumulative';
  formula?: string; // Excel formula with {rowIndex} placeholder
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
  const [uploadedData, setUploadedData] = useState<ActualDataItem[] | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [currentFile, setCurrentFile] = useState<{ name: string; type: string; data: string; file: File; metadata?: string } | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [currentProject, setCurrentProject] = useState<Partial<ProjectData> | null>(null);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const chatRef = useRef<any>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (!chatRef.current) {
      chatRef.current = ai.chats.create({
        model: "gemini-3-flash-preview",
        config: {
          systemInstruction: `You are a professional Production Planning Assistant. 
          Your goal is to collect information from the user to generate a comprehensive Excel production plan.
          
          CONVERSATION GUIDELINES:
          - Be professional, concise, and helpful.
          - Use bolding for key terms and lists for multiple items.
          - If the user provides a file, acknowledge the data and use it in your suggestions.
          - Always confirm the structure before calling the tool.
          
          REQUIRED PROJECT DETAILS:
          1. Project Name
          2. Overall Goal (numeric value)
          3. Unit (e.g., units, hours)
          4. Start Date (YYYY-MM-DD)
          5. End Date (YYYY-MM-DD)
          6. Resources/Teams
          
          DYNAMIC SCHEMA ANALYSIS:
          - Suggest DAILY KEY COLUMNS (Target, Actual, and others like Variance G{rowIndex}-F{rowIndex}).
          - Suggest PLAN summary columns.
          - Suggest PIVOT columns for weekly/monthly summaries.
          - Suggest DASHBOARD metrics for high-level KPIs.
          
          The raw data table name is 'DailyProductionTable'.
          Base columns: [Date(A), Day(B), Week(C), Month(D), Name(E)].
          Daily columns start at Column F (Index 6).
          
          IMPORTANT: Only help with production planning and Excel generation. Politey decline unrelated queries.`,
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

  const processFile = async (file: File) => {
    setFileName(file.name);
    const fileType = file.type;

    if (fileType === 'text/csv' || file.name.endsWith('.csv')) {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (results) => {
          const parsedData: ActualDataItem[] = results.data.map((row: any) => ({
            date: row.Date || row.date || '',
            name: row.Name || row.name || '',
            actual: parseFloat(row.Actual || row.actual || '0')
          })).filter(item => item.date && item.name);

          // Generate a markdown table snippet
          const snippet = parsedData.slice(0, 5).map(row =>
            `| ${row.date} | ${row.name} | ${row.actual} |`
          ).join('\n');
          const dataAwareContent = `\n\n**Data Preview (from ${file.name}):**\n| Date | Name | Actual |\n|---|---|---|\n${snippet}\n\nTotal rows: ${parsedData.length}`;

          setUploadedData(parsedData);
          setCurrentFile({
            name: file.name,
            type: fileType,
            data: '', // Not needed for CSV
            file: file,
            metadata: dataAwareContent
          });
        },
        error: (error) => {
          console.error("CSV Parse Error:", error);
          alert("Error parsing CSV file. Please ensure it has Date, Name, and Actual columns.");
        }
      });
    } else if (fileType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
      const reader = new FileReader();
      reader.onload = async (e) => {
        const buffer = e.target?.result as ArrayBuffer;
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);
        const worksheet = workbook.getWorksheet(1);
        if (!worksheet) return;

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

        // Generate a markdown table snippet
        const snippet = parsedData.slice(0, 5).map(row =>
          `| ${row.date} | ${row.name} | ${row.actual} |`
        ).join('\n');
        const dataAwareContent = `\n\n**Data Preview (from ${file.name}):**\n| Date | Name | Actual |\n|---|---|---|\n${snippet}\n\nTotal rows: ${parsedData.length}`;

        setUploadedData(parsedData);
        setCurrentFile({
          name: file.name,
          type: fileType,
          data: '', // Not needed for Excel
          file: file,
          metadata: dataAwareContent
        });
      };
      reader.readAsArrayBuffer(file);
    } else if (fileType.startsWith('image/')) {
      const reader = new FileReader();
      reader.onload = (e) => {
        const base64Data = e.target?.result as string;
        setCurrentFile({
          name: file.name,
          type: fileType,
          data: base64Data,
          file: file
        });
      };
      reader.readAsDataURL(file);
    } else {
      alert("Unsupported file type. Please upload CSV, Excel, or Image files.");
      setFileName(null);
    }
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    await processFile(file);
    e.target.value = '';
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
  };

  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);

    const file = e.dataTransfer.files?.[0];
    if (file) {
      await processFile(file);
    }
  };

  const handleSendMessage = async () => {
    if ((!inputValue.trim() && !currentFile) || isTyping) return;

    // Generate context preamble
    const contextPreamble = currentProject ?
      `[CURRENT PROJECT STATE: Name="${currentProject.name || '?'}", Goal=${currentProject.goal || '?'}, Unit="${currentProject.unit || '?'}", Dates=${currentProject.startDate || '?'}/${currentProject.endDate || '?'}, Resources=${currentProject.resources?.join(',') || '?'}]\n`
      : "";

    const fileMetadata = currentFile?.metadata || "";
    const fullPrompt = `${contextPreamble}${inputValue}${fileMetadata}`;

    const userMsg: Message = {
      id: Date.now().toString(),
      role: 'user',
      content: inputValue || (currentFile ? `Shared ${currentFile.type.startsWith('image/') ? 'an image' : 'a file'}: ${currentFile.name}` : ""),
      attachment: currentFile ? {
        name: currentFile.name,
        type: currentFile.type,
        data: currentFile.data
      } : undefined
    };

    setMessages(prev => [...prev, userMsg]);
    setInputValue('');
    setCurrentFile(null);
    setFileName(null);
    setIsTyping(true);

    try {
      let result;
      if (userMsg.attachment && userMsg.attachment.type.startsWith('image/')) {
        const base64Data = userMsg.attachment.data.split(',')[1];
        result = await chatRef.current.sendMessage({
          message: [
            { text: fullPrompt },
            {
              inlineData: {
                data: base64Data,
                mimeType: userMsg.attachment.type
              }
            }
          ]
        });
      } else {
        result = await chatRef.current.sendMessage({ message: fullPrompt });
      }

      const response = result;

      if (response.functionCalls) {
        for (const call of response.functionCalls) {
          if (call.name === 'generate_production_plan') {
            const projectData = call.args as ProjectData;
            setCurrentProject(projectData); // Update current project state

            // Merge actual data from both sources
            const combinedActualData = [...(projectData.actualData || [])];

            if (uploadedData) {
              uploadedData.forEach(upItem => {
                // Avoid duplicates if same date/name exists in both
                const exists = combinedActualData.some(combItem =>
                  combItem.date === upItem.date && combItem.name === upItem.name
                );
                if (!exists) {
                  combinedActualData.push(upItem);
                }
              });
            }

            projectData.actualData = combinedActualData.length > 0 ? combinedActualData : undefined;
            await generateExcelFile(projectData);
          }
        }
      } else {
        setMessages(prev => [...prev, {
          id: Date.now().toString(),
          role: 'agent',
          content: response.text || "I'm sorry, I didn't quite get that. Could you please provide the project details?"
        }]);
      }
    } catch (error) {
      console.error("Gemini Error:", error);
      setMessages(prev => [...prev, {
        id: Date.now().toString(),
        role: 'agent',
        content: "I'm having a bit of trouble connecting to my brain. Could you try again?"
      }]);
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
      // We calculate a weight for each item based on its chronological position
      // Learning (0-25%): 30% -> 60% weight
      // Progress (25-75%): 60% -> 100% weight
      // Behavior (75-100%): 100% weight
      const weights = scheduleItems.map((_, index) => {
        const t = index / (totalItems - 1 || 1);
        if (t < 0.25) {
          // Learning Phase: Linear ramp from 0.3 to 0.6
          return 0.3 + (0.6 - 0.3) * (t / 0.25);
        } else if (t < 0.75) {
          // Progress Phase: Linear ramp from 0.6 to 1.0
          return 0.6 + (1.0 - 0.6) * ((t - 0.25) / 0.5);
        } else {
          // Behavior Phase: Stable at 1.0 (peak performance)
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

      // Apply conditional formatting for whole vs decimal numbers
      const keyRows = scheduleItems.length + 1;
      const lastKeyColLetter = getColumnLetter(sheetKey.columns.length);
      sheetKey.addConditionalFormatting({
        ref: `F2:${lastKeyColLetter}${keyRows}`,
        rules: [
          {
            priority: 1,
            type: 'expression',
            formulae: ['MOD(F2,1)=0'],
            style: { numFmt: '#,##0' },
          },
          {
            priority: 2,
            type: 'expression',
            formulae: ['MOD(F2,1)<>0'],
            style: { numFmt: '#,##0.00' },
          },
        ],
      });

      // --- Sheet 2: Production Plan ---
      const sheetPlan = workbook.addWorksheet(sanitizeSheetName(`${projectData.name} Plan`));

      const unit = projectData.unit || 'Units';
      const unitLabel = unit.charAt(0).toUpperCase() + unit.slice(1);

      // Define Columns from projectData.columns
      // We always start with Date and Month
      const dynamicColumns: ProjectColumn[] = [
        { header: 'Date', key: 'date', width: 15, section: 'Target' as const },
        { header: 'Month', key: 'month', width: 15, section: 'Target' as const },
        ...projectData.columns
      ];

      sheetPlan.columns = dynamicColumns.map(col => ({
        key: col.key,
        width: col.width || 18
      }));

      // --- Header Row 1: Project Title ---
      const totalCols = dynamicColumns.length;
      const lastColLetter = getColumnLetter(totalCols);
      sheetPlan.mergeCells(`A1:${lastColLetter}1`);
      const titleCell = sheetPlan.getCell('A1');
      titleCell.value = `${projectData.name}: Production Plan & Daily Output Tracking`;
      titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF006633' } };
      titleCell.font = { color: { argb: 'FFFFFFFF' }, bold: true, size: 12 };
      titleCell.alignment = { horizontal: 'center', vertical: 'middle' };

      // --- Header Row 2 & 3: Sections ---
      // We need to group columns by section to merge them
      const sections = ['Target', 'Actual', 'Accumulative'] as const;
      let currentColIndex = 1;

      sections.forEach(section => {
        const sectionCols = dynamicColumns.filter(c => c.section === section);
        if (sectionCols.length > 0) {
          const startCol = getColumnLetter(currentColIndex);
          const endCol = getColumnLetter(currentColIndex + sectionCols.length - 1);

          // Row 2: Main Section
          const ref2 = `${startCol}2:${endCol}2`;
          sheetPlan.mergeCells(ref2);
          const cell2 = sheetPlan.getCell(`${startCol}2`);
          cell2.value = section === 'Target' ? `Target ${unitLabel} Output` : (section === 'Actual' ? `${unitLabel} Output Tracking` : section);

          // Styling based on section
          if (section === 'Target') {
            cell2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
          } else if (section === 'Actual' || section === 'Accumulative') {
            cell2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF70AD47' } };
            cell2.font = { color: { argb: 'FFFFFFFF' }, bold: true };
          }
          cell2.alignment = { horizontal: 'center', vertical: 'middle' };

          // Row 3: Sub Section (only for Actual/Accumulative based on user example)
          if (section === 'Actual' || section === 'Accumulative') {
            const ref3 = `${startCol}3:${endCol}3`;
            sheetPlan.mergeCells(ref3);
            const cell3 = sheetPlan.getCell(`${startCol}3`);
            cell3.value = section;
            cell3.fill = { type: 'pattern', pattern: 'solid', fgColor: section === 'Actual' ? { argb: 'FFC6E0B4' } : { argb: 'FFD9D9D9' } };
            cell3.font = { bold: true };
            cell3.alignment = { horizontal: 'center', vertical: 'middle' };
          } else {
            // For Target, just merge Row 2 and 3
            const ref23 = `${startCol}2:${endCol}3`;
            sheetPlan.unMergeCells(ref2);
            sheetPlan.mergeCells(ref23);
          }

          currentColIndex += sectionCols.length;
        }
      });

      // --- Header Row 4: Column Names ---
      const headerRow4 = sheetPlan.getRow(4);
      dynamicColumns.forEach((col, i) => {
        const cell = headerRow4.getCell(i + 1);
        cell.value = col.header;
        cell.font = { bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
        };

        // Background colors for Row 4 based on sections
        if (col.section === 'Target') cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
        else if (col.section === 'Actual') cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6E0B4' } };
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

          if (col.key === 'date') {
            cell.value = dateObj;
          } else if (col.key === 'month') {
            cell.value = { formula: `TEXT(A${rowIndex}, "mmmm")` };
          } else if (col.formula) {
            // Replace {rowIndex} placeholder in formula
            const finalFormula = col.formula.replace(/{rowIndex}/g, rowIndex.toString());
            cell.value = { formula: finalFormula };
          }

          // Apply borders
          cell.border = {
            top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
          };

          // Specific formatting
          if (col.header.toLowerCase().includes('rate') || col.header.toLowerCase().includes('%')) {
            cell.numFmt = '0.00%';
          }
        });
      });

      // --- Grand Total Row for Production Plan ---
      const totalRowIndex = uniqueDates.length + 5;
      const totalRow = sheetPlan.getRow(totalRowIndex);
      totalRow.getCell(1).value = 'Grand Total';
      totalRow.font = { bold: true };
      dynamicColumns.forEach((col, colIdx) => {
        if (colIdx === 0) return; // Skip 'Date'
        const cell = totalRow.getCell(colIdx + 1);
        const colLetter = getColumnLetter(colIdx + 1);

        // Sum numeric columns, skip month/rates
        const isRate = col.header.toLowerCase().includes('rate') || col.header.toLowerCase().includes('%');
        const isMonth = col.key === 'month';

        if (!isMonth && !isRate) {
          cell.value = { formula: `SUM(${colLetter}5:${colLetter}${totalRowIndex - 1})` };
        } else if (isRate) {
          // For rates, we might want a weighted average or just leave blank
          // User example shows single values, but for grand total we'll leave blank or use a specific formula if provided
        }

        cell.border = {
          top: { style: 'double' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
        };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
      });

      const planRows = uniqueDates.length + 5; // Including total row
      // Apply conditional formatting to all numeric columns in the data range
      sheetPlan.addConditionalFormatting({
        ref: `C5:${lastColLetter}${planRows}`,
        rules: [
          {
            priority: 1,
            type: 'expression',
            formulae: ['MOD(C5,1)=0'],
            style: { numFmt: '#,##0' },
          },
          {
            priority: 2,
            type: 'expression',
            formulae: ['MOD(C5,1)<>0'],
            style: { numFmt: '#,##0.00' },
          },
        ],
      });

      // --- Sheet 3: Pivot Summary ---

      // --- Sheet 3: Pivot Summary ---
      const sheetPivot = workbook.addWorksheet(sanitizeSheetName('Production_Pivot'));

      const basePivotCols = [
        { header: 'Week', key: 'week', width: 10 },
        { header: 'Month', key: 'month', width: 15 },
      ];

      const dynamicPivotCols = (projectData.pivotColumns || [
        { header: 'Total Target', formula: `SUMIFS(DailyProductionTable[Target], DailyProductionTable[Week], A{rowIndex})` },
        { header: 'Total Actual', formula: `SUMIFS(DailyProductionTable[Actual], DailyProductionTable[Week], A{rowIndex})` },
        { header: 'Total Variance', formula: `SUMIFS(DailyProductionTable[Variance], DailyProductionTable[Week], A{rowIndex})` },
        { header: 'Cumulative Actual', formula: `SUM($D$2:D{rowIndex})` },
      ])
        .filter(col => !basePivotCols.some(bp => bp.header === col.header))
        .map(col => ({
          header: col.header,
          formula: col.formula,
          width: 18
        }));

      sheetPivot.columns = [...basePivotCols, ...dynamicPivotCols.map(c => ({ header: c.header, key: c.header.replace(/\s+/g, ''), width: c.width }))];

      const weeks = Array.from(new Set(scheduleItems.map(s => format(s.date, 'w')))).sort((a, b) => Number(a) - Number(b));

      sheetPivot.addTable({
        name: 'PivotSummaryTable',
        ref: 'A1',
        headerRow: true,
        totalsRow: true,
        style: { theme: 'TableStyleLight1', showRowStripes: true },
        columns: [
          { name: 'Week', filterButton: true },
          { name: 'Month', filterButton: true },
          ...dynamicPivotCols.map(col => {
            const isRate = col.header.toLowerCase().includes('rate') || col.header.toLowerCase().includes('%');
            return {
              name: col.header,
              filterButton: true,
              totalsRowFunction: (isRate ? 'none' : 'sum') as any
            };
          })
        ],
        rows: weeks.map((week, index) => {
          const rowIndex = index + 2;
          const row = [
            Number(week),
            { formula: `INDEX(DailyProductionTable[Month], MATCH(${week}, DailyProductionTable[Week], 0))` }
          ];
          dynamicPivotCols.forEach(col => {
            const finalFormula = col.formula.replace(/{rowIndex}/g, rowIndex.toString());
            row.push({ formula: finalFormula } as any);
          });
          return row;
        })
      });

      const pivotRows = weeks.length + 1;
      const lastPivotColLetter = getColumnLetter(sheetPivot.columns.length);
      sheetPivot.addConditionalFormatting({
        ref: `C2:${lastPivotColLetter}${pivotRows}`,
        rules: [
          {
            priority: 1,
            type: 'expression',
            formulae: ['MOD(C2,1)=0'],
            style: { numFmt: '#,##0' },
          },
          {
            priority: 2,
            type: 'expression',
            formulae: ['MOD(C2,1)<>0'],
            style: { numFmt: '#,##0.00' },
          },
        ],
      });

      // --- Sheet 4: Dashboard ---
      const sheetDash = workbook.addWorksheet(sanitizeSheetName('Summary_Dashboard'));
      sheetDash.mergeCells('A1:B1');
      sheetDash.getCell('A1').value = 'Project Summary Dashboard';
      sheetDash.getCell('A1').font = { bold: true, size: 16 };
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
        const cell = sheetDash.getCell(`B${r}`);
        cell.value = { formula: item.formula };

        if (item.format) {
          cell.numFmt = item.format;
        } else {
          sheetDash.addConditionalFormatting({
            ref: `B${r}`,
            rules: [
              {
                priority: 1,
                type: 'expression',
                formulae: [`MOD(B${r},1)=0`],
                style: { numFmt: '#,##0' },
              },
              {
                priority: 2,
                type: 'expression',
                formulae: [`MOD(B${r},1)<>0`],
                style: { numFmt: '#,##0.00' },
              },
            ],
          });
        }
        sheetDash.getCell(`A${r}`).font = { bold: true };
      });
      sheetDash.getColumn(1).width = 30;
      sheetDash.getColumn(2).width = 25;

      const buffer = await workbook.xlsx.writeBuffer();

      setMessages(prev => [...prev, {
        id: Date.now().toString(),
        role: 'agent',
        content: `I've generated the production plan for **${projectData.name}**. You can download it below.`,
        type: 'file',
        fileData: {
          name: `${projectData.name.replace(/\s+/g, '_')}_Production_Planning.xlsx`,
          buffer: buffer as ExcelJS.Buffer
        }
      }]);

    } catch (error) {
      console.error(error);
      setMessages(prev => [...prev, {
        id: Date.now().toString(),
        role: 'agent',
        content: "I encountered an error generating the Excel file. Please check the details and try again."
      }]);
    }
  };

  const handleDownload = (fileName: string, buffer: ExcelJS.Buffer) => {
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, fileName);
  };

  const resetChat = () => {
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

  const handlePaste = async (e: React.ClipboardEvent) => {
    const items = e.clipboardData.items;
    for (let i = 0; i < items.length; i++) {
      if (items[i].type.startsWith('image/')) {
        const file = items[i].getAsFile();
        if (file) {
          // Rename pasted files to something descriptive
          const pastedFile = new File([file], `screenshot-${Date.now()}.png`, { type: file.type });
          await processFile(pastedFile);
        }
      }
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
    <div
      className={`max-w-2xl mx-auto h-[700px] flex flex-col bg-white rounded-2xl shadow-xl border border-gray-100 overflow-hidden relative ${isDragging ? 'ring-4 ring-blue-500 ring-inset' : ''}`}
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
    >
      {isDragging && (
        <div className="absolute inset-0 z-50 bg-blue-600/10 backdrop-blur-sm flex items-center justify-center pointer-events-none">
          <div className="bg-white p-8 rounded-3xl shadow-2xl border-2 border-blue-500 border-dashed flex flex-col items-center gap-4 animate-in zoom-in-95 duration-200">
            <div className="w-16 h-16 bg-blue-100 rounded-full flex items-center justify-center text-blue-600">
              <Download className="w-8 h-8 animate-bounce" />
            </div>
            <div className="text-center">
              <p className="text-xl font-bold text-gray-900">Drop files here</p>
              <p className="text-sm text-gray-500 mt-1">CSV, Excel, or Images</p>
            </div>
          </div>
        </div>
      )}
      {/* Header */}
      <div className="bg-white border-b border-gray-100 p-4 flex justify-between items-center">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 bg-blue-600 rounded-full flex items-center justify-center text-white">
            <Bot className="w-6 h-6" />
          </div>
          <div>
            <h1 className="font-bold text-gray-900">Production Plan Agent</h1>
            <p className="text-xs text-green-600 flex items-center gap-1">
              <span className="w-2 h-2 bg-green-500 rounded-full animate-pulse"></span>
              Powered by Gemini AI
            </p>
          </div>
        </div>
        <button
          onClick={resetChat}
          className="p-2 text-gray-400 hover:text-gray-600 hover:bg-gray-50 rounded-full transition-colors"
          title="Reset Chat"
        >
          <RefreshCw className="w-5 h-5" />
        </button>
      </div>

      {/* Messages */}
      <div className="flex-1 overflow-y-auto p-4 space-y-6 bg-gray-50/50">
        {messages.map((msg) => (
          <div
            key={msg.id}
            className={`flex gap-3 ${msg.role === 'user' ? 'flex-row-reverse' : ''}`}
          >
            <div className={`w-8 h-8 rounded-full flex items-center justify-center flex-shrink-0 ${msg.role === 'agent' ? 'bg-blue-600 text-white' : 'bg-gray-900 text-white'
              }`}>
              {msg.role === 'agent' ? <Bot className="w-5 h-5" /> : <UserIcon className="w-5 h-5" />}
            </div>

            <div className={`max-w-[80%] space-y-2`}>
              <div className={`p-4 rounded-2xl shadow-sm ${msg.role === 'agent'
                ? 'bg-white text-gray-800 rounded-tl-none border border-gray-100'
                : 'bg-gray-900 text-white rounded-tr-none'
                }`}>
                <p className="whitespace-pre-wrap leading-relaxed">{msg.content}</p>
                {msg.attachment && (
                  <div className="mt-2 space-y-2">
                    {msg.attachment.type.startsWith('image/') ? (
                      <div className="rounded-lg overflow-hidden border border-gray-200 bg-white">
                        <img src={msg.attachment.data} alt={msg.attachment.name} className="max-w-full h-auto max-h-64 object-contain mx-auto" />
                      </div>
                    ) : (
                      <div className={`flex items-center gap-3 p-3 rounded-xl border ${msg.role === 'user' ? 'bg-white/10 border-white/20 text-white' : 'bg-gray-50 border-gray-100 text-gray-700'}`}>
                        <div className={`w-10 h-10 rounded-lg flex items-center justify-center ${msg.role === 'user' ? 'bg-white/20' : 'bg-white shadow-sm text-blue-600'}`}>
                          <FileSpreadsheet className="w-6 h-6" />
                        </div>
                        <div className="flex-1 overflow-hidden">
                          <p className="text-sm font-bold truncate">{msg.attachment.name}</p>
                          <p className={`text-[10px] uppercase font-black tracking-wider opacity-60`}>
                            {msg.attachment.type.includes('csv') ? 'CSV FILE' : 'EXCEL DOCUMENT'}
                          </p>
                        </div>
                      </div>
                    )}
                  </div>
                )}
              </div>

              {msg.type === 'file' && msg.fileData && (
                <button
                  onClick={() => handleDownload(msg.fileData!.name, msg.fileData!.buffer)}
                  className="flex items-center gap-3 bg-green-50 border border-green-100 p-4 rounded-xl w-full hover:bg-green-100 transition-colors group text-left"
                >
                  <div className="w-10 h-10 bg-green-100 group-hover:bg-green-200 rounded-lg flex items-center justify-center text-green-600 transition-colors">
                    <FileSpreadsheet className="w-6 h-6" />
                  </div>
                  <div className="flex-1">
                    <p className="font-medium text-green-900">{msg.fileData.name}</p>
                    <p className="text-xs text-green-700">Click to download</p>
                  </div>
                  <Download className="w-5 h-5 text-green-600" />
                </button>
              )}
            </div>
          </div>
        ))}

        {isTyping && (
          <div className="flex gap-3">
            <div className="w-8 h-8 bg-blue-600 rounded-full flex items-center justify-center text-white flex-shrink-0">
              <Bot className="w-5 h-5" />
            </div>
            <div className="bg-white border border-gray-100 p-4 rounded-2xl rounded-tl-none shadow-sm flex items-center gap-2">
              <Loader2 className="w-4 h-4 animate-spin text-blue-600" />
              <span className="text-sm text-gray-500">Agent is analyzing...</span>
            </div>
          </div>
        )}
        <div ref={messagesEndRef} />
      </div>

      {/* Input */}
      <div className="p-4 bg-white border-t border-gray-100 space-y-3">
        {fileName && (
          <div className="flex items-center justify-between bg-blue-50 px-3 py-2 rounded-lg border border-blue-100 animate-in fade-in slide-in-from-bottom-2 duration-200">
            <div className="flex items-center gap-2 text-sm text-blue-700">
              {currentFile?.type.startsWith('image/') ? (
                <div className="w-10 h-10 rounded border border-blue-200 overflow-hidden bg-white shadow-sm">
                  <img src={currentFile.data} alt="preview" className="w-full h-full object-cover" />
                </div>
              ) : (
                <div className="w-10 h-10 rounded border border-blue-200 flex items-center justify-center bg-white shadow-sm text-blue-600">
                  <FileSpreadsheet className="w-6 h-6" />
                </div>
              )}
              <div className="flex flex-col">
                <span className="font-bold truncate max-w-[200px]">{fileName}</span>
                <span className="text-[10px] opacity-70 uppercase font-black tracking-wider">
                  {currentFile?.type.startsWith('image/') ? 'Image Attachment' : 'Data File (Ready to Analyze)'}
                </span>
              </div>
            </div>
            <button
              onClick={() => { setFileName(null); setUploadedData(null); setCurrentFile(null); }}
              className="p-1.5 text-blue-400 hover:text-blue-600 hover:bg-blue-100 rounded-full transition-colors"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        )}
        <div className="flex items-end gap-2">
          <button
            onClick={() => fileInputRef.current?.click()}
            className="p-3 mb-0.5 text-gray-400 hover:text-blue-600 hover:bg-blue-50 rounded-xl transition-all"
            title="Upload data (CSV, Excel) or Images"
          >
            <Paperclip className="w-5 h-5" />
          </button>
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileUpload}
            accept=".csv, .xlsx, .xls, image/*"
            className="hidden"
          />
          <textarea
            ref={textareaRef}
            rows={1}
            value={inputValue}
            onChange={(e) => setInputValue(e.target.value)}
            onKeyDown={handleKeyDown}
            onPaste={handlePaste}
            placeholder="Describe your project..."
            disabled={isTyping}
            className="flex-1 px-4 py-3 bg-gray-50 border border-gray-200 rounded-xl focus:ring-2 focus:ring-blue-500 focus:bg-white outline-none transition-all disabled:opacity-50 disabled:cursor-not-allowed resize-none overflow-y-auto max-h-[200px]"
          />
          <button
            onClick={handleSendMessage}
            disabled={(!inputValue.trim() && !currentFile) || isTyping}
            className={`p-3 mb-0.5 rounded-xl transition-all shadow-sm ${(inputValue.trim() || currentFile) && !isTyping
              ? 'bg-blue-600 text-white hover:bg-blue-700'
              : 'bg-gray-100 text-gray-300 cursor-not-allowed'
              }`}
          >
            <Send className="w-5 h-5" />
          </button>
        </div>
      </div>
    </div>
  );
}
