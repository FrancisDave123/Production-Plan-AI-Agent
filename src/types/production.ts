import ExcelJS from 'exceljs';

export interface Message {
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
        data: string; // Base64 for images
    };
}

export interface ActualDataItem {
    date: string;
    name: string;
    actual: number;
    [key: string]: any;
}

export interface ProjectColumn {
    header: string;
    key: string;
    section: 'Target' | 'Actual' | 'Accumulative';
    formula?: string; // Excel formula with {rowIndex} placeholder
    width?: number;
}

export interface DailyColumn {
    header: string;
    key: string;
    formula?: string;
}

export interface ProjectData {
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

export interface FileAttachment {
    name: string;
    type: string;
    data: string;
    file: File;
    metadata?: string;
    parsedData?: ActualDataItem[];
}
