import React, { useState, useEffect, useRef } from 'react';
import { saveAs } from 'file-saver';
import { Bot, RefreshCw, Download } from 'lucide-react';
import { GoogleGenAI, Type } from "@google/genai";

// Modular Imports
import { Message, ProjectData, ActualDataItem, FileAttachment } from '../types/production';
import { generateExcelFile } from '../utils/excelGenerator';
import { handleFileProcessing } from '../utils/fileHandlers';
import ChatMessage from './chat/ChatMessage';
import ChatInput from './chat/ChatInput';
import FilePreview from './chat/FilePreview';

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

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
  const [currentFile, setCurrentFile] = useState<FileAttachment | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [currentProject, setCurrentProject] = useState<Partial<ProjectData> | null>(null);

  const messagesEndRef = useRef<HTMLDivElement>(null);
  const chatRef = useRef<any>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const initChat = () => {
    chatRef.current = ai.chats.create({
      model: "gemini-2.0-flash",
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
        - Suggest DAILY KEY COLUMNS (Target, Actual, and others).
        - Suggest PLAN summary columns.
        - Suggest PIVOT columns for weekly/monthly summaries.
        - Suggest DASHBOARD metrics for high-level KPIs.
        
        Raw data table: 'DailyProductionTable'. Base columns: [Date, Day, Week, Month, Name].
        Daily columns start at Column F (Index 6).`,
        tools: [{
          functionDeclarations: [{
            name: "generate_production_plan",
            description: "Generates the production planning Excel file once core project details and column structure are confirmed.",
            parameters: {
              type: Type.OBJECT,
              properties: {
                name: { type: Type.STRING },
                goal: { type: Type.NUMBER },
                unit: { type: Type.STRING },
                startDate: { type: Type.STRING },
                endDate: { type: Type.STRING },
                resources: { type: Type.ARRAY, items: { type: Type.STRING } },
                columns: {
                  type: Type.ARRAY,
                  items: {
                    type: Type.OBJECT,
                    properties: {
                      header: { type: Type.STRING },
                      key: { type: Type.STRING },
                      section: { type: Type.STRING, enum: ["Target", "Actual", "Accumulative"] },
                      formula: { type: Type.STRING }
                    },
                    required: ["header", "key", "section", "formula"]
                  }
                },
                dailyColumns: {
                  type: Type.ARRAY,
                  items: {
                    type: Type.OBJECT,
                    properties: {
                      header: { type: Type.STRING },
                      key: { type: Type.STRING },
                      formula: { type: Type.STRING }
                    },
                    required: ["header", "key"]
                  }
                },
                pivotColumns: {
                  type: Type.ARRAY,
                  items: {
                    type: Type.OBJECT,
                    properties: {
                      header: { type: Type.STRING },
                      formula: { type: Type.STRING }
                    },
                    required: ["header", "formula"]
                  }
                },
                dashboardMetrics: {
                  type: Type.ARRAY,
                  items: {
                    type: Type.OBJECT,
                    properties: {
                      label: { type: Type.STRING },
                      formula: { type: Type.STRING },
                      format: { type: Type.STRING }
                    },
                    required: ["label", "formula"]
                  }
                },
                actualData: {
                  type: Type.ARRAY,
                  items: {
                    type: Type.OBJECT,
                    properties: {
                      date: { type: Type.STRING },
                      name: { type: Type.STRING },
                      actual: { type: Type.NUMBER }
                    },
                    required: ["date", "name", "actual"]
                  }
                }
              },
              required: ["name", "goal", "unit", "startDate", "endDate", "resources", "columns"]
            }
          }]
        }]
      }
    });
  };

  // Initialize Gemini Chat
  useEffect(() => {
    if (!chatRef.current) {
      initChat();
    }
  }, []);

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages, isTyping]);

  const processFile = async (file: File) => {
    try {
      const processed = await handleFileProcessing(file);
      setFileName(processed.name!);
      if (processed.parsedData) {
        setUploadedData(processed.parsedData);
      }
      setCurrentFile(processed as FileAttachment);
    } catch (error) {
      alert(error instanceof Error ? error.message : "Error processing file.");
    }
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    await processFile(file);
    e.target.value = '';
  };

  const handleSendMessage = async () => {
    if ((!inputValue.trim() && !currentFile) || isTyping) return;

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
      if (!chatRef.current) initChat();

      let result;
      if (userMsg.attachment && userMsg.attachment.type.startsWith('image/')) {
        const base64Data = userMsg.attachment.data.split(',')[1];
        result = await chatRef.current!.sendMessage({
          message: [
            { text: fullPrompt },
            { inlineData: { data: base64Data, mimeType: userMsg.attachment.type } }
          ]
        });
      } else {
        result = await chatRef.current!.sendMessage({ message: fullPrompt });
      }

      const functionCalls = result?.functionCalls;
      if (functionCalls && functionCalls.length > 0) {
        for (const call of functionCalls) {
          if (call.name === 'generate_production_plan') {
            const projectData = call.args as ProjectData;
            setCurrentProject(projectData);

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
            const buffer = await generateExcelFile(projectData);

            setMessages(prev => [...prev, {
              id: Date.now().toString(),
              role: 'agent',
              content: `I've generated the production plan for **${projectData.name}**. You can download it below.`,
              type: 'file',
              fileData: {
                name: `${projectData.name.replace(/\s+/g, '_')}_Production_Planning.xlsx`,
                buffer: buffer
              }
            }]);
          }
        }
      } else {
        const textResponse = result?.text;
        setMessages(prev => [...prev, {
          id: Date.now().toString(),
          role: 'agent',
          content: textResponse || "I'm sorry, I didn't quite get that."
        }]);
      }
    } catch (error) {
      console.error("Gemini Error:", error);
      setMessages(prev => [...prev, {
        id: Date.now().toString(),
        role: 'agent',
        content: "I'm having a bit of trouble. Could you try again?"
      }]);
    } finally {
      setIsTyping(false);
    }
  };

  const handleDownload = (fileName: string, buffer: any) => {
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, fileName);
  };

  const resetChat = () => {
    setMessages([{
      id: '1',
      role: 'agent',
      content: "Hello! I'm your Production Plan Agent. I can help you create a detailed Excel production plan."
    }]);
    setUploadedData(null);
    setFileName(null);
    setCurrentProject(null);
    initChat();
  };

  return (
    <div
      className={`max-w-2xl mx-auto h-[700px] flex flex-col bg-white rounded-2xl shadow-xl border border-gray-100 overflow-hidden relative ${isDragging ? 'ring-4 ring-blue-500 ring-inset' : ''}`}
      onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
      onDragLeave={() => setIsDragging(false)}
      onDrop={async (e) => {
        e.preventDefault();
        setIsDragging(false);
        const file = e.dataTransfer.files?.[0];
        if (file) await processFile(file);
      }}
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
        <button onClick={resetChat} className="p-2 text-gray-400 hover:text-gray-600 hover:bg-gray-50 rounded-full transition-colors">
          <RefreshCw className="w-5 h-5" />
        </button>
      </div>

      {/* Messages */}
      <div className="flex-1 overflow-y-auto p-4 space-y-6 bg-gray-50/50">
        {messages.map((msg) => (
          <ChatMessage key={msg.id} message={msg} onDownload={handleDownload} />
        ))}
        {isTyping && (
          <div className="flex gap-3">
            <div className="w-8 h-8 bg-blue-600 rounded-full flex items-center justify-center text-white flex-shrink-0 animate-pulse">
              <Bot className="w-5 h-5" />
            </div>
            <div className="bg-white border border-gray-100 p-4 rounded-2xl rounded-tl-none shadow-sm text-sm text-gray-500">
              Agent is analyzing...
            </div>
          </div>
        )}
        <div ref={messagesEndRef} />
      </div>

      {/* Input */}
      <div className="p-4 bg-white border-t border-gray-100 space-y-3">
        {fileName && currentFile && (
          <FilePreview
            fileName={fileName}
            fileType={currentFile.type}
            fileData={currentFile.data}
            onRemove={() => { setFileName(null); setCurrentFile(null); setUploadedData(null); }}
          />
        )}
        <ChatInput
          value={inputValue}
          onChange={setInputValue}
          onSend={handleSendMessage}
          onFileClick={() => fileInputRef.current?.click()}
          onPaste={async (e) => {
            const item = e.clipboardData.items[0];
            if (item?.type.startsWith('image/')) {
              const file = item.getAsFile();
              if (file) await processFile(new File([file], `screenshot-${Date.now()}.png`, { type: file.type }));
            }
          }}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && !e.shiftKey) {
              e.preventDefault();
              handleSendMessage();
            }
          }}
          hasAttachment={!!currentFile}
          isTyping={isTyping}
        />
        <input
          type="file"
          ref={fileInputRef}
          onChange={handleFileUpload}
          accept=".csv, .xlsx, .xls, image/*"
          className="hidden"
        />
      </div>
    </div>
  );
}
