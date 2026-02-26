import React, { useState, useEffect, useRef } from "react";
import { saveAs } from "file-saver";
import {
  Bot,
  RefreshCw,
  Download,
  User as UserIcon,
  FileSpreadsheet,
  Loader2,
  Paperclip,
  Send,
  X,
  FileText,
  History,
  ChevronLeft,
} from "lucide-react";
import { GoogleGenAI, Type } from "@google/genai";
import ReactMarkdown from "react-markdown";

// Modular Imports
import {
  Message,
  ProjectData,
  ActualDataItem,
  FileAttachment,
} from "../types/production";
import { generateExcelFile } from "../utils/excelGenerator";
import { handleFileProcessing } from "../utils/fileHandlers";
import ChatHistorySidebar, {
  ChatSession,
  loadSessions,
  saveSessions,
  generateSessionTitle,
} from "./chat/ChatHistorySidebar";

const ai = new GoogleGenAI({
  apiKey: import.meta.env.VITE_GEMINI_API_KEY || "",
});

const DEFAULT_MESSAGE: Message = {
  id: "1",
  role: "agent",
  content:
    "Hello! I'm your Production Plan Agent. I can help you create a detailed Excel production plan. \n\nTo get started, please tell me about your project: **What is the project name, your total goal, the start/end dates, and who is working on it?**",
};

export default function ProductionPlanMaker() {
  const [sessions, setSessions] = useState<ChatSession[]>([]);
  const [activeSessionId, setActiveSessionId] = useState<string>("");
  const [showSidebar, setShowSidebar] = useState(false);
  const [messages, setMessages] = useState<Message[]>([DEFAULT_MESSAGE]);
  const [inputValue, setInputValue] = useState("");
  const [isTyping, setIsTyping] = useState(false);
  const [isStreaming, setIsStreaming] = useState(false);
  const [uploadedData, setUploadedData] = useState<ActualDataItem[] | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [currentFile, setCurrentFile] = useState<FileAttachment | null>(null);
  const [currentProject, setCurrentProject] = useState<Partial<ProjectData> | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [previewImage, setPreviewImage] = useState<{ url: string; name: string } | null>(null);

  const messagesEndRef = useRef<HTMLDivElement>(null);
  const chatRef = useRef<any>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const streamIntervalRef = useRef<ReturnType<typeof setInterval> | null>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const deletedSessionsRef = useRef<Set<string>>(new Set());

  const initChat = () => {
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
        tools: [
          {
            functionDeclarations: [
              {
                name: "generate_production_plan",
                description:
                  "Generates the production planning Excel file once core project details and column structure are confirmed.",
                parameters: {
                  type: Type.OBJECT,
                  properties: {
                    name: { type: Type.STRING, description: "The name of the project" },
                    goal: { type: Type.NUMBER, description: "The total numeric goal" },
                    unit: { type: Type.STRING, description: "The unit of measurement (e.g., 'units', 'hours')" },
                    startDate: { type: Type.STRING, description: "Start date in YYYY-MM-DD format" },
                    endDate: { type: Type.STRING, description: "End date in YYYY-MM-DD format" },
                    resources: { type: Type.ARRAY, items: { type: Type.STRING }, description: "List of names of teams or individuals" },
                    columns: {
                      type: Type.ARRAY,
                      items: {
                        type: Type.OBJECT,
                        properties: {
                          header: { type: Type.STRING, description: "The display name of the column" },
                          key: { type: Type.STRING, description: "A unique key for the column" },
                          section: { type: Type.STRING, enum: ["Target", "Actual", "Accumulative"], description: "Which section the column belongs to" },
                          formula: { type: Type.STRING, description: "Excel formula referencing DailyProductionTable. Use {rowIndex} for the current row." },
                        },
                        required: ["header", "key", "section", "formula"],
                      },
                    },
                    dailyColumns: {
                      type: Type.ARRAY,
                      items: {
                        type: Type.OBJECT,
                        properties: {
                          header: { type: Type.STRING },
                          key: { type: Type.STRING },
                          formula: { type: Type.STRING },
                        },
                        required: ["header", "key"],
                      },
                    },
                    pivotColumns: {
                      type: Type.ARRAY,
                      items: {
                        type: Type.OBJECT,
                        properties: {
                          header: { type: Type.STRING },
                          formula: { type: Type.STRING },
                        },
                        required: ["header", "formula"],
                      },
                    },
                    dashboardMetrics: {
                      type: Type.ARRAY,
                      items: {
                        type: Type.OBJECT,
                        properties: {
                          label: { type: Type.STRING },
                          formula: { type: Type.STRING },
                          format: { type: Type.STRING },
                        },
                        required: ["label", "formula"],
                      },
                    },
                    actualData: {
                      type: Type.ARRAY,
                      items: {
                        type: Type.OBJECT,
                        properties: {
                          date: { type: Type.STRING },
                          name: { type: Type.STRING },
                          actual: { type: Type.NUMBER },
                        },
                        required: ["date", "name", "actual"],
                      },
                    },
                  },
                  required: ["name", "goal", "unit", "startDate", "endDate", "resources", "columns"],
                },
              },
            ],
          },
        ],
      },
    });
  };

  // Initialize Gemini Chat
  useEffect(() => {
    if (!chatRef.current) initChat();
  }, []);

  // Load sessions on mount
  useEffect(() => {
    const stored = loadSessions();
    if (stored.length > 0) {
      setSessions(stored);
      const last = stored[stored.length - 1];
      setActiveSessionId(last.id);
      setMessages(last.messages.length > 0 ? last.messages : [DEFAULT_MESSAGE]);
    } else {
      setActiveSessionId(Date.now().toString());
    }
  }, []);

  // Save sessions whenever messages change
  useEffect(() => {
    if (!activeSessionId) return;
    setSessions((prev) => {
      const filtered = prev.filter(s => !deletedSessionsRef.current.has(s.id));
      const exists = filtered.find((s) => s.id === activeSessionId);
      let updated: ChatSession[];
      if (exists) {
        updated = filtered.map((s) =>
          s.id === activeSessionId ? { ...s, messages, title: generateSessionTitle(messages) } : s,
        );
      } else {
        updated = [...filtered, { id: activeSessionId, title: generateSessionTitle(messages), createdAt: new Date().toISOString(), messages }];
      }
      saveSessions(updated);
      return updated;
    });
  }, [messages, activeSessionId]);

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages, isTyping]);

  useEffect(() => {
    return () => {
      if (streamIntervalRef.current) clearInterval(streamIntervalRef.current);
    };
  }, []);

  useEffect(() => {
    if (textareaRef.current) {
      textareaRef.current.style.height = "auto";
      textareaRef.current.style.height = `${Math.min(textareaRef.current.scrollHeight, 200)}px`;
    }
  }, [inputValue]);

  const typewriterEffect = (fullText: string, msgId: string) => {
    let i = 0;
    setIsStreaming(true);
    if (streamIntervalRef.current) clearInterval(streamIntervalRef.current);
    streamIntervalRef.current = setInterval(() => {
      i += 5;
      setMessages((prev) =>
        prev.map((m) => (m.id === msgId ? { ...m, content: fullText.slice(0, i) } : m)),
      );
      if (i >= fullText.length) {
        if (streamIntervalRef.current) clearInterval(streamIntervalRef.current);
        streamIntervalRef.current = null;
        setIsStreaming(false);
      }
    }, 5);
  };

  const processFile = async (file: File) => {
    try {
      const processed = await handleFileProcessing(file);
      setFileName(processed.name!);
      if (processed.parsedData) setUploadedData(processed.parsedData);
      setCurrentFile(processed as FileAttachment);
    } catch (error) {
      alert(error instanceof Error ? error.message : "Error processing file.");
    }
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    await processFile(file);
    e.target.value = "";
  };

  const handleSendMessage = async () => {
    if ((!inputValue.trim() && !currentFile) || isTyping || isStreaming) return;

    const contextPreamble = currentProject
      ? `[CURRENT PROJECT STATE: Name="${currentProject.name || "?"}", Goal=${currentProject.goal || "?"}, Unit="${currentProject.unit || "?"}", Dates=${currentProject.startDate || "?"}/${currentProject.endDate || "?"}, Resources=${currentProject.resources?.join(",") || "?"}]\n`
      : "";

    const fileMetadata = currentFile?.metadata || "";
    const fullPrompt = `${contextPreamble}${inputValue}${fileMetadata}`;

    const userMsg: Message = {
      id: Date.now().toString(),
      role: "user",
      content:
        inputValue ||
        (currentFile
          ? `Shared ${currentFile.type.startsWith("image/") ? "an image" : "a file"}: ${currentFile.name}`
          : ""),
      attachment: currentFile
        ? { name: currentFile.name, type: currentFile.type, data: currentFile.data }
        : undefined,
    };

    setMessages((prev) => [...prev, userMsg]);
    setInputValue("");
    setCurrentFile(null);
    setFileName(null);
    setIsTyping(true);

    try {
      if (!chatRef.current) initChat();

      let result;
      if (userMsg.attachment && userMsg.attachment.type.startsWith("image/")) {
        const base64Data = userMsg.attachment.data.split(",")[1];
        result = await chatRef.current!.sendMessage({
          message: [
            { text: fullPrompt },
            { inlineData: { data: base64Data, mimeType: userMsg.attachment.type } },
          ],
        });
      } else {
        result = await chatRef.current!.sendMessage({ message: fullPrompt });
      }

      const functionCalls = result?.functionCalls;
      if (functionCalls && functionCalls.length > 0) {
        for (const call of functionCalls) {
          if (call.name === "generate_production_plan") {
            const projectData = call.args as ProjectData;
            setCurrentProject(projectData);

            const combinedActualData = [...(projectData.actualData || [])];
            if (uploadedData) {
              uploadedData.forEach((upItem) => {
                const exists = combinedActualData.some(
                  (c) => c.date === upItem.date && c.name === upItem.name,
                );
                if (!exists) combinedActualData.push(upItem);
              });
            }
            projectData.actualData = combinedActualData.length > 0 ? combinedActualData : undefined;
            const buffer = await generateExcelFile(projectData);

            const msgId = Date.now().toString();
            const successText = `I've generated the production plan for **${projectData.name}**. You can download it below.`;
            setMessages((prev) => [
              ...prev,
              {
                id: msgId,
                role: "agent",
                content: "",
                type: "file",
                fileData: {
                  name: `${projectData.name.replace(/\s+/g, "_")}_Production_Planning.xlsx`,
                  buffer: buffer,
                },
              },
            ]);
            typewriterEffect(successText, msgId);
          }
        }
      } else {
        const textResponse =
          result?.text ||
          "I'm sorry, I didn't quite get that. Could you please provide more details about your project?";
        const msgId = Date.now().toString();
        setMessages((prev) => [...prev, { id: msgId, role: "agent", content: "" }]);
        typewriterEffect(textResponse, msgId);
      }
    } catch (error) {
      console.error("Gemini Error:", error);
      const msgId = Date.now().toString();
      setMessages((prev) => [...prev, { id: msgId, role: "agent", content: "" }]);
      typewriterEffect("I'm having a bit of trouble connecting to my brain. Could you try again?", msgId);
    } finally {
      setIsTyping(false);
    }
  };

  const handleDownload = (fileName: string, buffer: any) => {
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    saveAs(blob, fileName);
  };

  const downloadAttachment = (dataUrl: string, fileName: string) => {
    const link = document.createElement("a");
    link.href = dataUrl;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSendMessage();
    }
  };

  const resetChat = () => {
    if (streamIntervalRef.current) clearInterval(streamIntervalRef.current);
    setIsStreaming(false);
    setMessages([DEFAULT_MESSAGE]);
    setUploadedData(null);
    setFileName(null);
    setCurrentProject(null);
    initChat();
  };

  const startNewSession = () => {
    if (streamIntervalRef.current) clearInterval(streamIntervalRef.current);
    setIsStreaming(false);
    const newId = Date.now().toString();
    setSessions((prev) => {
      const updated = [...prev, { id: newId, title: "New Chat", createdAt: new Date().toISOString(), messages: [DEFAULT_MESSAGE] }];
      saveSessions(updated);
      return updated;
    });
    setActiveSessionId(newId);
    setMessages([DEFAULT_MESSAGE]);
    setUploadedData(null);
    setFileName(null);
    setCurrentFile(null);
    setCurrentProject(null);
    chatRef.current = null;
    setShowSidebar(false);
    initChat();
  };

  const loadSession = (session: ChatSession) => {
    if (streamIntervalRef.current) clearInterval(streamIntervalRef.current);
    setIsStreaming(false);
    setActiveSessionId(session.id);
    setMessages(session.messages.length > 0 ? session.messages : [DEFAULT_MESSAGE]);
    setUploadedData(null);
    setFileName(null);
    setCurrentFile(null);
    setCurrentProject(null);
    chatRef.current = null;
    setShowSidebar(false);
    initChat();
  };

  const handleDeleteSession = (sessionId: string) => {
    deletedSessionsRef.current.add(sessionId);
    setSessions((prev) => {
      const updated = prev.filter((s) => s.id !== sessionId);
      saveSessions(updated);
      return updated;
    });
    if (sessionId === activeSessionId) {
      const remaining = sessions.filter((s) => s.id !== sessionId);
      if (remaining.length > 0) {
        loadSession(remaining[remaining.length - 1]);
      } else {
        startNewSession();
      }
    }
  };

  return (
    <div
      className="max-w-4xl mx-auto h-[700px] flex rounded-2xl shadow-xl overflow-hidden relative"
      style={{ border: "1px solid #e5e0d5" }}
      onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
      onDragLeave={(e) => {
        const relatedTarget = e.relatedTarget as Node | null;
        if (!e.currentTarget.contains(relatedTarget)) setIsDragging(false);
      }}
      onDrop={async (e) => {
        e.preventDefault();
        setIsDragging(false);
        const file = e.dataTransfer.files?.[0];
        if (file) await processFile(file);
      }}
    >
      {/* ── Sidebar ── */}
      <ChatHistorySidebar
        sessions={sessions}
        activeSessionId={activeSessionId}
        showSidebar={showSidebar}
        onNewSession={startNewSession}
        onLoadSession={loadSession}
        onDeleteSession={handleDeleteSession}
      />

      {/* ── Main Chat ── */}
      <div className="flex-1 flex flex-col min-w-0 relative">

        {/* Drag Overlay */}
        {isDragging && (
          <div className="absolute inset-0 z-50 bg-blue-600/10 backdrop-blur-sm flex items-center justify-center pointer-events-none">
            <div className="bg-white p-8 rounded-3xl shadow-2xl border-2 border-blue-500 border-dashed flex flex-col items-center gap-4 animate-in zoom-in-95 duration-200">
              <div className="w-16 h-16 bg-blue-100 rounded-full flex items-center justify-center text-blue-600">
                <Download className="w-8 h-8 animate-bounce" />
              </div>
              <div className="text-center">
                <p className="text-xl font-bold text-gray-900">Drop files here</p>
                <p className="text-sm text-gray-500 mt-1">CSV, Excel, PDF, Docs, or Images</p>
              </div>
            </div>
          </div>
        )}

        {/* Header */}
        <div
          className="p-4 flex justify-between items-center flex-shrink-0"
          style={{ backgroundColor: "#133020", borderBottom: "1px solid #046241" }}
        >
          <div className="flex items-center gap-3">
            <button
              onClick={() => setShowSidebar((v) => !v)}
              className="p-1.5 rounded-lg transition-opacity hover:opacity-70"
              style={{ color: "#FFC370" }}
              title="Toggle History"
            >
              {showSidebar ? <ChevronLeft className="w-5 h-5" /> : <History className="w-5 h-5" />}
            </button>
            <div
              className="w-10 h-10 rounded-full flex items-center justify-center text-white"
              style={{ backgroundColor: "#046241" }}
            >
              <Bot className="w-6 h-6" />
            </div>
            <div>
              <h1 className="font-bold text-white">Production Plan Agent</h1>
              <p className="text-xs flex items-center gap-1" style={{ color: "#FFB347" }}>
                <span className="w-2 h-2 rounded-full animate-pulse inline-block" style={{ backgroundColor: "#FFB347" }}></span>
                Powered by Gemini AI
              </p>
            </div>
          </div>
          <button
            onClick={startNewSession}
            className="p-2 rounded-full transition-opacity hover:opacity-70"
            style={{ color: "#FFC370" }}
            title="New Chat"
          >
            <RefreshCw className="w-5 h-5" />
          </button>
        </div>

        {/* Messages */}
        <div className="flex-1 overflow-y-auto p-4 space-y-6" style={{ backgroundColor: "#f5eedb" }}>
          {messages.map((msg) => (
            <div key={msg.id} className={`flex gap-3 ${msg.role === "user" ? "flex-row-reverse" : ""}`}>
              <div
                className="w-8 h-8 rounded-full flex items-center justify-center flex-shrink-0 text-white"
                style={{ backgroundColor: msg.role === "agent" ? "#046241" : "#133020" }}
              >
                {msg.role === "agent" ? <Bot className="w-5 h-5" /> : <UserIcon className="w-5 h-5" />}
              </div>

              <div className="max-w-[80%] space-y-2">
                <div
                  className="p-4 shadow-sm"
                  style={
                    msg.role === "agent"
                      ? { backgroundColor: "#ffffff", color: "#133020", borderRadius: "0 1rem 1rem 1rem", border: "1px solid #e5e0d5" }
                      : { backgroundColor: "#133020", color: "#ffffff", borderRadius: "1rem 0 1rem 1rem" }
                  }
                >
                  <div className="leading-relaxed prose prose-sm max-w-none">
                    <ReactMarkdown
                      components={{
                        p: ({ children }: any) => <p className="mb-2 last:mb-0">{children}</p>,
                        strong: ({ children }: any) => <strong className="font-semibold">{children}</strong>,
                        ul: ({ children }: any) => <ul className="list-disc list-inside mb-2 space-y-1">{children}</ul>,
                        ol: ({ children }: any) => <ol className="list-decimal list-inside mb-2 space-y-1">{children}</ol>,
                        li: ({ children }: any) => <li className="text-sm">{children}</li>,
                        code: ({ children }: any) => (
                          <code className="px-1 rounded text-xs font-mono" style={{ backgroundColor: "#F9F7F7", color: "#133020" }}>
                            {children}
                          </code>
                        ),
                      }}
                    >
                      {msg.content}
                    </ReactMarkdown>
                  </div>

                  {/* Attachment Preview */}
                  {msg.attachment && (
                    <div className="mt-3 pt-3 border-t border-white/10">
                      {msg.attachment.type.startsWith("image/") ? (
                        <div
                          className="relative group cursor-pointer overflow-hidden rounded-lg border border-white/20 w-fit"
                          onClick={() => setPreviewImage({ url: msg.attachment!.data, name: msg.attachment!.name })}
                        >
                          <img
                            src={msg.attachment.data}
                            alt={msg.attachment.name}
                            className="block max-w-full h-auto max-h-[300px] transition-transform group-hover:scale-105"
                          />
                          <div className="absolute inset-0 bg-black/0 group-hover:bg-black/20 transition-colors flex items-center justify-center opacity-0 group-hover:opacity-100">
                            <div className="bg-black/50 p-2 rounded-full text-white">
                              <Download className="w-5 h-5" />
                            </div>
                          </div>
                        </div>
                      ) : (
                        <div
                          className="flex items-center gap-3 p-3 rounded-lg bg-white/10 border border-white/20 cursor-pointer hover:bg-white/20 transition-colors"
                          onClick={() => downloadAttachment(msg.attachment!.data, msg.attachment!.name)}
                        >
                          <div className="p-2 bg-white/20 rounded-md">
                            <FileText className="w-5 h-5 text-white" />
                          </div>
                          <div className="flex-1 min-w-0">
                            <p className="text-sm font-medium text-white truncate">{msg.attachment.name}</p>
                            <p className="text-xs text-white/70">Click to download</p>
                          </div>
                          <Download className="w-4 h-4 text-white/70" />
                        </div>
                      )}
                    </div>
                  )}
                </div>

                {/* Download Section */}
                {msg.type === "file" && msg.fileData && !isStreaming && (
                  <button
                    onClick={() => handleDownload(msg.fileData!.name, msg.fileData!.buffer)}
                    className="flex items-center gap-3 p-4 rounded-xl w-full transition-opacity text-left hover:opacity-90"
                    style={{ backgroundColor: "#FFC370", border: "1px solid #FFB347" }}
                  >
                    <div className="w-10 h-10 rounded-lg flex items-center justify-center" style={{ backgroundColor: "rgba(255,255,255,0.3)" }}>
                      <FileSpreadsheet className="w-6 h-6" style={{ color: "#133020" }} />
                    </div>
                    <div className="flex-1">
                      <p className="font-medium" style={{ color: "#133020" }}>{msg.fileData.name}</p>
                      <p className="text-xs" style={{ color: "#046241" }}>Click to download</p>
                    </div>
                    <Download className="w-5 h-5" style={{ color: "#133020" }} />
                  </button>
                )}
              </div>
            </div>
          ))}

          {/* Typing indicator */}
          {isTyping && (
            <div className="flex gap-3">
              <div className="w-8 h-8 rounded-full flex items-center justify-center text-white flex-shrink-0" style={{ backgroundColor: "#046241" }}>
                <Bot className="w-5 h-5" />
              </div>
              <div className="p-4 rounded-2xl shadow-sm flex items-center gap-2" style={{ backgroundColor: "#ffffff", border: "1px solid #e5e0d5" }}>
                <Loader2 className="w-4 h-4 animate-spin" style={{ color: "#046241" }} />
                <span className="text-sm" style={{ color: "#133020" }}>Agent is analyzing...</span>
              </div>
            </div>
          )}
          <div ref={messagesEndRef} />
        </div>

        {/* Input Area */}
        <div className="p-4 space-y-3 flex-shrink-0" style={{ backgroundColor: "#ffffff", borderTop: "1px solid #e5e0d5" }}>
          {fileName && (
            <div className="flex items-center justify-between px-3 py-2 rounded-lg" style={{ backgroundColor: "#f5eedb", border: "1px solid #FFC370" }}>
              <div className="flex items-center gap-2 text-sm" style={{ color: "#046241" }}>
                <Paperclip className="w-4 h-4" />
                <span className="font-medium truncate max-w-[200px]">{fileName}</span>
              </div>
              <button
                onClick={() => { setFileName(null); setUploadedData(null); setCurrentFile(null); }}
                className="hover:opacity-70"
                style={{ color: "#FFB347" }}
              >
                <X className="w-4 h-4" />
              </button>
            </div>
          )}

          <div className="flex items-end gap-2">
            <button
              onClick={() => fileInputRef.current?.click()}
              className="p-3 mb-0.5 rounded-xl transition-opacity hover:opacity-70"
              style={{ color: "#046241" }}
              title="Upload file (CSV, Excel, PDF, Doc, PPT, Image)"
            >
              <Paperclip className="w-5 h-5" />
            </button>
            <input
              type="file"
              ref={fileInputRef}
              onChange={handleFileUpload}
              accept=".csv, .xlsx, .xls, .pdf, .docx, .doc, .pptx, .ppt, .txt, .md, .json, image/*"
              className="hidden"
            />
            <textarea
              ref={textareaRef}
              rows={1}
              value={inputValue}
              onChange={(e) => setInputValue(e.target.value)}
              onKeyDown={handleKeyDown}
              placeholder="Describe your project..."
              disabled={isTyping || isStreaming}
              className="flex-1 px-4 py-3 rounded-xl outline-none transition-all disabled:opacity-50 disabled:cursor-not-allowed resize-none overflow-y-auto max-h-[200px]"
              style={{ backgroundColor: "#F9F7F7", border: "1px solid #e5e0d5", color: "#133020" }}
            />
            <button
              onClick={handleSendMessage}
              disabled={(!inputValue.trim() && !currentFile) || isTyping || isStreaming}
              className="p-3 mb-0.5 rounded-xl transition-opacity shadow-sm disabled:opacity-50 disabled:cursor-not-allowed text-white"
              style={{ backgroundColor: "#046241" }}
            >
              <Send className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* Image Preview Modal */}
        {previewImage && (
          <div className="absolute inset-0 z-50 bg-black/80 backdrop-blur-sm flex items-center justify-center p-4 animate-in fade-in duration-200">
            <div className="relative max-w-full max-h-full flex flex-col items-center">
              <div className="absolute top-4 right-4 flex gap-3 z-10">
                <button
                  onClick={() => downloadAttachment(previewImage.url, previewImage.name)}
                  className="p-3 bg-black/60 hover:bg-black/80 rounded-full text-white backdrop-blur-md transition-all border border-white/20 shadow-lg hover:scale-105"
                  title="Download"
                >
                  <Download className="w-6 h-6" />
                </button>
                <button
                  onClick={() => setPreviewImage(null)}
                  className="p-3 bg-black/60 hover:bg-black/80 rounded-full text-white backdrop-blur-md transition-all border border-white/20 shadow-lg hover:scale-105"
                  title="Close"
                >
                  <X className="w-6 h-6" />
                </button>
              </div>
              <img
                src={previewImage.url}
                alt={previewImage.name}
                className="max-w-full max-h-[80vh] rounded-lg shadow-2xl object-contain"
              />
              <p className="mt-4 text-white/80 font-medium">{previewImage.name}</p>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}