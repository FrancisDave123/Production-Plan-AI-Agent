import React from 'react';
import { Plus, Trash2 } from 'lucide-react';

export interface Message {
    id: string;
    role: 'agent' | 'user';
    content: string;
    type?: 'text' | 'file';
    fileData?: {
        name: string;
        buffer: any;
    };
}

export interface ChatSession {
    id: string;
    title: string;
    createdAt: string;
    messages: Message[];
}

export const STORAGE_KEY = 'productionPlanChatHistory';

export const loadSessions = (): ChatSession[] => {
    try {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (!raw) return [];
        const sessions: ChatSession[] = JSON.parse(raw);
        return sessions.map(s => ({
            ...s,
            messages: s.messages.map(m => ({ ...m, fileData: undefined }))
        }));
    } catch { return []; }
};

export const saveSessions = (sessions: ChatSession[]) => {
    try {
        const stripped = sessions.map(s => ({
            ...s,
            messages: s.messages.map(m => ({ ...m, fileData: undefined }))
        }));
        localStorage.setItem(STORAGE_KEY, JSON.stringify(stripped));
    } catch { }
};

export const generateSessionTitle = (messages: Message[]): string => {
    const firstUser = messages.find(m => m.role === 'user');
    if (firstUser) return firstUser.content.length > 40 ? firstUser.content.slice(0, 40) + '…' : firstUser.content;
    return 'New Chat';
};

interface ChatHistorySidebarProps {
    sessions: ChatSession[];
    activeSessionId: string;
    showSidebar: boolean;
    onNewSession: () => void;
    onLoadSession: (session: ChatSession) => void;
    onDeleteSession: (sessionId: string) => void;
}

export default function ChatHistorySidebar({
    sessions,
    activeSessionId,
    showSidebar,
    onNewSession,
    onLoadSession,
    onDeleteSession,
}: ChatHistorySidebarProps) {
    return (
        <div
            className="flex flex-col transition-all duration-300 overflow-hidden flex-shrink-0"
            style={{
                width: showSidebar ? '260px' : '0px',
                backgroundColor: '#1a3d28',
                borderRight: showSidebar ? '1px solid #046241' : 'none'
            }}
        >
            {showSidebar && (
                <>
                    {/* Sidebar Header */}
                    <div className="p-4 flex items-center justify-between flex-shrink-0" style={{ borderBottom: '1px solid #046241' }}>
                        <span className="font-semibold text-white text-sm">Chat History</span>
                        <button
                            onClick={onNewSession}
                            className="p-1.5 rounded-lg hover:opacity-70 transition-opacity"
                            style={{ backgroundColor: '#046241', color: '#FFC370' }}
                            title="New Chat"
                        >
                            <Plus className="w-4 h-4" />
                        </button>
                    </div>

                    {/* Session List */}
                    <div className="flex-1 overflow-y-auto p-2 space-y-1">
                        {sessions.length === 0 && (
                            <p className="text-xs text-center p-4" style={{ color: '#FFC370', opacity: 0.6 }}>No history yet</p>
                        )}
                        {[...sessions].reverse().map(session => (
                            <div
                                key={session.id}
                                className="flex items-center gap-2 p-3 rounded-xl cursor-pointer group transition-opacity hover:opacity-90"
                                style={{ backgroundColor: session.id === activeSessionId ? '#046241' : 'rgba(255,255,255,0.05)' }}
                            >
                                <div
                                    className="flex-1 min-w-0"
                                    onClick={() => onLoadSession(session)}
                                >
                                    <p className="text-xs font-medium truncate text-white">{session.title}</p>
                                    <p className="text-xs mt-0.5" style={{ color: '#FFB347', opacity: 0.8 }}>
                                        {new Date(session.createdAt).toLocaleDateString()}
                                    </p>
                                </div>
                                <button
                                    type="button"
                                    onClick={(e) => {
                                        e.stopPropagation();
                                        console.log('delete clicked', session.id);
                                        onDeleteSession(session.id);
                                    }}
                                    className="opacity-0 group-hover:opacity-100 transition-opacity p-1 rounded"
                                    style={{ color: '#FFC370' }}
                                >
                                    <Trash2 className="w-3 h-3" />
                                </button>
                            </div>
                        ))}
                    </div>
                </>
            )}
        </div>
    );
}