import React from 'react';
import ReactMarkdown from 'react-markdown';
import { Bot, User as UserIcon, FileSpreadsheet, Download } from 'lucide-react';
import { Message } from '../../types/production';

interface ChatMessageProps {
    message: Message;
    onDownload?: (fileName: string, buffer: any) => void;
}

const ChatMessage: React.FC<ChatMessageProps> = ({ message, onDownload }) => {
    const isAgent = message.role === 'agent';

    return (
        <div className={`flex gap-3 ${!isAgent ? 'flex-row-reverse' : ''}`}>
            <div className={`w-8 h-8 rounded-full flex items-center justify-center flex-shrink-0 ${isAgent ? 'bg-blue-600 text-white' : 'bg-gray-900 text-white'
                }`}>
                {isAgent ? <Bot className="w-5 h-5" /> : <UserIcon className="w-5 h-5" />}
            </div>

            <div className="max-w-[80%] space-y-2">
                <div className={`p-4 rounded-2xl shadow-sm ${isAgent
                        ? 'bg-white text-gray-800 rounded-tl-none border border-gray-100'
                        : 'bg-gray-900 text-white rounded-tr-none'
                    }`}>
                    <div className="prose prose-sm max-w-none prose-p:leading-relaxed prose-pre:bg-gray-800 prose-pre:text-gray-100">
                        <ReactMarkdown>{message.content}</ReactMarkdown>
                    </div>

                    {message.attachment && (
                        <div className="mt-3 space-y-2">
                            {message.attachment.type.startsWith('image/') ? (
                                <div className="rounded-lg overflow-hidden border border-gray-200 bg-white">
                                    <img
                                        src={message.attachment.data}
                                        alt={message.attachment.name}
                                        className="max-w-full h-auto max-h-64 object-contain mx-auto"
                                    />
                                </div>
                            ) : (
                                <div className={`flex items-center gap-3 p-3 rounded-xl border ${!isAgent ? 'bg-white/10 border-white/20 text-white' : 'bg-gray-50 border-gray-100 text-gray-700'
                                    }`}>
                                    <div className={`w-10 h-10 rounded-lg flex items-center justify-center ${!isAgent ? 'bg-white/20' : 'bg-white shadow-sm text-blue-600'
                                        }`}>
                                        <FileSpreadsheet className="w-6 h-6" />
                                    </div>
                                    <div className="flex-1 overflow-hidden">
                                        <p className="text-sm font-bold truncate">{message.attachment.name}</p>
                                        <p className="text-[10px] uppercase font-black tracking-wider opacity-60">
                                            {message.attachment.type.includes('csv') ? 'CSV FILE' : 'EXCEL DOCUMENT'}
                                        </p>
                                    </div>
                                </div>
                            )}
                        </div>
                    )}
                </div>

                {message.type === 'file' && message.fileData && onDownload && (
                    <button
                        onClick={() => onDownload(message.fileData!.name, message.fileData!.buffer)}
                        className="flex items-center gap-3 bg-green-50 border border-green-100 p-4 rounded-xl w-full hover:bg-green-100 transition-colors group text-left"
                    >
                        <div className="w-10 h-10 bg-green-100 group-hover:bg-green-200 rounded-lg flex items-center justify-center text-green-600 transition-colors">
                            <FileSpreadsheet className="w-6 h-6" />
                        </div>
                        <div className="flex-1">
                            <p className="font-medium text-green-900">{message.fileData.name}</p>
                            <p className="text-xs text-green-700">Click to download</p>
                        </div>
                        <Download className="w-5 h-5 text-green-600" />
                    </button>
                )}
            </div>
        </div>
    );
};

export default ChatMessage;
