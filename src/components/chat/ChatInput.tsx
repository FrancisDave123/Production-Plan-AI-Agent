import React, { useRef, useEffect } from 'react';
import { Paperclip, Send } from 'lucide-react';

interface ChatInputProps {
    value: string;
    onChange: (value: string) => void;
    onSend: () => void;
    onFileClick: () => void;
    onPaste: (e: React.ClipboardEvent) => void;
    onKeyDown: (e: React.KeyboardEvent<HTMLTextAreaElement>) => void;
    disabled?: boolean;
    hasAttachment?: boolean;
    isTyping?: boolean;
}

const ChatInput: React.FC<ChatInputProps> = ({
    value,
    onChange,
    onSend,
    onFileClick,
    onPaste,
    onKeyDown,
    disabled,
    hasAttachment,
    isTyping
}) => {
    const textareaRef = useRef<HTMLTextAreaElement>(null);

    useEffect(() => {
        if (textareaRef.current) {
            textareaRef.current.style.height = 'auto';
            textareaRef.current.style.height = `${Math.min(textareaRef.current.scrollHeight, 200)}px`;
        }
    }, [value]);

    const canSend = (value.trim() || hasAttachment) && !isTyping;

    return (
        <div className="flex items-end gap-2">
            <button
                onClick={onFileClick}
                className="p-3 mb-0.5 text-gray-400 hover:text-blue-600 hover:bg-blue-50 rounded-xl transition-all"
                title="Upload data (CSV, Excel) or Images"
                disabled={isTyping}
            >
                <Paperclip className="w-5 h-5" />
            </button>

            <textarea
                ref={textareaRef}
                rows={1}
                value={value}
                onChange={(e) => onChange(e.target.value)}
                onKeyDown={onKeyDown}
                onPaste={onPaste}
                placeholder="Describe your project..."
                disabled={disabled || isTyping}
                className="flex-1 px-4 py-3 bg-gray-50 border border-gray-200 rounded-xl focus:ring-2 focus:ring-blue-500 focus:bg-white outline-none transition-all disabled:opacity-50 disabled:cursor-not-allowed resize-none overflow-y-auto max-h-[200px]"
            />

            <button
                onClick={onSend}
                disabled={!canSend}
                className={`p-3 mb-0.5 rounded-xl transition-all shadow-sm ${canSend
                        ? 'bg-blue-600 text-white hover:bg-blue-700'
                        : 'bg-gray-100 text-gray-300 cursor-not-allowed'
                    }`}
            >
                <Send className="w-5 h-5" />
            </button>
        </div>
    );
};

export default ChatInput;
