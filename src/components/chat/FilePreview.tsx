import React from 'react';
import { X, FileSpreadsheet } from 'lucide-react';

interface FilePreviewProps {
    fileName: string;
    fileType: string;
    fileData?: string;
    onRemove: () => void;
}

const FilePreview: React.FC<FilePreviewProps> = ({ fileName, fileType, fileData, onRemove }) => {
    return (
        <div className="flex items-center justify-between bg-blue-50 px-3 py-2 rounded-lg border border-blue-100 animate-in fade-in slide-in-from-bottom-2 duration-200">
            <div className="flex items-center gap-2 text-sm text-blue-700 w-full overflow-hidden">
                {fileType.startsWith('image/') && fileData ? (
                    <div className="w-10 h-10 rounded border border-blue-200 overflow-hidden bg-white shadow-sm flex-shrink-0">
                        <img src={fileData} alt="preview" className="w-full h-full object-cover" />
                    </div>
                ) : (
                    <div className="w-10 h-10 rounded border border-blue-200 flex items-center justify-center bg-white shadow-sm text-blue-600 flex-shrink-0">
                        <FileSpreadsheet className="w-6 h-6" />
                    </div>
                )}
                <div className="flex flex-col flex-1 min-w-0">
                    <span className="font-bold truncate">{fileName}</span>
                    <span className="text-[10px] opacity-70 uppercase font-black tracking-wider">
                        {fileType.startsWith('image/') ? 'Image Attachment' : 'Data File (Ready to Analyze)'}
                    </span>
                </div>
            </div>
            <button
                onClick={onRemove}
                className="p-1 px-1.5 text-blue-400 hover:text-blue-600 hover:bg-blue-100 rounded-full transition-colors flex-shrink-0"
                title="Remove attachment"
            >
                <X className="w-4 h-4" />
            </button>
        </div>
    );
};

export default FilePreview;
