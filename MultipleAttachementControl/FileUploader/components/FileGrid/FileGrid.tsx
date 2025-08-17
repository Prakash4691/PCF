import * as React from "react";
import { FileCard } from "./FileCard";
import { getFileInfo } from "../../utils/fileHelpers";

interface FileGridProps {
  files: File[];
  onRemove: (index: number) => void;
  onPreview?: (index: number) => void;
  onDownload?: (index: number) => void;
}

export const FileGrid: React.FC<FileGridProps> = ({
  files,
  onRemove,
  onPreview,
  onDownload,
}) => {
  return (
    <div className="file-grid">
      {files.map((file, index) => {
        const fileInfo = getFileInfo(file);
        return (
          <FileCard
            key={`${file.name}-${index}`}
            fileInfo={fileInfo}
            index={index}
            onRemove={onRemove}
            onPreview={onPreview}
            onDownload={onDownload}
          />
        );
      })}
    </div>
  );
};
