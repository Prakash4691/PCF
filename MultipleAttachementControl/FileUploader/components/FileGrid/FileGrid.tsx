import * as React from "react";
import { FileCard } from "./FileCard";
import { getFileInfo, getSafeFileName } from "../../utils/fileHelpers";
import { FileWithContent } from "../../types/interfaces";

interface FileGridProps {
  files: File[];
  fileMetadata?: Map<string, FileWithContent>; // Map file name to metadata
  onRemove: (index: number) => void;
  onPreview?: (index: number) => void;
  onDownload?: (index: number) => void;
}

export const FileGrid: React.FC<FileGridProps> = ({
  files,
  fileMetadata,
  onRemove,
  onPreview,
  onDownload,
}) => {
  return (
    <div className="file-grid">
      {files.map((file, index) => {
        const safeName = getSafeFileName(file);
        // Get metadata for this file if available
        const metadata = fileMetadata?.get(safeName);
        const additionalInfo = metadata
          ? {
              source: metadata.source,
              subject: metadata.subject,
              noteText: metadata.noteText,
              createdOn: metadata.createdOn,
              modifiedOn: metadata.modifiedOn,
            }
          : undefined;

        const fileInfo = getFileInfo(file, additionalInfo);
        const key = (metadata && metadata.guid) || safeName || `${index}`;

        return (
          <FileCard
            key={key}
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
