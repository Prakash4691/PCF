import * as React from "react";
import { Icon } from "@fluentui/react/lib/Icon";
import { FileCard } from "./FileCard";
import { getFileInfo } from "../../utils/fileHelpers";

interface FileGridProps {
  files: File[];
  hasOverflow: boolean;
  onRemove: (index: number) => void;
}

export const FileGrid: React.FC<FileGridProps> = ({
  files,
  hasOverflow,
  onRemove,
}) => {
  const fileGridRef = React.useRef<HTMLDivElement>(null);

  return (
    <>
      {hasOverflow && (
        <div className="scroll-indicator">
          <Icon iconName="DoubleChevronLeft8" /> Scroll to see more files{" "}
          <Icon iconName="DoubleChevronRight8" />
        </div>
      )}
      <div className="file-grid" ref={fileGridRef}>
        {files.map((file, index) => {
          const fileInfo = getFileInfo(file);
          return (
            <FileCard
              key={`${file.name}-${index}`}
              fileInfo={fileInfo}
              index={index}
              onRemove={onRemove}
            />
          );
        })}
      </div>
    </>
  );
};
