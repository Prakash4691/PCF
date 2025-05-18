import * as React from "react";
import { Icon } from "@fluentui/react/lib/Icon";
import { Label } from "@fluentui/react/lib/Label";
import { Separator } from "@fluentui/react/lib/Separator";

interface FileHeaderProps {
  fileCount: number;
  onAddFiles: () => void;
}

export const FileHeader: React.FC<FileHeaderProps> = ({
  fileCount,
  onAddFiles,
}) => {
  return (
    <>
      <div className="files-header">
        <Label>Selected Files ({fileCount})</Label>
        <button className="add-files-button" onClick={onAddFiles}>
          <Icon iconName="Add" />
          <span>Add Files</span>
        </button>
      </div>
      <Separator />
    </>
  );
};
