import * as React from "react";
import { Icon } from "@fluentui/react/lib/Icon";
import { Text } from "@fluentui/react/lib/Text";
import { TooltipHost } from "@fluentui/react/lib/Tooltip";
import { FileInfo } from "../../types/interfaces";

interface FileCardProps {
  fileInfo: FileInfo;
  index: number;
  onRemove: (index: number) => void;
}

export const FileCard: React.FC<FileCardProps> = ({
  fileInfo,
  index,
  onRemove,
}) => {
  const { file, icon, fileType, sizeText, isExistingFile } = fileInfo;

  return (
    <div className={`file-card ${isExistingFile ? "existing-file" : ""}`}>
      <div className="file-icon">
        <Icon iconName={icon} style={{ fontSize: "32px" }} />
      </div>
      <div className="file-details">
        <div className="file-name" title={file.name}>
          {file.name}
        </div>
        <div className="file-meta">
          <Text variant="small">{fileType}</Text>
          <Text variant="small" className="file-size">
            {isExistingFile && (
              <TooltipHost content="This file exists as a notes record">
                <span className="existing-file-indicator">
                  <Icon
                    iconName="InfoSolid"
                    style={{
                      fontSize: "10px",
                      marginRight: "4px",
                    }}
                  />
                  {sizeText}
                </span>
              </TooltipHost>
            )}
            {!isExistingFile && sizeText}
          </Text>
        </div>
      </div>
      <button
        className="remove-button"
        onClick={(e) => {
          e.stopPropagation();
          onRemove(index);
        }}
        title="Remove file"
      >
        <Icon iconName="Cancel" />
      </button>
    </div>
  );
};
