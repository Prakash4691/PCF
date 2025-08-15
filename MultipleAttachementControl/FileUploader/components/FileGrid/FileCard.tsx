import * as React from "react";
import { Icon } from "@fluentui/react/lib/Icon";
import { Text } from "@fluentui/react/lib/Text";
import { TooltipHost } from "@fluentui/react/lib/Tooltip";
import { IconButton } from "@fluentui/react/lib/Button";
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
        <TooltipHost content={file.name}>
          <Text
            variant="medium"
            className="file-name"
            styles={{
              root: {
                fontWeight: "var(--fontWeightSemibold, 600)",
                color: "var(--colorNeutralForeground1, #323130)",
                fontSize: "var(--fontSizeBase300, 14px)",
                lineHeight: "var(--lineHeightBase300, 20px)",
              },
            }}
          >
            {file.name}
          </Text>
        </TooltipHost>
        <div className="file-meta">
          <Text
            variant="small"
            styles={{
              root: {
                color: "var(--colorNeutralForeground2, #616161)",
                fontSize: "var(--fontSizeBase200, 12px)",
              },
            }}
          >
            {fileType}
          </Text>
          <Text
            variant="small"
            className="file-size"
            styles={{
              root: {
                color: "var(--colorNeutralForeground3, #8a8886)",
                fontSize: "var(--fontSizeBase200, 12px)",
              },
            }}
          >
            {isExistingFile && (
              <TooltipHost content="This file exists as a notes record">
                <span className="existing-file-indicator">
                  <Icon
                    iconName="InfoSolid"
                    style={{
                      fontSize: "10px",
                      marginRight: "4px",
                      color: "var(--colorBrandForeground1, #0078d4)",
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
      <IconButton
        className="remove-button"
        iconProps={{ iconName: "Cancel" }}
        onClick={(e) => {
          e.stopPropagation();
          onRemove(index);
        }}
        title="Remove file"
        ariaLabel="Remove file"
        styles={{
          root: {
            backgroundColor: "transparent",
            border: "none",
            color: "var(--colorStatusDangerForeground1, #d13438)",
            width: "24px",
            height: "24px",
            minWidth: "24px",
            borderRadius: "var(--borderRadiusSmall, 2px)",
          },
          rootHovered: {
            backgroundColor: "var(--colorStatusDangerBackground1, #fdf3f4)",
            color: "var(--colorStatusDangerForeground1, #d13438)",
          },
          icon: {
            fontSize: "16px",
          },
        }}
      />
    </div>
  );
};
