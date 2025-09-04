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
  onPreview?: (index: number) => void;
  onDownload?: (index: number) => void;
  showActions?: boolean; // allow grid to control
}

export const FileCard: React.FC<FileCardProps> = ({
  fileInfo,
  index,
  onRemove,
  onPreview,
  onDownload,
  showActions = true,
}) => {
  const { file, icon, fileType, sizeText, isExistingFile, source, subject, noteText, createdOn } = fileInfo;
  
  // Utility function to strip HTML tags from text
  const stripHtmlTags = (htmlString?: string): string => {
    if (!htmlString) return '';
    return htmlString.replace(/<[^>]*>/g, '').trim();
  };
  
  // Create enhanced tooltip content based on source and available metadata
  const getTooltipContent = () => {
    const displayTitle = subject || file.name;
    const cleanDescription = stripHtmlTags(noteText);
    const description = cleanDescription ? `Description: ${cleanDescription}` : '';
    const sourceInfo = source === 'timeline' ? 'Source: Timeline Notes' : 'Source: File Upload Control';
    const dateInfo = createdOn ? `Created: ${createdOn.toLocaleDateString()}` : '';
    
    const tooltipParts = [displayTitle];
    if (description) tooltipParts.push(description);
    tooltipParts.push(sourceInfo);
    if (dateInfo) tooltipParts.push(dateInfo);
    
    return tooltipParts.join('\n');
  };

  return (
    <div className={`file-card ${isExistingFile ? "existing-file" : ""} ${source ? `source-${source}` : ""}`}>
      <div className="file-icon">
        <Icon iconName={icon} style={{ fontSize: "32px" }} />
        {/* Source indicator badge */}
        {source && (
          <div className={`source-badge source-${source}`}>
            <Icon 
              iconName={source === 'timeline' ? 'Timeline' : 'Upload'} 
              style={{ fontSize: '10px' }}
            />
          </div>
        )}
      </div>
      <div className="file-details">
        <TooltipHost content={getTooltipContent()}>
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
            {subject || file.name}
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
              <TooltipHost content={
                source === 'timeline' 
                  ? "This file is from timeline notes" 
                  : "This file is from file upload"
              }>
                <span className="existing-file-indicator">
                  <Icon
                    iconName={source === 'timeline' ? "Timeline" : "InfoSolid"}
                    style={{
                      fontSize: "10px",
                      marginRight: "4px",
                      color: source === 'timeline' 
                        ? "var(--colorPalettePurpleForeground1, #8764b8)" 
                        : "var(--colorBrandForeground1, #0078d4)",
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
      {showActions && (
        <div className="file-card-actions">
          <IconButton
            className="action-button"
            iconProps={{ iconName: "View" }}
            onClick={(e) => {
              e.stopPropagation();
              if (onPreview) {
                onPreview(index);
              }
            }}
            title="Preview file"
            ariaLabel="Preview file"
          />
          <IconButton
            className="action-button"
            iconProps={{ iconName: "Download" }}
            onClick={(e) => {
              e.stopPropagation();
              if (onDownload) {
                onDownload(index);
              }
            }}
            title="Download file"
            ariaLabel="Download file"
          />
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
      )}
    </div>
  );
};
