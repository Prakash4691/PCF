import * as React from "react";
import { Icon } from "@fluentui/react/lib/Icon";
import { Text } from "@fluentui/react/lib/Text";
import { ProgressIndicator } from "@fluentui/react/lib/ProgressIndicator";
import { IconButton } from "@fluentui/react/lib/Button";
import { Stack } from "@fluentui/react/lib/Stack";
import { FileUploadProgress, FileWithContent } from "../types/interfaces";

interface UploadProgressProps {
  uploadProgress: FileUploadProgress;
  onRemoveFile?: (index: number) => void;
}

export const UploadProgress: React.FC<UploadProgressProps> = ({
  uploadProgress,
  onRemoveFile,
}) => {
  const { currentFileIndex, totalFiles, filesWithProgress } = uploadProgress;

  // Update progress bar widths
  React.useEffect(() => {
    const progressBars = document.querySelectorAll(
      ".upload-progress-bar-fill[data-progress]"
    );
    progressBars.forEach((bar: Element) => {
      const progress = bar.getAttribute("data-progress");
      if (progress && bar instanceof HTMLElement) {
        bar.style.width = `${progress}%`;
      }
    });
  }, [filesWithProgress]);

  const getFileIcon = (fileName: string): string => {
    const extension = fileName.split(".").pop()?.toLowerCase() || "";

    switch (extension) {
      case "pdf":
        return "PDF";
      case "doc":
      case "docx":
        return "WordDocument";
      case "xls":
      case "xlsx":
        return "ExcelDocument";
      case "ppt":
      case "pptx":
        return "PowerPointDocument";
      case "png":
      case "jpg":
      case "jpeg":
      case "gif":
      case "bmp":
        return "FileImage";
      case "zip":
      case "rar":
        return "ZipFolder";
      default:
        return "Page";
    }
  };

  const formatFileSize = (bytes: number): string => {
    if (bytes === 0) return "0 B";
    if (bytes <= 1) return "-- B"; // For existing files with dummy size

    const k = 1024;
    const sizes = ["B", "KB", "MB", "GB"];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + " " + sizes[i];
  };

  const completedCount = filesWithProgress.filter(
    (f) => f.uploadStatus === "completed"
  ).length;

  return (
    <div className="upload-progress-container">
      <div className="upload-progress-header">
        <Stack
          horizontal
          verticalAlign="center"
          styles={{ root: { marginBottom: "16px" } }}
        >
          <Icon
            iconName="CloudUpload"
            styles={{
              root: { fontSize: "16px", marginRight: "8px", color: "#0078D4" },
            }}
          />
          <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
            {completedCount} of {totalFiles} files uploaded
          </Text>
        </Stack>
      </div>

      <div className="upload-progress-files">
        {filesWithProgress.map((fileWithContent, index) => {
          const progress = fileWithContent.uploadProgress || 0;
          const status = fileWithContent.uploadStatus || "pending";

          return (
            <div
              key={`${fileWithContent.file.name}-${index}`}
              className="upload-progress-file"
            >
              <Stack
                horizontal
                verticalAlign="center"
                tokens={{ childrenGap: 12 }}
              >
                {/* File Icon */}
                <div className="upload-progress-file-icon">
                  <Icon
                    iconName={getFileIcon(fileWithContent.file.name)}
                    styles={{
                      root: {
                        fontSize: "24px",
                        color: status === "error" ? "#D13438" : "#605E5C",
                      },
                    }}
                  />
                </div>

                {/* File Details */}
                <Stack styles={{ root: { flex: 1, minWidth: 0 } }}>
                  <Stack
                    horizontal
                    verticalAlign="center"
                    tokens={{ childrenGap: 8 }}
                  >
                    <Text
                      variant="medium"
                      styles={{
                        root: {
                          fontWeight: 600,
                          color: status === "error" ? "#D13438" : "#323130",
                          overflow: "hidden",
                          textOverflow: "ellipsis",
                          whiteSpace: "nowrap",
                          maxWidth: "200px",
                        },
                      }}
                    >
                      {fileWithContent.file.name}
                    </Text>
                    <Text
                      variant="small"
                      styles={{
                        root: {
                          color: "#605E5C",
                          fontSize: "12px",
                        },
                      }}
                    >
                      {formatFileSize(fileWithContent.file.size)}
                    </Text>
                  </Stack>

                  {/* Progress Bar */}
                  <div className="upload-progress-bar-container">
                    <div className="upload-progress-bar-background">
                      <div
                        className={`upload-progress-bar-fill ${status}`}
                        data-progress={progress}
                      />
                    </div>
                  </div>

                  {/* Error Message */}
                  {status === "error" && fileWithContent.uploadError && (
                    <Text
                      variant="small"
                      styles={{
                        root: {
                          color: "#D13438",
                          fontSize: "11px",
                          marginTop: "2px",
                        },
                      }}
                    >
                      {fileWithContent.uploadError}
                    </Text>
                  )}
                </Stack>

                {/* Progress Percentage and Remove Button */}
                <Stack
                  horizontal
                  verticalAlign="center"
                  tokens={{ childrenGap: 8 }}
                >
                  <Text
                    variant="small"
                    styles={{
                      root: {
                        color: status === "error" ? "#D13438" : "#605E5C",
                        fontSize: "12px",
                        fontWeight: 600,
                        minWidth: "35px",
                        textAlign: "right",
                      },
                    }}
                  >
                    {status === "error" ? (
                      <Icon
                        iconName="ErrorBadge"
                        styles={{ root: { fontSize: "16px" } }}
                      />
                    ) : status === "completed" ? (
                      <Icon
                        iconName="CheckMark"
                        styles={{
                          root: { fontSize: "16px", color: "#107C10" },
                        }}
                      />
                    ) : (
                      `${Math.round(progress)}%`
                    )}
                  </Text>

                  {onRemoveFile &&
                    (status === "pending" || status === "error") && (
                      <IconButton
                        iconProps={{ iconName: "Cancel" }}
                        onClick={(e) => {
                          e.stopPropagation();
                          onRemoveFile(index);
                        }}
                        title="Remove file"
                        ariaLabel="Remove file"
                        styles={{
                          root: {
                            backgroundColor: "transparent",
                            border: "none",
                            color: "#605E5C",
                            width: "20px",
                            height: "20px",
                            minWidth: "20px",
                          },
                          rootHovered: {
                            backgroundColor: "#F3F2F1",
                            color: "#D13438",
                          },
                          icon: {
                            fontSize: "12px",
                          },
                        }}
                      />
                    )}
                </Stack>
              </Stack>
            </div>
          );
        })}
      </div>
    </div>
  );
};
