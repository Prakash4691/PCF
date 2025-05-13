import * as React from "react";
import { Icon } from "@fluentui/react/lib/Icon";
import { initializeIcons } from "@fluentui/react/lib/Icons";
import { IInputs } from "./generated/ManifestTypes";
import { Text } from "@fluentui/react/lib/Text";
import { Label } from "@fluentui/react/lib/Label";
import { Separator } from "@fluentui/react/lib/Separator";
import { TooltipHost } from "@fluentui/react/lib/Tooltip";
import { PrimaryButton } from "@fluentui/react/lib/Button";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";

// Initialize the FluentUI icons
initializeIcons();

interface FileUploaderComponentProps {
  selectedFiles: File[];
  onFilesSelected: (files: File[]) => void;
  onFileRemoved: (index: number) => void;
  onSubmitFiles: () => void;
  context: ComponentFramework.Context<IInputs>;
  isUploading: boolean;
  uploadProgress?: number;
  uploadMessage?: { text: string; type: MessageBarType } | null;
  hasExistingFiles: boolean;
  filesUploaded: boolean;
  operationType?: string;
}

interface FileInfo {
  file: File;
  icon: string;
  sizeMb: string;
  sizeText: string;
  fileType: string;
  isExistingFile: boolean;
}

export const FileUploaderComponent: React.FC<FileUploaderComponentProps> = (
  props
) => {
  const {
    selectedFiles,
    onFilesSelected,
    onFileRemoved,
    onSubmitFiles,
    context,
    isUploading,
    uploadProgress,
    uploadMessage,
    hasExistingFiles,
    filesUploaded,
    operationType,
  } = props;
  const [isDragging, setIsDragging] = React.useState(false);
  const [hasOverflow, setHasOverflow] = React.useState(false);
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const fileGridRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    const checkOverflow = () => {
      if (fileGridRef.current) {
        setHasOverflow(
          fileGridRef.current.scrollWidth > fileGridRef.current.clientWidth
        );
      }
    };

    checkOverflow();
    // Add resize listener to check when window size changes
    window.addEventListener("resize", checkOverflow);
    return () => window.removeEventListener("resize", checkOverflow);
  }, [selectedFiles]);

  const getFileInfo = (file: File): FileInfo => {
    const fileExtension = file.name.split(".").pop()?.toLowerCase() || "";
    const sizeMb = (file.size / 1024 / 1024).toFixed(2);
    let fileType = file.type;
    let icon = "Document";
    // If file size is very small (1 byte), mark as existing file (dummy file)
    const isExistingFile = file.size <= 1;

    // Map file types to appropriate icons
    if (
      fileType.includes("image") ||
      ["jpg", "jpeg", "png", "gif", "bmp", "webp"].includes(fileExtension)
    ) {
      icon = "FileImage";
      fileType = "Image";
    } else if (fileType.includes("pdf") || fileExtension === "pdf") {
      icon = "PDF";
      fileType = "PDF Document";
    } else if (
      fileType.includes("spreadsheet") ||
      ["xls", "xlsx", "csv"].includes(fileExtension)
    ) {
      icon = "ExcelDocument";
      fileType = "Spreadsheet";
    } else if (
      fileType.includes("word") ||
      ["doc", "docx", "rtf"].includes(fileExtension)
    ) {
      icon = "WordDocument";
      fileType = "Word Document";
    } else if (
      fileType.includes("presentation") ||
      ["ppt", "pptx"].includes(fileExtension)
    ) {
      icon = "PowerPointDocument";
      fileType = "Presentation";
    } else if (
      fileType.includes("text") ||
      ["txt", "text"].includes(fileExtension)
    ) {
      icon = "TextDocument";
      fileType = "Text Document";
    } else if (["zip", "rar", "7z", "tar", "gz"].includes(fileExtension)) {
      icon = "ZipFolder";
      fileType = "Archive";
    } else if (
      ["mp4", "mov", "avi", "wmv", "flv", "webm"].includes(fileExtension)
    ) {
      icon = "Video";
      fileType = "Video";
    } else if (["mp3", "wav", "ogg", "flac", "m4a"].includes(fileExtension)) {
      icon = "MusicInCollection";
      fileType = "Audio";
    }

    // Format the size text (e.g., "2.50 MB")
    let sizeText = "";
    if (isExistingFile) {
      sizeText = "Existing File";
    } else if (file.size < 1024) {
      sizeText = `${file.size} B`;
    } else if (file.size < 1024 * 1024) {
      sizeText = `${(file.size / 1024).toFixed(2)} KB`;
    } else {
      sizeText = `${sizeMb} MB`;
    }

    return {
      file,
      icon,
      sizeMb,
      sizeText,
      fileType,
      isExistingFile,
    };
  };

  const handleFileSelection = (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (files && files.length > 0) {
      const fileArray = Array.from(files);
      onFilesSelected(fileArray);
      // Clear the input value to allow the same file to be selected again
      if (fileInputRef.current) {
        fileInputRef.current.value = "";
      }
    }
  };

  const handleDragOver = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    event.stopPropagation();
    setIsDragging(true);
  };

  const handleDragLeave = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    event.stopPropagation();
    setIsDragging(false);
  };

  const handleDrop = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    event.stopPropagation();
    setIsDragging(false);

    const files = event.dataTransfer.files;
    if (files && files.length > 0) {
      const fileArray = Array.from(files);
      onFilesSelected(fileArray);
    }
  };

  const handlePickFilesClick = async () => {
    try {
      const fileObjs = await context.device.pickFile({
        accept: "*",
        allowMultipleFiles: true,
        maximumAllowedFileSize: 10485760, // 10MB - typical Dataverse attachment limit
      });
      if (fileObjs && fileObjs.length > 0) {
        const files: File[] = [];
        for (const fileObj of fileObjs) {
          // Create a File object from the FileObject
          const file = new File([fileObj.fileContent], fileObj.fileName, {
            type: fileObj.mimeType,
          });
          files.push(file);
        }
        onFilesSelected(files);
      }
    } catch (error) {
      console.error("Error picking files:", error);
    }
  };

  return (
    <div
      className={`file-uploader-container ${isDragging ? "drag-active" : ""}`}
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
    >
      {selectedFiles.length === 0 ? (
        <div className="file-uploader-content" onClick={handlePickFilesClick}>
          <Icon
            iconName="Upload"
            style={{ fontSize: "24px", color: "#0078d4" }}
          />
          <div className="file-uploader-text">
            Drag and drop files here or click to browse
          </div>

          <input
            type="file"
            ref={fileInputRef}
            style={{ display: "none" }}
            multiple
            onChange={handleFileSelection}
          />
        </div>
      ) : (
        <>
          <div className="files-header">
            <Label>Selected Files ({selectedFiles.length})</Label>
            <button className="add-files-button" onClick={handlePickFilesClick}>
              <Icon iconName="Add" />
              <span>Add Files</span>
            </button>
          </div>
          <Separator />
          {hasOverflow && (
            <div className="scroll-indicator">
              <Icon iconName="DoubleChevronLeft8" /> Scroll to see more files{" "}
              <Icon iconName="DoubleChevronRight8" />
            </div>
          )}
          <div className="file-grid" ref={fileGridRef}>
            {selectedFiles.map((file, index) => {
              const fileInfo = getFileInfo(file);
              return (
                <div
                  key={`${file.name}-${index}`}
                  className={`file-card ${
                    fileInfo.isExistingFile ? "existing-file" : ""
                  }`}
                >
                  <div className="file-icon">
                    <Icon
                      iconName={fileInfo.icon}
                      style={{ fontSize: "32px" }}
                    />
                  </div>
                  <div className="file-details">
                    <div className="file-name" title={file.name}>
                      {file.name}
                    </div>
                    <div className="file-meta">
                      <Text variant="small">{fileInfo.fileType}</Text>
                      <Text variant="small" className="file-size">
                        {fileInfo.isExistingFile && (
                          <TooltipHost content="This file exists in the record but needs to be uploaded to create notes">
                            <span className="existing-file-indicator">
                              <Icon
                                iconName="InfoSolid"
                                style={{ fontSize: "10px", marginRight: "4px" }}
                              />
                              {fileInfo.sizeText}
                            </span>
                          </TooltipHost>
                        )}
                        {!fileInfo.isExistingFile && fileInfo.sizeText}
                      </Text>
                    </div>
                  </div>
                  <button
                    className="remove-button"
                    onClick={(e) => {
                      e.stopPropagation();
                      onFileRemoved(index);
                    }}
                    title="Remove file"
                  >
                    <Icon iconName="Cancel" />
                  </button>
                </div>
              );
            })}
          </div>

          {/* Upload Status and Button Section */}
          <div className="upload-section">
            {isUploading && (
              <div className="progress-container">
                <Spinner
                  size={SpinnerSize.large}
                  label={operationType}
                  labelPosition="right"
                  styles={{ label: { marginLeft: "8px" } }}
                />
                {/* <div
                  className="progress-text"
                  style={{ marginTop: "4px", textAlign: "center" }}
                >
                  {`${Math.round(uploadProgress * 100)}% complete`}
                </div> */}
              </div>
            )}

            {uploadMessage && (
              <MessageBar
                messageBarType={uploadMessage.type}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
              >
                {uploadMessage.text}
              </MessageBar>
            )}

            <div className="submit-button-container">
              <PrimaryButton
                text={isUploading ? "Uploading..." : "Upload Files"}
                /*  onClick={async () => {
                  try {
                    await onSubmitFiles();
                  } catch (error) {
                    console.error("Error in upload:", error);
                  }
                }}*/
                onClick={onSubmitFiles}
                disabled={
                  isUploading ||
                  selectedFiles.length === 0 ||
                  !selectedFiles.some((file) => file.size > 1) || // Enable only if there's at least one new file (size > 1 byte)
                  filesUploaded
                }
                iconProps={{
                  iconName: isUploading ? "Spinner" : "CloudUpload",
                }}
              />
            </div>
          </div>
        </>
      )}
    </div>
  );
};
