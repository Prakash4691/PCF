import * as React from "react";
import { Icon } from "@fluentui/react/lib/Icon";
import { initializeIcons } from "@fluentui/react/lib/Icons";
import { FileUploaderComponentProps } from "./types/interfaces";
import { DeleteConfirmationDialog } from "./components/dialogs/DeleteConfirmationDialog";
import { SizeLimitDialog } from "./components/dialogs/SizeLimitDialog";
import { FileGrid } from "./components/FileGrid/FileGrid";
import { FileHeader } from "./components/FileHeader";
import { UploadSection } from "./components/UploadSection";
import { UploadProgress } from "./components/UploadProgress";
import { getMimeTypeFromExtension } from "./utils/mimeTypes";

// Initialize the FluentUI icons
initializeIcons();

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
    uploadMessage,
    filesUploaded,
    uploadedFilesCount,
    newFilesCount,
    operationType,
    maxFileSizeForAttachment,
    blockedFileExtension,
    showDialog,
    uploadProgress,
  } = props;

  const [isDragging, setIsDragging] = React.useState(false);
  const [showDeleteDialog, setShowDeleteDialog] = React.useState(false);
  const [fileToDelete, setFileToDelete] = React.useState<{
    index: number;
    name: string;
  } | null>(null);
  const [showSizeLimitDialog, setShowSizeLimitDialog] = React.useState(false);
  const [dialogContent, setDialogContent] = React.useState<{
    title: string;
    subText: string;
  }>({
    title: "",
    subText: "",
  });
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  React.useEffect(() => {
    if (showDialog) {
      setDialogContent(showDialog);
      setShowSizeLimitDialog(true);
    }
  }, [showDialog]);

  React.useEffect(() => {
    if (selectedFiles.length === 0) {
      setShowDeleteDialog(false);
      setFileToDelete(null);
    }
  }, [selectedFiles]);

  const handleFileSelection = (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (files && files.length > 0) {
      const fileArray = Array.from(files);
      onFilesSelected(fileArray);
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
      onFilesSelected(fileArray, true);
    }
  };

  const handlePickFilesClick = async () => {
    try {
      const fileObjs = await context.device.pickFile({
        accept: "*",
        allowMultipleFiles: true,
        maximumAllowedFileSize: maxFileSizeForAttachment,
      });
      if (fileObjs && fileObjs.length > 0) {
        const blockedExtensions = blockedFileExtension
          ? blockedFileExtension
              .split(";")
              .map((ext: string) => ext.trim().toLowerCase())
          : [];

        const validFiles: File[] = [];
        const invalidFiles: { file: File; reason: string }[] = [];

        for (const fileObj of fileObjs) {
          const file = new File([fileObj.fileContent], fileObj.fileName, {
            type:
              fileObj.mimeType || getMimeTypeFromExtension(fileObj.fileName),
          });

          const fileExtension = file.name.split(".").pop()?.toLowerCase() || "";

          if (blockedExtensions.includes(fileExtension)) {
            invalidFiles.push({
              file,
              reason: `File type .${fileExtension} is not allowed`,
            });
            continue;
          }

          validFiles.push(file);
        }

        if (validFiles.length > 0) {
          onFilesSelected(validFiles, false);
        }

        if (invalidFiles.length > 0) {
          const errorMessages = invalidFiles.map(
            (invalid) => `${invalid.file.name}: ${invalid.reason}`
          );

          setDialogContent({
            title: "Blocked File Type",
            subText: `The following file(s) have blocked extensions: ${errorMessages.join(
              "; "
            )}. Please select files with allowed extensions.`,
          });
          setShowSizeLimitDialog(true);
        }
      }
    } catch (error) {
      setDialogContent({
        title: "File Size Limit Exceeded",
        subText: `The selected file(s) exceed the maximum allowed size of ${(
          maxFileSizeForAttachment / 1024
        ).toLocaleString()} KB (${maxFileSizeForAttachment.toLocaleString()} bytes). Please select smaller files.`,
      });
      setShowSizeLimitDialog(true);
    }
  };

  const handleFileRemove = (index: number) => {
    const fileInfo = selectedFiles[index];
    if (fileInfo.size <= 1) {
      setFileToDelete({ index, name: fileInfo.name });
      setShowDeleteDialog(true);
    } else {
      onFileRemoved(index);
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
          <Icon iconName="Upload" className="upload-icon" />
          <div className="file-uploader-text">
            <strong>Drag and drop files here</strong> or{" "}
            <strong>click to browse</strong>
          </div>
          <div className="file-uploader-subtext">
            Supported file types and size limits apply
          </div>
          <input
            type="file"
            ref={fileInputRef}
            className="hidden-file-input"
            multiple
            onChange={handleFileSelection}
            aria-label="File upload input"
            title="Choose files to upload"
          />
        </div>
      ) : (
        <>
          <FileHeader
            fileCount={
              // If there are new files pending upload show that count with Selected, else show uploaded count
              newFilesCount > 0 ? newFilesCount : uploadedFilesCount
            }
            mode={newFilesCount > 0 ? "selected" : "uploaded"}
            onAddFiles={handlePickFilesClick}
          />
          <FileGrid files={selectedFiles} onRemove={handleFileRemove} />
          {uploadProgress && uploadProgress.filesWithProgress.length > 0 && (
            <div className="upload-progress-wrapper">
              <UploadProgress
                uploadProgress={uploadProgress}
                onRemoveFile={handleFileRemove}
              />
            </div>
          )}
          <UploadSection
            isUploading={isUploading}
            operationType={operationType}
            uploadMessage={uploadMessage}
            onSubmit={onSubmitFiles}
            disabled={
              isUploading ||
              selectedFiles.length === 0 ||
              !selectedFiles.some((file) => file.size > 1) ||
              filesUploaded
            }
          />
        </>
      )}

      <DeleteConfirmationDialog
        isOpen={showDeleteDialog && fileToDelete !== null}
        fileName={fileToDelete?.name || ""}
        onDismiss={() => setShowDeleteDialog(false)}
        onConfirm={() => {
          if (fileToDelete !== null) {
            onFileRemoved(fileToDelete.index);
            setShowDeleteDialog(false);
          }
        }}
      />

      <SizeLimitDialog
        isOpen={showSizeLimitDialog}
        title={dialogContent.title}
        subText={dialogContent.subText}
        onDismiss={() => setShowSizeLimitDialog(false)}
      />
    </div>
  );
};
