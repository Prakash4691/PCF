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
import { base64ToUint8Array } from "./utils/fileUtils";
import { PreviewDialog } from "./components/dialogs/PreviewDialog";
import {
  base64ToBlob,
  getObjectUrl,
  isPreviewable,
  isTextType,
  readFileAsText,
  triggerBrowserDownload,
} from "./utils/filePreviewUtils";
import { DataverseNotesOperations } from "./services/DataverseNotesOperations";

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
  const notesServiceRef = React.useRef<DataverseNotesOperations | null>(null);

  const [previewData, setPreviewData] = React.useState<{
    fileName: string;
    mimeType: string;
    objectUrl?: string;
    textContent?: string;
  } | null>(null);
  const [isPreviewOpen, setIsPreviewOpen] = React.useState(false);
  const [isPreviewLoading, setIsPreviewLoading] = React.useState(false);

  if (!notesServiceRef.current) {
    notesServiceRef.current = new DataverseNotesOperations(context);
  }

  // Safely extract parentRecordId without using 'any'
  interface ParentRecordContextParam {
    parameters: {
      parentRecordId?: { raw?: string };
    };
  }
  const parentRecordId =
    (context as unknown as ParentRecordContextParam).parameters.parentRecordId
      ?.raw || "";

  // Extend File interface for notesId metadata (existing uploaded placeholder files)
  interface FileWithNoteId extends File {
    notesId?: string;
  }
  const hasNotesId = (f: File): f is FileWithNoteId =>
    Object.prototype.hasOwnProperty.call(f, "notesId") &&
    typeof (f as FileWithNoteId).notesId === "string" &&
    !!(f as FileWithNoteId).notesId;

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
          // Device.pickFile returns base64 WITHOUT data URL prefix (per docs). We must decode to bytes to prevent corruption.
          let bytes: Uint8Array;
          try {
            bytes = base64ToUint8Array(fileObj.fileContent);
          } catch (e) {
            console.warn(
              "Failed to decode base64 from pickFile, falling back to raw string",
              e
            );
            bytes = new TextEncoder().encode(fileObj.fileContent);
          }
          const mime =
            fileObj.mimeType || getMimeTypeFromExtension(fileObj.fileName);
          // Use a copy of the underlying ArrayBuffer to satisfy BlobPart typing across environments
          const arrayBuffer: ArrayBuffer = new Uint8Array(bytes).buffer;
          const file = new File([arrayBuffer], fileObj.fileName, {
            type: mime,
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

  const handlePreview = async (index: number) => {
    const file: File = selectedFiles[index];
    setIsPreviewOpen(true);
    setIsPreviewLoading(true);
    try {
      // Determine if existing (placeholder size <=1) meaning we need to fetch from notes
      if (file.size <= 1) {
        const svc = notesServiceRef.current!;
        // Fallback: attempt to resolve notesId even if not yet attached
        let notesId: string | undefined;
        if (hasNotesId(file)) {
          notesId = file.notesId;
        } else {
          const fileName: string = (file as File).name;
          notesId = await svc.getNotesId(fileName, parentRecordId);
        }
        if (!notesId) {
          setPreviewData({
            fileName: file.name,
            mimeType: file.type || "application/octet-stream",
          });
          return;
        }
        if (!hasNotesId(file)) {
          (file as FileWithNoteId).notesId = notesId;
        }
        const note = await svc.retrieveNote(notesId);
        if (note) {
          const blob = base64ToBlob(note.base64, note.mimeType);
          if (isTextType(note.mimeType)) {
            const text = await blob.text();
            setPreviewData({
              fileName: note.fileName,
              mimeType: note.mimeType,
              textContent: text,
            });
          } else if (isPreviewable(note.mimeType)) {
            const url = getObjectUrl(notesId, blob);
            setPreviewData({
              fileName: note.fileName,
              mimeType: note.mimeType,
              objectUrl: url,
            });
          } else {
            setPreviewData({
              fileName: note.fileName,
              mimeType: note.mimeType,
            });
          }
        }
      } else {
        // New file chosen locally
        const mimeType = file.type || getMimeTypeFromExtension(file.name);
        if (isTextType(mimeType)) {
          const text = await readFileAsText(file);
          setPreviewData({ fileName: file.name, mimeType, textContent: text });
        } else if (isPreviewable(mimeType)) {
          const blob = file;
          const url = getObjectUrl(file.name + file.size, blob);
          setPreviewData({ fileName: file.name, mimeType, objectUrl: url });
        } else {
          setPreviewData({ fileName: file.name, mimeType });
        }
      }
    } finally {
      setIsPreviewLoading(false);
    }
  };

  const handleDownload = async (index: number) => {
    const file: File = selectedFiles[index];
    try {
      if (file.size <= 1) {
        const svc = notesServiceRef.current!;
        let notesId: string | undefined;
        if (hasNotesId(file)) {
          notesId = file.notesId;
        } else {
          const fileName: string = (file as File).name;
          notesId = await svc.getNotesId(fileName, parentRecordId);
        }
        if (!notesId) return;
        if (!hasNotesId(file)) {
          (file as FileWithNoteId).notesId = notesId;
        }
        const note = await svc.retrieveNote(notesId);
        if (note) {
          const blob = base64ToBlob(note.base64, note.mimeType);
          triggerBrowserDownload(blob, note.fileName);
        }
      } else {
        triggerBrowserDownload(file, file.name);
      }
    } catch (e) {
      console.error("Download failed", e);
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
          <FileGrid
            files={selectedFiles}
            onRemove={handleFileRemove}
            onPreview={handlePreview}
            onDownload={handleDownload}
          />
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
      <PreviewDialog
        isOpen={isPreviewOpen}
        data={previewData}
        isLoading={isPreviewLoading}
        onDismiss={() => {
          setIsPreviewOpen(false);
          setPreviewData(null);
        }}
      />
    </div>
  );
};
