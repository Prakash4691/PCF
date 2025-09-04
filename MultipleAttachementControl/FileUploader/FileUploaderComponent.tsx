import * as React from "react";
import { Icon } from "@fluentui/react/lib/Icon";
import { initializeIcons } from "@fluentui/react/lib/Icons";
import {
  FileUploaderComponentProps,
  FileWithContent,
} from "./types/interfaces";
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
import { inferMimeTypeFromFileName } from "./utils/mimeTypes";

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
    allocatedWidth,
    allocatedHeight,
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

  // Removed timeline message state - using only the main uploadMessage from parent

  // Timeline integration state
  const [timelineFiles, setTimelineFiles] = React.useState<FileWithContent[]>(
    []
  );
  const [isLoadingTimeline, setIsLoadingTimeline] = React.useState(false);
  const [mergedFiles, setMergedFiles] = React.useState<File[]>([]);

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

  // Function to fetch timeline notes
  const fetchTimelineFiles = React.useCallback(async () => {
    if (!parentRecordId || !notesServiceRef.current) return;

    setIsLoadingTimeline(true);
    try {
      const timelineNotes = await notesServiceRef.current.getAllNotesForRecord(
        parentRecordId
      );
      const timelineFileObjects: FileWithContent[] = [];

      for (const note of timelineNotes) {
        // Create placeholder File object similar to existing files
        const extension = note.filename.split(".").pop()?.toLowerCase() || "";
        const mimeType = note.mimetype || inferMimeTypeFromFileName(extension);

        const file = new File([new ArrayBuffer(1)], note.filename, {
          type: mimeType,
        });

        // Attach notesId to file object for preview/download
        (file as FileWithNoteId).notesId = note.annotationid;

        const timelineFile: FileWithContent = {
          file,
          isExisting: true,
          notesId: note.annotationid,
          uploadProgress: 100,
          uploadStatus: "completed",
          source: "timeline",
          guid: note.annotationid,
          subject: note.subject,
          noteText: note.notetext,
          createdOn: note.createdon ? new Date(note.createdon) : undefined,
          modifiedOn: note.modifiedon ? new Date(note.modifiedon) : undefined,
          createdBy: note.createdby,
        };

        timelineFileObjects.push(timelineFile);
      }

      setTimelineFiles(timelineFileObjects);
    } catch (error) {
      console.error("Error fetching timeline files:", error);
      setTimelineFiles([]); // Set empty array on error
    } finally {
      setIsLoadingTimeline(false);
    }
  }, [parentRecordId]);

  // Function to merge files intelligently (prevent duplicates)
  const mergeFiles = React.useCallback(
    (pcfFiles: File[], timelineFiles: FileWithContent[]): File[] => {
      console.log(
        "Merging files - PCF:",
        pcfFiles.length,
        "Timeline:",
        timelineFiles.length
      );

      const merged: File[] = [];
      const processedGuids = new Set<string>();
      const processedNames = new Set<string>();
      const timelineFilesByGuid = new Map<string, FileWithContent>();
      const timelineFilesByName = new Map<string, FileWithContent>();

      // Index timeline files by both GUID and name for efficient lookup
      timelineFiles.forEach((timelineFile) => {
        if (timelineFile.guid) {
          timelineFilesByGuid.set(timelineFile.guid, timelineFile);
        }
        timelineFilesByName.set(timelineFile.file.name, timelineFile);
      });

      // Process PCF files first - these are authoritative (from actual control fileData)
      pcfFiles.forEach((file) => {
        const fileWithNoteId = file as FileWithNoteId;
        const pcfGuid = fileWithNoteId.notesId;

        // Always add PCF file to merged list - it's authoritative
        merged.push(file);
        processedNames.add(file.name);

        if (pcfGuid) {
          processedGuids.add(pcfGuid);
        }
      });

      // Add only timeline files that DON'T have corresponding PCF files
      timelineFiles.forEach((timelineFile) => {
        const alreadyProcessedByGuid =
          timelineFile.guid && processedGuids.has(timelineFile.guid);
        const alreadyProcessedByName = processedNames.has(
          timelineFile.file.name
        );

        // Only add timeline files that are NOT already represented by PCF files
        if (!alreadyProcessedByGuid && !alreadyProcessedByName) {
          merged.push(timelineFile.file);
          if (timelineFile.guid) {
            processedGuids.add(timelineFile.guid);
          }
          processedNames.add(timelineFile.file.name);
        }
      });

      console.log("Merged", merged.length, "files total");

      return merged;
    },
    []
  );

  // Effect to fetch timeline files when parent record changes
  React.useEffect(() => {
    if (parentRecordId) {
      fetchTimelineFiles();
    }
  }, [parentRecordId, fetchTimelineFiles]);

  // Effect to merge PCF and timeline files
  React.useEffect(() => {
    const merged = mergeFiles(selectedFiles, timelineFiles);
    setMergedFiles(merged);
  }, [selectedFiles, timelineFiles, mergeFiles]);

  // Create metadata map for enhanced file display
  const fileMetadata = React.useMemo(() => {
    const metadataMap = new Map<string, FileWithContent>();

    // Create metadata with proper priority: PCF files (from control) take precedence over timeline
    let pcfCount = 0,
      timelineCount = 0;

    mergedFiles.forEach((file) => {
      const pcfFile = selectedFiles.find((sf) => sf.name === file.name);
      const timelineFile = timelineFiles.find(
        (tf) => tf.file.name === file.name
      );

      // Priority 1: PCF file exists - this means user uploaded via control, use fileupload source
      if (pcfFile) {
        pcfCount++;
        metadataMap.set(file.name, {
          file,
          isExisting: file.size <= 1,
          source: "fileupload",
          uploadStatus: file.size <= 1 ? "completed" : "pending",
          uploadProgress: file.size <= 1 ? 100 : 0,
          guid: (pcfFile as any).notesId, // Preserve notesId if available
        });
      }
      // Priority 2: Timeline-only file (no corresponding PCF file)
      else if (timelineFile) {
        timelineCount++;
        metadataMap.set(file.name, timelineFile);
      }
    });

    console.log(
      `Metadata created: ${pcfCount} PCF files, ${timelineCount} timeline-only files`
    );

    return metadataMap;
  }, [selectedFiles, timelineFiles, mergedFiles]);

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

  // Handler for timeline file deletion
  const handleTimelineFileDelete = async (fileName: string) => {
    try {
      const metadata = fileMetadata.get(fileName);
      if (!metadata?.guid || metadata.source !== "timeline") return;

      const svc = notesServiceRef.current!;
      await svc.deleteNote(metadata.guid);

      // Remove from timeline files state
      setTimelineFiles((prev) => prev.filter((f) => f.file.name !== fileName));

      setShowDeleteDialog(false);
    } catch (error) {
      console.error("Error deleting timeline file:", error);
    }
  };

  const handleFileRemove = (index: number) => {
    const file = mergedFiles[index];
    const metadata = fileMetadata.get(file.name);

    // Determine if this is a timeline file or PCF file
    if (metadata?.source === "timeline") {
      // Timeline file - handle delete directly
      if (file.size <= 1) {
        setFileToDelete({ index, name: file.name });
        setShowDeleteDialog(true);
      }
    } else {
      // PCF file - find the corresponding index in selectedFiles and delegate to parent
      const selectedFileIndex = selectedFiles.findIndex(
        (f) => f.name === file.name
      );
      if (selectedFileIndex !== -1) {
        if (file.size <= 1) {
          setFileToDelete({ index: selectedFileIndex, name: file.name });
          setShowDeleteDialog(true);
        } else {
          onFileRemoved(selectedFileIndex);
        }
      }
    }
  };

  const handlePreview = async (index: number) => {
    const file: File = mergedFiles[index];
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
    const file: File = mergedFiles[index];
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
      className={`file-uploader-container ${isDragging ? "drag-active" : ""} ${
        // Treat phone form factor as extra-small to enforce single-column stacking
        (context.client &&
          context.client.getFormFactor &&
          context.client.getFormFactor() === 3) ||
        (allocatedWidth && allocatedWidth < 450)
          ? "is-xs"
          : allocatedWidth && allocatedWidth < 700
          ? "is-sm"
          : allocatedWidth && allocatedWidth < 1024
          ? "is-md"
          : "is-lg"
      } ${
        typeof allocatedHeight === "number" && allocatedHeight > 0
          ? "has-host-height"
          : "auto-height"
      }`}
      data-allocated-width={allocatedWidth}
      data-allocated-height={allocatedHeight}
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
    >
      {mergedFiles.length === 0 && !isLoadingTimeline ? (
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
      ) : isLoadingTimeline && mergedFiles.length === 0 ? (
        <div className="loading-timeline">
          <Icon iconName="Loading" className="loading-icon" />
          <div>Loading timeline files...</div>
        </div>
      ) : (
        <>
          <FileHeader
            fileCount={newFilesCount > 0 ? newFilesCount : mergedFiles.length} // Show total count of all files (PCF + timeline)
            mode={newFilesCount > 0 ? "selected" : "uploaded"}
            onAddFiles={handlePickFilesClick}
          />
          <FileGrid
            files={mergedFiles}
            fileMetadata={fileMetadata}
            onRemove={handleFileRemove}
            onPreview={handlePreview}
            onDownload={handleDownload}
          />
          {/* Upload message is now displayed only in UploadSection component */}
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
            const metadata = fileMetadata.get(fileToDelete.name);
            if (metadata?.source === "timeline") {
              handleTimelineFileDelete(fileToDelete.name);
            } else {
              onFileRemoved(fileToDelete.index);
              setShowDeleteDialog(false);
            }
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
