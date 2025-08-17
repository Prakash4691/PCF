import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import { FileUploaderComponent } from "./FileUploaderComponent";
import { MessageBarType } from "@fluentui/react/lib/MessageBar";
import {
  FileWithContent,
  UploadMessage,
  FileUploadProgress,
} from "./types/interfaces";
import { DataverseNotesOperations } from "./services/DataverseNotesOperations";
import { readFileAsDataURL } from "./utils/fileUtils";
import { inferMimeTypeFromFileName } from "./utils/mimeTypes";

export class MultipleFileUploader
  implements ComponentFramework.ReactControl<IInputs, IOutputs>
{
  private notifyOutputChanged: () => void;
  private selectedFiles: FileWithContent[] = [];
  private fileDataValue = "";
  private existingFileNames: string[] = [];
  private isUploading = false;
  private uploadMessage: UploadMessage | null = null;
  private filesUploaded = false;
  private context: ComponentFramework.Context<IInputs>;
  private parentRecordId: string;
  private operationType: string;
  private blockedFileExtension = "";
  private maxFileSizeForAttachment: number;
  private showDialog: boolean;
  private dataverseService: DataverseNotesOperations;
  private uploadProgress: FileUploadProgress | null = null;
  //private triggerSaveRefresh = "";

  /**
   * Used to initialize the control instance
   */
  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary
  ): void {
    this.notifyOutputChanged = notifyOutputChanged;
    this.context = context;
    this.dataverseService = new DataverseNotesOperations(context);
    // IMPORTANT: assign parentRecordId BEFORE parsing existing files so notesId lookups work
    this.parentRecordId = context.parameters.parentRecordId.raw;

    // Request resize tracking so allocatedWidth/allocatedHeight are updated (per Microsoft docs trackContainerResize)
    try {
      // true indicates we want height updates as well
      this.context.mode.trackContainerResize(true);
    } catch (e) {
      // Fail silently if API not available (older harness versions)
      console.warn("trackContainerResize not available", e);
    }

    if (context.parameters.fileData.raw) {
      this.parseExistingFieldValue(context.parameters.fileData.raw);
    }
    this.blockedFileExtension =
      context.parameters.blockedFileExtension.raw ?? "";
    this.maxFileSizeForAttachment = context.parameters.maxFileSizeForAttachment
      .raw
      ? parseInt(context.parameters.maxFileSizeForAttachment.raw) * 1024
      : 0;
    this.showDialog = false;
  }

  /**
   * Called when any value in the property bag has changed
   */
  public updateView(
    context: ComponentFramework.Context<IInputs>
  ): React.ReactElement {
    if (
      context.parameters.fileData.raw !== this.fileDataValue &&
      context.parameters.fileData.raw
    ) {
      this.parseExistingFieldValue(context.parameters.fileData.raw);
    }

    this.context = context;

    // Fallback: if existing files present but none have resolved notesId yet and we now have parentRecordId, attempt resolution
    if (
      this.parentRecordId &&
      this.selectedFiles.some(
        (f) => f.isExisting && (!f.notesId || typeof f.notesId !== "string")
      )
    ) {
      this.resolveExistingNotesIds();
    }

    return React.createElement(FileUploaderComponent, {
      selectedFiles: this.selectedFiles.map(
        (fileWithContent) => fileWithContent.file
      ),
      onFilesSelected: this.onFilesSelected,
      onFileRemoved: this.onFileRemoved,
      onSubmitFiles: this.onSubmitFiles,
      context: context,
      isUploading: this.isUploading,
      uploadMessage: this.uploadMessage,
      hasExistingFiles: this.existingFileNames.length > 0,
      filesUploaded: this.filesUploaded,
      uploadedFilesCount: this.selectedFiles.filter(
        (f) => f.isExisting || f.uploadStatus === "completed"
      ).length,
      newFilesCount: this.selectedFiles.filter(
        (f) => !f.isExisting && f.uploadStatus !== "completed"
      ).length,
      operationType: this.operationType,
      maxFileSizeForAttachment: this.maxFileSizeForAttachment,
      blockedFileExtension: this.blockedFileExtension,
      showDialog: null,
      uploadProgress: this.uploadProgress,
      allocatedWidth: context.mode?.allocatedWidth,
      allocatedHeight: context.mode?.allocatedHeight,
    });
  }

  /**
   * Parse existing field value into file names array
   */
  private parseExistingFieldValue(fieldValue: string): void {
    if (!fieldValue) return;

    const rawSegments = fieldValue
      .split(",")
      .map((s) => s.trim())
      .filter((s) => s.length > 0);

    const decoded: string[] = [];

    // Helper: known extensions to determine filename boundaries for legacy (unencoded) storage
    const knownExt = [
      "jpg",
      "jpeg",
      "png",
      "gif",
      "bmp",
      "webp",
      "svg",
      "ico",
      "pdf",
      "doc",
      "docx",
      "xls",
      "xlsx",
      "ppt",
      "pptx",
      "txt",
    ];

    // Detect if new encoded format (every segment decodes cleanly & contains no leftover % except for encoded sequences)
    const looksEncoded = rawSegments.some((seg) => /%[0-9A-Fa-f]{2}/.test(seg));

    if (looksEncoded) {
      for (const seg of rawSegments) {
        try {
          decoded.push(decodeURIComponent(seg));
        } catch {
          decoded.push(seg); // fallback
        }
      }
    } else {
      // Legacy: filenames may have been split on commas that actually belonged to the filename.
      let accumulator: string[] = [];
      const flush = () => {
        if (accumulator.length) {
          decoded.push(accumulator.join(", "));
          accumulator = [];
        }
      };
      for (const seg of rawSegments) {
        accumulator.push(seg);
        const candidate = accumulator.join(", ");
        const ext = candidate.split(".").pop()?.toLowerCase() || "";
        if (knownExt.includes(ext)) {
          flush();
        }
      }
      flush();
    }

    this.existingFileNames = decoded.filter((n) => n.length > 0);
    this.createDummyFiles();
  }

  /**
   * Create dummy File objects for existing file names
   */
  private createDummyFiles(): void {
    if (this.existingFileNames.length === 0) return;

    const dummyFiles: FileWithContent[] = this.existingFileNames.map(
      (fileName) => {
        const extension = fileName.split(".").pop()?.toLowerCase() || "";
        const mimeType = inferMimeTypeFromFileName(extension);

        const file = new File([new ArrayBuffer(1)], fileName, {
          type: mimeType,
        });

        let notesId: Promise<string | null> | undefined;
        if (this.parentRecordId) {
          notesId = this.dataverseService.getNotesId(
            fileName,
            this.parentRecordId
          );
        }

        const existingFile = this.selectedFiles.find(
          (f) => f.file.name === fileName
        );
        return {
          file,
          isExisting: true,
          notesId: existingFile?.notesId || notesId,
          uploadProgress: 100,
          uploadStatus: "completed" as const,
        };
      }
    );

    this.selectedFiles = dummyFiles;
    this.updateFileDataValue();
    // Resolve notes ids asynchronously and attach to File objects
    this.resolveExistingNotesIds();
  }

  /** Resolve pending notesId promises to concrete strings and attach to file objects */
  private resolveExistingNotesIds(): void {
    const pending = this.selectedFiles.filter(
      (f) => f.isExisting && f.notesId && typeof f.notesId !== "string"
    );
    if (pending.length === 0) return;
    pending.forEach(async (f) => {
      try {
        const id = await f.notesId;
        if (id) {
          f.file.notesId = id; // augment File so React component can access
          f.notesId = id; // cache resolved
          this.notifyOutputChanged();
        }
      } catch (e) {
        console.error("Failed resolving notesId for", f.file.name, e);
      }
    });
  }

  private updateFileDataValue(): void {
    // Store as comma-separated list of URI-encoded names to avoid ambiguity with commas inside filenames
    this.fileDataValue = this.selectedFiles
      .map((f) => encodeURIComponent(f.file.name))
      .join(",");
  }

  /**
   * Update progress for a specific file
   */
  private updateFileProgress(
    fileIndex: number,
    progress: number,
    status: "pending" | "uploading" | "completed" | "error",
    error?: string
  ): void {
    if (fileIndex < this.selectedFiles.length) {
      this.selectedFiles[fileIndex].uploadProgress = progress;
      this.selectedFiles[fileIndex].uploadStatus = status;
      if (error) {
        this.selectedFiles[fileIndex].uploadError = error;
      }

      // Update upload progress state
      const filesToUpload = this.selectedFiles.filter((f) => !f.isExisting);
      this.uploadProgress = {
        currentFileIndex: fileIndex,
        totalFiles: filesToUpload.length,
        filesWithProgress: [...filesToUpload],
      };

      this.notifyOutputChanged();
    }
  }

  /**
   * Initialize upload progress tracking
   */
  private initializeUploadProgress(): void {
    const filesToUpload = this.selectedFiles.filter((f) => !f.isExisting);

    // Initialize all files with pending status
    filesToUpload.forEach((file, index) => {
      file.uploadProgress = 0;
      file.uploadStatus = "pending";
      file.uploadError = undefined;
    });

    this.uploadProgress = {
      currentFileIndex: 0,
      totalFiles: filesToUpload.length,
      filesWithProgress: [...filesToUpload],
    };
  }

  /**
   * Called when the user selects files
   */
  private onFilesSelected = async (
    files: File[],
    isDragAndDrop: boolean
  ): Promise<void> => {
    try {
      const existingNames = this.selectedFiles.map((f) =>
        f.file.name.toLowerCase()
      );
      const nonDuplicateFiles: File[] = [];
      const duplicateFiles: File[] = [];

      for (const file of files) {
        if (existingNames.includes(file.name.toLowerCase())) {
          duplicateFiles.push(file);
        } else {
          nonDuplicateFiles.push(file);
        }
      }

      if (duplicateFiles.length > 0) {
        this.showDialog = true;
        this.notifyOutputChanged();
        return;
      }

      if (nonDuplicateFiles.length > 0) {
        const newFiles = nonDuplicateFiles.map((file) => ({
          file,
          isExisting: false,
          isDragAndDrop: isDragAndDrop,
          uploadProgress: 0,
          uploadStatus: "pending" as const,
        }));
        this.selectedFiles = [...this.selectedFiles, ...newFiles];
        this.updateFileDataValue();
        this.filesUploaded = false;
        this.notifyOutputChanged();
      }
    } catch (error) {
      console.error("Error in onFilesSelected:", error);
      this.uploadMessage = {
        text: `Error processing selected files: ${(error as Error).message}`,
        type: MessageBarType.error,
      };
      this.notifyOutputChanged();
    }
  };

  /**
   * Called when a file is removed
   */
  private onFileRemoved = async (index: number): Promise<void> => {
    const removedFile = this.selectedFiles[index];
    if (removedFile.isExisting) {
      this.uploadMessage = null;
      this.operationType = "Deleting Notes...";
      this.isUploading = true;
      this.notifyOutputChanged();

      try {
        const fileName = removedFile.file.name;
        this.existingFileNames = this.existingFileNames.filter(
          (name) => name !== fileName
        );

        let notesId: string | null = null;
        if (removedFile.notesId) {
          if (typeof removedFile.notesId === "string") {
            notesId = removedFile.notesId;
          } else {
            try {
              notesId = await removedFile.notesId;
            } catch (error) {
              console.error("Error retrieving notes ID:", error);
              throw new Error(`Failed to get notes ID for file ${fileName}`);
            }
          }
        }

        if (!notesId && this.parentRecordId) {
          try {
            notesId = await this.dataverseService.getNotesId(
              fileName,
              this.parentRecordId
            );
          } catch (error) {
            console.error("Error retrieving notes ID:", error);
            throw new Error(`Failed to get notes ID for file ${fileName}`);
          }
        }

        if (notesId) {
          await this.dataverseService.deleteNote(notesId);
        } else {
          console.warn(
            `No notes ID found for file ${fileName}. File may have already been deleted.`
          );
        }

        this.selectedFiles = [
          ...this.selectedFiles.slice(0, index),
          ...this.selectedFiles.slice(index + 1),
        ];

        this.updateFileDataValue();
        this.isUploading = false;
        this.uploadMessage = {
          text: `Successfully deleted notes for ${fileName}`,
          type: MessageBarType.success,
        };
      } catch (error) {
        console.error("Error deleting file:", error);
        this.isUploading = false;
        this.uploadMessage = {
          text: `Error deleting file: ${(error as Error).message}`,
          type: MessageBarType.error,
        };
      }
    } else {
      this.selectedFiles = [
        ...this.selectedFiles.slice(0, index),
        ...this.selectedFiles.slice(index + 1),
      ];
      this.updateFileDataValue();
    }

    this.notifyOutputChanged();
  };

  /**
   * Called when the user clicks the upload button
   */
  private onSubmitFiles = async (): Promise<void> => {
    let errorMessage: string = null;
    let completedFiles = 0;
    this.isUploading = true;
    this.uploadMessage = null;
    this.operationType = "Creating Notes...";

    const filesToUpload = this.selectedFiles.filter(
      (fileWithContent) => !fileWithContent.isExisting
    );

    if (filesToUpload.length === 0) {
      this.isUploading = false;
      this.notifyOutputChanged();
      return;
    }

    // Initialize progress tracking
    this.initializeUploadProgress();
    this.notifyOutputChanged();

    try {
      const parentEntityName = this.context.parameters.parentEntityName.raw;
      const recordId = this.parentRecordId;
      if (!parentEntityName || !recordId) {
        this.isUploading = false;
        this.uploadMessage = {
          text: "Parent entity name or record ID is missing. Please check configuration.",
          type: MessageBarType.error,
        };
        this.notifyOutputChanged();
        return;
      }

      const totalFiles = filesToUpload.length;
      const filesWithNotesId: FileWithContent[] = [];

      let i = 0;
      for (const fileWithContent of filesToUpload) {
        const originalIndex = this.selectedFiles.findIndex(
          (f) => f.file.name === fileWithContent.file.name
        );

        try {
          // Update status to uploading
          this.updateFileProgress(originalIndex, 0, "uploading");

          const file = fileWithContent.file;

          // Simulate progress during file reading
          this.updateFileProgress(originalIndex, 25, "uploading");

          // Always use DataURL based base64 reading for consistency and memory safety
          const fileContent = await readFileAsDataURL(file);

          // Progress after file processing
          this.updateFileProgress(originalIndex, 50, "uploading");

          const fileExtension = file.name.split(".").pop()?.toLowerCase() || "";
          const mimeType = inferMimeTypeFromFileName(fileExtension);

          // Progress before upload
          this.updateFileProgress(originalIndex, 75, "uploading");

          const notesId = await this.dataverseService.createNoteWithAttachment(
            file,
            fileContent,
            recordId,
            parentEntityName,
            mimeType
          );

          // Complete this file
          this.updateFileProgress(originalIndex, 100, "completed");

          fileWithContent.notesId = notesId;
          fileWithContent.file.notesId = notesId; // attach for preview
          filesWithNotesId.push(fileWithContent);
          completedFiles++;
        } catch (fileError) {
          console.error(
            `Error uploading ${fileWithContent.file.name}:`,
            fileError
          );

          // Mark this file as error
          this.updateFileProgress(
            originalIndex,
            0,
            "error",
            `Error uploading: ${(fileError as Error).message}`
          );

          errorMessage = `Error uploading ${fileWithContent.file.name}: ${
            (fileError as Error).message
          }`;
        }

        i++;
      }

      this.filesUploaded = true;
      this.selectedFiles = this.selectedFiles.filter(
        (file) =>
          file.isExisting ||
          filesWithNotesId.some((f) => f.file.name === file.file.name)
      );

      const successfullyUploadedFiles = this.selectedFiles.map(
        (file) => file.file.name
      );
      this.isUploading = false;
      this.parseExistingFieldValue(successfullyUploadedFiles.join(","));

      this.uploadMessage = {
        text: `Successfully created notes for ${completedFiles} of ${totalFiles} files${
          completedFiles < totalFiles
            ? `. Failed files have been removed from the list because ${errorMessage}`
            : ""
        }`,
        type:
          completedFiles === totalFiles && !errorMessage
            ? MessageBarType.success
            : MessageBarType.error,
      };

      // Clear upload progress after a delay
      setTimeout(() => {
        this.uploadProgress = null;
        this.notifyOutputChanged();
      }, 3000);
    } catch (error) {
      console.error("Error in file upload process:", error);
      this.isUploading = false;
      this.uploadMessage = {
        text: `Error in file upload process: ${(error as Error).message}`,
        type: MessageBarType.error,
      };

      // Clear upload progress on error
      this.uploadProgress = null;
    }
    // After all uploads and error handling
    if (
      completedFiles === filesToUpload.length &&
      !errorMessage &&
      this.context &&
      this.context.events &&
      typeof this.context.events.filesUploaded === "function"
    ) {
      //this.triggerSaveRefresh = new Date().toISOString(); // Use timestamp to ensure change is detected
      this.context.events.filesUploaded();
    }
  };

  /**
   * It is called by the framework prior to a control receiving new data.
   */
  public getOutputs(): IOutputs {
    return {
      fileData: this.fileDataValue,
      isUploading: this.isUploading,
      //triggerSaveRefresh: this.triggerSaveRefresh,
    };
  }

  /**
   * Called when the control is to be removed from the DOM tree
   */
  public destroy(): void {
    // Add code to cleanup control if necessary
  }
}
