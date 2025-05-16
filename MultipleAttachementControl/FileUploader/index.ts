import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import { FileUploaderComponent } from "./FileUploaderComponent";
import { MessageBarType } from "@fluentui/react/lib/MessageBar";

interface FileWithContent {
  file: File;
  content?: string;
  isExisting: boolean;
  notesId?: string;
}

/**
 * Properly converts an ArrayBuffer to a base64 string
 * This is crucial for binary files like PDFs and Office documents
 */
function arrayBufferToBase64(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  // Process in chunks (e.g. 1000 bytes at a time) to avoid call-stack or performance issues
  const chunkSize = 1000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    // Convert each chunk of bytes into characters
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
  }
  // Encode the combined binary string as base64
  return window.btoa(binary);
}

export class MultipleFileUploader
  implements ComponentFramework.ReactControl<IInputs, IOutputs>
{
  private notifyOutputChanged: () => void;
  private selectedFiles: FileWithContent[] = [];
  private fileDataValue = "";
  private existingFileNames: string[] = [];
  private isUploading = false;
  private uploadMessage: { text: string; type: MessageBarType } | null = null;
  private filesUploaded = false;
  private context: ComponentFramework.Context<IInputs>;
  private parentRecordId: string;
  private operationType: string;
  private blockedFileExtension = "";
  private maxFileSizeForAttachment: number;
  private showDialog: { title: string; subText: string } | null = null;

  /**
   * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
   * Data-set values are not initialized here, use updateView.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
   * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
   * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
   */
  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary
  ): void {
    this.notifyOutputChanged = notifyOutputChanged;
    this.context = context;

    // Parse existing field value if available
    if (context.parameters.fileData.raw) {
      this.parseExistingFieldValue(context.parameters.fileData.raw);
    }

    this.parentRecordId = context.parameters.parentRecordId.raw;
    this.blockedFileExtension =
      context.parameters.blockedFileExtension.raw ?? "";
    this.maxFileSizeForAttachment = context.parameters.maxFileSizeForAttachment
      .raw
      ? parseInt(context.parameters.maxFileSizeForAttachment.raw) * 1024
      : 0;
  }

  /**
   * Parse existing field value into file names array
   * @param fieldValue The raw field value from Dataverse
   */
  private parseExistingFieldValue(fieldValue: string): void {
    if (!fieldValue) return;

    // Split the comma-separated values and clean them up
    this.existingFileNames = fieldValue
      .split(",")
      .map((name) => name.trim())
      .filter((name) => name.length > 0);

    // Create dummy File objects for existing files
    this.createDummyFiles();
  }

  /**
   * Create dummy File objects for existing file names
   */
  private createDummyFiles(): void {
    if (this.existingFileNames.length === 0) return;

    // For each file name, create a dummy File object with appropriate metadata
    const dummyFiles: FileWithContent[] = this.existingFileNames.map(
      (fileName) => {
        // Get file extension to determine mime type
        const extension = fileName.split(".").pop()?.toLowerCase() || "";
        let mimeType = "application/octet-stream";

        // Map common extensions to mime types
        if (["jpg", "jpeg", "png", "gif", "bmp", "webp"].includes(extension)) {
          mimeType = `image/${extension === "jpg" ? "jpeg" : extension}`;
        } else if (extension === "pdf") {
          mimeType = "application/pdf";
        } else if (["doc", "docx"].includes(extension)) {
          mimeType = "application/msword";
        } else if (["xls", "xlsx"].includes(extension)) {
          mimeType = "application/vnd.ms-excel";
        } else if (["ppt", "pptx"].includes(extension)) {
          mimeType = "application/vnd.ms-powerpoint";
        } else if (extension === "txt") {
          mimeType = "text/plain";
        }

        // Create a small placeholder file with the given name and mime type
        const file = new File([new ArrayBuffer(1)], fileName, {
          type: mimeType,
        });
        const notesId =
          this.selectedFiles.find((f) => f.file.name === fileName)?.notesId ||
          "";
        return { file, isExisting: true, notesId };
      }
    );

    // Add the dummy files to the selectedFiles array
    this.selectedFiles = dummyFiles;
    this.updateFileDataValue();
  }

  /**
   * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
   */
  public updateView(
    context: ComponentFramework.Context<IInputs>
  ): React.ReactElement {
    // Check if the field value has changed externally and needs to be reparsed
    if (
      context.parameters.fileData.raw !== this.fileDataValue &&
      context.parameters.fileData.raw
    ) {
      this.parseExistingFieldValue(context.parameters.fileData.raw);
    }

    // Update the context
    this.context = context;

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
      operationType: this.operationType,
      maxFileSizeForAttachment: this.maxFileSizeForAttachment,
      blockedFileExtension: this.blockedFileExtension,
      showDialog: this.showDialog,
    });
  }

  /**
   * Called when the user selects files
   * Properly reads files as ArrayBuffer and converts to base64
   */
  private onFilesSelected = async (files: File[]): Promise<void> => {
    try {
      // Process each file to get its base64 content
      const newFiles = await Promise.all(
        files.map(async (file) => {
          try {
            // Read the file as base64 using our helper method
            const content = await this.readFileAsBase64(file);
            return { file, content, isExisting: false };
          } catch (error) {
            console.error(`Error processing file ${file.name}:`, error);
            // Return the file without content - will be read again later if needed
            return { file, isExisting: false };
          }
        })
      );

      // Add the new files to our collection
      this.selectedFiles = [...this.selectedFiles, ...newFiles];
      this.updateFileDataValue();
      this.filesUploaded = false;
      this.notifyOutputChanged();
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
   * @param index The index of the file to remove
   */
  private onFileRemoved = async (index: number): Promise<void> => {
    // Check if this is an existing file
    const removedFile = this.selectedFiles[index];
    if (removedFile.isExisting) {
      this.uploadMessage = null; // Clear any previous messages
      this.operationType = "Deleting Notes...";
      this.isUploading = true;
      this.notifyOutputChanged();

      try {
        // If it's an existing file, remove it from existingFileNames array
        const fileName = removedFile.file.name;
        this.existingFileNames = this.existingFileNames.filter(
          (name) => name !== fileName
        );

        // Delete record from dataverse for annotation record
        if (removedFile.notesId) {
          await this.context.webAPI.deleteRecord(
            "annotation",
            removedFile.notesId
          );
        }

        // Remove the file from selectedFiles array
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
      // For non-existing files, just remove from the array
      this.selectedFiles = [
        ...this.selectedFiles.slice(0, index),
        ...this.selectedFiles.slice(index + 1),
      ];
      this.updateFileDataValue();
    }

    this.notifyOutputChanged();
  };

  /**
   * Updates the fileDataValue based on selected files
   * This is called whenever files are added or removed
   */
  private updateFileDataValue(): void {
    this.fileDataValue = this.selectedFiles
      .map((fileWithContent) => fileWithContent.file.name)
      .join(", ");
  }

  /**
   * Called when the user clicks the upload button
   * @returns Promise that resolves when the upload is complete
   */
  private onSubmitFiles = async (): Promise<void> => {
    let errorMessage: string = null;
    // Disable button immediately
    this.isUploading = true;
    this.uploadMessage = null;
    this.operationType = "Creating Notes...";
    this.notifyOutputChanged();

    try {
      // Check if we have any files to upload (that aren't just dummy files)
      const filesToUpload = this.selectedFiles.filter(
        (fileWithContent) => !fileWithContent.isExisting
      );
      if (filesToUpload.length === 0) {
        this.isUploading = false;
        this.notifyOutputChanged();
        return;
      }

      // Get the parent record ID and entity name
      const parentEntityName = this.context.parameters.parentEntityName.raw;
      if (!parentEntityName) {
        this.isUploading = false;
        this.notifyOutputChanged();
        return;
      }

      const recordId = this.parentRecordId;
      if (!recordId) {
        this.isUploading = false;
        this.notifyOutputChanged();
        return;
      }

      // Process each file one by one, updating progress as we go
      const totalFiles = filesToUpload.length;
      let completedFiles = 0;
      const filesWithNotesId: FileWithContent[] = [];
      for (const fileWithContent of filesToUpload) {
        try {
          const notesId = await this.createNoteWithAttachment(
            fileWithContent,
            recordId,
            parentEntityName
          );

          fileWithContent.notesId = notesId; // Store the notesId in the fileWithContent object
          filesWithNotesId.push(fileWithContent);
          completedFiles++;
        } catch (fileError) {
          console.error(
            `Error uploading ${fileWithContent.file.name}:`,
            fileError
          );
          errorMessage = `Error uploading ${fileWithContent.file.name}: ${
            (fileError as Error).message
          }`;
        }
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
      this.notifyOutputChanged();
    } catch (error) {
      console.error("Error in file upload process:", error);
      this.isUploading = false;
      this.uploadMessage = {
        text: `Error in file upload process: ${(error as Error).message}`,
        type: MessageBarType.error,
      };
      this.notifyOutputChanged();
    }
  };

  /**
   * Creates a note record with the file as an attachment
   * This is the crucial method for proper file handling
   */
  private async createNoteWithAttachment(
    fileWithContent: FileWithContent,
    parentId: string,
    parentEntityName: string
  ): Promise<string> {
    try {
      // If content is not already available, read it
      let fileContent = fileWithContent.content;
      if (!fileContent) {
        console.log(`Reading content for file: ${fileWithContent.file.name}`);
        fileContent = await this.readFileAsBase64(fileWithContent.file);
      }

      // Validate content
      if (!fileContent || fileContent.length === 0) {
        throw new Error(
          `Failed to get content for file: ${fileWithContent.file.name}`
        );
      }

      // Get the file name and determine MIME type
      const fileName = fileWithContent.file.name;
      const fileExtension = fileName.split(".").pop()?.toLowerCase() || "";

      // Map file extensions to their proper MIME types
      let mimeType: string;
      switch (fileExtension) {
        case "pdf":
          mimeType = "application/pdf";
          break;
        case "doc":
          mimeType = "application/msword";
          break;
        case "docx":
          mimeType =
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
          break;
        case "xls":
          mimeType = "application/vnd.ms-excel";
          break;
        case "xlsx":
          mimeType =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
          break;
        case "ppt":
          mimeType = "application/vnd.ms-powerpoint";
          break;
        case "pptx":
          mimeType =
            "application/vnd.openxmlformats-officedocument.presentationml.presentation";
          break;
        case "txt":
          mimeType = "text/plain";
          break;
        default:
          mimeType = fileWithContent.file.type || "application/octet-stream";
      }

      // Log key information for debugging
      console.log(
        `Creating note with attachment: ${fileName}, MIME: ${mimeType}, Size: ${fileWithContent.file.size} bytes, Base64 length: ${fileContent.length}`
      );

      // Create the annotation entity
      const entity = {
        subject: fileName,
        notetext: `Attachment: ${fileName}`,
        filename: fileName,
        mimetype: mimeType,
        documentbody: fileContent, // The base64 encoded file content
        isdocument: true,
        filesize: fileWithContent.file.size,
        objecttypecode: parentEntityName,
        [`objectid_${parentEntityName}@odata.bind`]: `/${parentEntityName}s(${parentId})`,
      };

      // Create the annotation record
      const notes = await this.context.webAPI.createRecord(
        "annotation",
        entity
      );
      console.log(`Successfully created note with ID: ${notes.id}`);

      return notes.id;
    } catch (error) {
      console.error(
        `Error creating note with attachment for ${fileWithContent.file.name}:`,
        error
      );
      throw error;
    }
  }

  /**
   * Reads a file and returns its content as base64 string
   * This is the core method for handling binary files properly
   */
  private readFileAsBase64(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
      // Log file information for debugging
      console.log(
        `Reading file: ${file.name}, type: ${file.type}, size: ${file.size} bytes`
      );

      const reader = new FileReader();

      reader.onload = () => {
        try {
          if (!reader.result) {
            reject(new Error("FileReader result is null"));
            return;
          }

          const base64 = arrayBufferToBase64(reader.result as ArrayBuffer);

          alert(base64);

          resolve(base64);
        } catch (error) {
          console.error(`Error converting ${file.name} to base64:`, error);
          reject(error);
        }
      };

      reader.onerror = () => {
        console.error(`FileReader error for ${file.name}:`, reader.error);
        reject(reader.error);
      };

      // Always read as ArrayBuffer for binary files
      reader.readAsArrayBuffer(file);
    });
  }

  /**
   * It is called by the framework prior to a control receiving new data.
   * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
   */
  public getOutputs(): IOutputs {
    return {
      fileData: this.fileDataValue,
      isUploading: this.isUploading,
    };
  }

  /**
   * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
   * i.e. cancelling any pending remote calls, removing listeners, etc.
   */
  public destroy(): void {
    // No need to clean up DOM elements as React will handle that
  }
}
