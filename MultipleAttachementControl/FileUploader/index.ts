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
  let binary = "";
  const bytes = new Uint8Array(buffer);
  const len = bytes.byteLength;

  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i]);
  }

  return binary;
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
        const mimeType: string = this.inferMimeTypeFromFileName(extension);
        /*  let mimeType = "application/octet-stream";

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
        } */

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
   */
  private onFilesSelected = async (files: File[]): Promise<void> => {
    try {
      // We no longer need to read file content here since we'll use block upload
      const newFiles = files.map((file) => ({
        file,
        isExisting: false,
      }));

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
   * Creates a note record with the file as an attachment using block upload
   * This implements the recommended Dataverse approach for file uploads
   */
  private async createNoteWithAttachment(
    fileWithContent: FileWithContent,
    parentId: string,
    parentEntityName: string
  ): Promise<string> {
    try {
      let fileContent = fileWithContent.content;
      if (!fileContent) {
        console.log(`Reading content for file: ${fileWithContent.file.name}`);
        fileContent = await this.readFileAsBase64(fileWithContent.file);
      }

      // Get file information
      const file = fileWithContent.file;
      const fileName = file.name;
      const fileExtension = fileName.split(".").pop()?.toLowerCase() || "";
      const type = file.type;

      // Set proper MIME type based on extension
      //let mimeType: string;
      const mimeType: string = this.inferMimeTypeFromFileName(fileExtension);
      /* switch (fileExtension) {
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
          mimeType = file.type || "application/octet-stream";
      } */

      // Create the target entity for initialization
      const target = {
        filename: fileName,
        mimetype: mimeType,
        subject: fileName + fileExtension,
        [`objectid_${parentEntityName}@odata.bind`]: `/${parentEntityName}s(${parentId})`,
        isdocument: true,
        documentbody: fileContent,
      };

      const notes = await this.context.webAPI.createRecord(
        "annotation",
        target
      );

      // This code initializes a block upload session for a file attachment in Dataverse.
      // It sends a POST request to the InitializeAnnotationBlocksUpload endpoint, providing metadata about the note (annotation) record.
      // The response will contain a FileContinuationToken, which is required for uploading file blocks in subsequent steps.
      // This call can take a long time to execute if the Dataverse API is slow to respond,
      // or if there are network latency issues. The endpoint initializes a block upload session,
      // which may involve server-side processing and database operations.
      // To help diagnose slowness, you can add timing logs:

      /* const initStart = performance.now();
      const initResponse = await fetch(
        `/api/data/v9.2/InitializeAnnotationBlocksUpload`,
        {
          method: "POST",
          headers: {
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            Accept: "application/json",
            "Content-Type": "application/json; charset=utf-8",
          },
          body: JSON.stringify({
            Target: {
              annotationid: notes.id, // The GUID of the note record just created
              subject: file.name, // Subject for the note
              filename: file.name, // Name of the file being uploaded
              mimetype: file.type, // MIME type of the file
              notetext: "File uploaded in chunks", // Description or note text
              "@odata.type": "Microsoft.Dynamics.CRM.annotation", // OData type for the annotation entity
              [`objectid_${parentEntityName}@odata.bind`]: `/${parentEntityName}s(${parentId})`,
            },
          }),
        }
      );
      const initEnd = performance.now();
      console.log(
        `InitializeAnnotationBlocksUpload took ${(initEnd - initStart).toFixed(
          2
        )} ms`
      );

      // To avoid long-running requests, check for a successful response before parsing JSON
      if (!initResponse.ok) {
        throw new Error(
          `Failed to initialize block upload: ${initResponse.statusText}`
        );
      }
      const initData = await initResponse.json();
      const fileContinuationToken = initData.FileContinuationToken;

      // Step 2: Upload the file in blocks
      const blockIds: string[] = [];
      const blockSize = 4 * 1024 * 1024; // 4 MB
      const fileData = await this.readFileAsArrayBuffer(file);
      const totalBytes = fileData.byteLength;

      // Convert ArrayBuffer to Uint8Array for processing
      const bytes = new Uint8Array(fileData);

      // Calculate how many blocks we need
      const blocksCount = Math.ceil(totalBytes / blockSize);
      console.log(`Uploading ${fileName} in ${blocksCount} blocks of 4MB each`);

      // Upload each block
      for (let i = 0; i < totalBytes; i += blockSize) {
        // Create a block of the appropriate size
        const blockData = bytes.slice(i, Math.min(i + blockSize, totalBytes));

        // Generate a block ID
        const blockId = btoa((Math.random() * 10000000000).toString()); // Convert to base64
        blockIds.push(blockId);

        // Convert Uint8Array to base64
        const blockBase64 = this.uint8ArrayToBase64(blockData);

        // Upload the block
        console.log(
          `Uploading block ${blockIds.length} of ${blocksCount}, size: ${blockData.length} bytes`
        );

        await fetch(`/api/data/v9.2/UploadBlock`, {
          method: "POST",
          headers: {
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            Accept: "application/json",
            "Content-Type": "application/json; charset=utf-8",
          },
          body: JSON.stringify({
            FileContinuationToken: fileContinuationToken,
            BlockId: blockId,
            BlockData: blockData,
          }),
        });
      }

      // Step 3: Commit the upload
      await fetch(`/api/data/v9.2/CommitAnnotationBlocksUpload`, {
        method: "POST",
        headers: {
          "OData-MaxVersion": "4.0",
          "OData-Version": "4.0",
          Accept: "application/json",
          "Content-Type": "application/json; charset=utf-8",
        },
        body: JSON.stringify({
          FileContinuationToken: fileContinuationToken,
          BlockList: blockIds,
          Target: {
            annotationid: notes.id,
            "@odata.type": "Microsoft.Dynamics.CRM.annotation",
            subject: file.name, // Subject for the note
            filename: file.name, // Name of the file being uploaded
            mimetype: file.type, // MIME type of the file
          },
        }),
      }); */

      // Return the ID of the created note
      return notes.id;
    } catch (error) {
      console.error(
        `Error creating note with attachment for ${fileWithContent.file.name}:`,
        error
      );
      throw error;
    }
  }

  private readFileAsBase64(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.readAsArrayBuffer(file);

      reader.onload = () => {
        try {
          if (!reader.result) {
            reject(new Error("FileReader result is null"));
            return;
          }

          const base64 = arrayBufferToBase64(reader.result as ArrayBuffer);

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
    });
  }

  public inferMimeTypeFromFileName(extension: string) {
    const mimeTypes: Record<string, string> = {
      txt: "text/plain",
      html: "text/html",
      css: "text/css",
      js: "text/javascript",
      json: "application/json",
      xml: "application/xml",
      pdf: "application/pdf",
      doc: "application/msword",
      docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      xls: "application/vnd.ms-excel",
      xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      ppt: "application/vnd.ms-powerpoint",
      pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      png: "image/png",
      jpg: "image/jpeg",
      jpeg: "image/jpeg",
      gif: "image/gif",
      bmp: "image/bmp",
      svg: "image/svg+xml",
      mp3: "audio/mpeg",
      mp4: "video/mp4",
      wav: "audio/wav",
      zip: "application/zip",
      csv: "text/csv",
    };

    return mimeTypes[extension] || "application/octet-stream";
  }

  /**
   * Converts Uint8Array to base64 string
   */
  private uint8ArrayToBase64(array: Uint8Array): string {
    let binary = "";
    const chunkSize = 1000; // Process in smaller chunks to avoid call stack issues

    for (let i = 0; i < array.length; i += chunkSize) {
      // Convert each chunk of bytes to string characters
      binary += String.fromCharCode.apply(
        null,
        Array.from(array.subarray(i, i + chunkSize))
      );
    }

    return window.btoa(binary);
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
