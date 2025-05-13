import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import { FileUploaderComponent } from "./FileUploaderComponent";
import { MessageBarType } from "@fluentui/react/lib/MessageBar";

interface FileWithContent {
  file: File;
  content?: string;
  isExisting: boolean;
}

export class MultipleFileUploader
  implements ComponentFramework.ReactControl<IInputs, IOutputs>
{
  private notifyOutputChanged: () => void;
  private selectedFiles: FileWithContent[] = [];
  private fileDataValue = "";
  private existingFileNames: string[] = [];
  private isUploading = false;
  private uploadProgress = 0;
  private uploadMessage: { text: string; type: MessageBarType } | null = null;
  private filesUploaded = false; // New state variable to track if files have been uploaded
  private context: ComponentFramework.Context<IInputs>;
  private parentRecordId: string;

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
        return { file, isExisting: true };
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
      isUploading: this.isUploading || this.filesUploaded, // Consider both states for button disabling
      uploadProgress: this.uploadProgress,
      uploadMessage: this.uploadMessage,
      hasExistingFiles: this.existingFileNames.length > 0,
      filesUploaded: this.filesUploaded,
    });
  }
  private onFilesSelected = async (files: File[]): Promise<void> => {
    const newFiles = await Promise.all(
      files.map(async (file) => ({
        file,
        content: await this.readFileAsBase64(file),
        isExisting: false,
      }))
    );
    this.selectedFiles = [...this.selectedFiles, ...newFiles];
    this.filesUploaded = false; // Reset filesUploaded when new files are added
    this.updateFileDataValue();
    this.notifyOutputChanged();
  };

  private onFileRemoved = (index: number): void => {
    // Check if this is an existing file
    const removedFile = this.selectedFiles[index];
    if (removedFile.isExisting) {
      // If it's an existing file, remove it from existingFileNames array
      const fileName = removedFile.file.name;
      this.existingFileNames = this.existingFileNames.filter(
        (name) => name !== fileName
      );
    }

    this.selectedFiles = [
      ...this.selectedFiles.slice(0, index),
      ...this.selectedFiles.slice(index + 1),
    ];

    this.updateFileDataValue();
    this.notifyOutputChanged();
  };

  private updateFileDataValue(): void {
    this.fileDataValue = this.selectedFiles
      .map((fileWithContent) => fileWithContent.file.name)
      .join(", ");
  }
  private onSubmitFiles = async (): Promise<void> => {
    // Disable button immediately
    this.isUploading = true;
    this.notifyOutputChanged();

    try {
      // Reset other state
      this.uploadProgress = 0;
      this.uploadMessage = null;
      this.notifyOutputChanged();

      // Check if we have any files to upload (that aren't just dummy files)
      const filesToUpload = this.selectedFiles.filter(
        (fileWithContent) => !fileWithContent.isExisting
      );
      if (filesToUpload.length === 0) {
        this.uploadMessage = {
          text: "No files to upload. Please add new files.",
          type: MessageBarType.warning,
        };
        this.isUploading = false;
        this.notifyOutputChanged();
        return;
      }

      // Get the parent record ID and entity name
      const parentEntityName = this.context.parameters.parentEntityName.raw;
      if (!parentEntityName) {
        this.uploadMessage = {
          text: "Parent entity name not specified. Please check control configuration.",
          type: MessageBarType.error,
        };
        this.isUploading = false;
        this.notifyOutputChanged();
        return;
      }

      const recordId = this.parentRecordId;

      if (!recordId) {
        this.uploadMessage = {
          text: "Could not determine parent record ID. Please ensure you're on a saved record.",
          type: MessageBarType.error,
        };
        this.isUploading = false;
        this.notifyOutputChanged();
        return;
      } // Process each file one by one, updating progress as we go
      const totalFiles = filesToUpload.length;
      let completedFiles = 0;
      let successfullyUploadedFiles: string[] = [];
      for (const fileWithContent of filesToUpload) {
        try {
          // First update the message and progress before starting the upload
          this.uploadMessage = {
            text: `Creating note for ${fileWithContent.file.name} (${
              completedFiles + 1
            }/${totalFiles})...`,
            type: MessageBarType.info,
          };
          // Calculate progress for the current file (divide each file upload into two parts)
          this.uploadProgress = completedFiles / totalFiles;
          this.notifyOutputChanged();

          // Create note record with attachment
          await this.createNoteWithAttachment(
            fileWithContent,
            recordId,
            parentEntityName
          );

          // Update message and progress after successful creation
          this.uploadMessage = {
            text: `Successfully created note for ${
              fileWithContent.file.name
            } (${completedFiles + 1}/${totalFiles})`,
            type: MessageBarType.success,
          };
          // Update progress to show this file is complete
          this.uploadProgress = (completedFiles + 1) / totalFiles;
          this.notifyOutputChanged();

          // Add to successful uploads list
          successfullyUploadedFiles.push(fileWithContent.file.name);
          completedFiles++;
        } catch (fileError) {
          console.error(
            `Error uploading file ${fileWithContent.file.name}:`,
            fileError
          );
          this.uploadMessage = {
            text: `Error uploading ${fileWithContent.file.name}: ${
              (fileError as Error).message
            }`,
            type: MessageBarType.error,
          };
          this.notifyOutputChanged();
        }
      } // All done
      this.filesUploaded = true;
      this.uploadProgress = 1;

      // Short delay to show the last file's success message
      await new Promise((resolve) => setTimeout(resolve, 1000));

      // Remove failed files from selectedFiles array
      this.selectedFiles = this.selectedFiles.filter(
        (file) =>
          file.isExisting || // Keep existing files
          successfullyUploadedFiles.includes(file.file.name) // Keep only successfully uploaded new files
      );

      // Update the field value to only include successful files
      this.updateFileDataValue();

      this.uploadMessage = {
        text: `Successfully created notes for ${completedFiles} of ${totalFiles} files${
          completedFiles < totalFiles
            ? `. Failed files have been removed from the list.`
            : ""
        }`,
        type:
          completedFiles === totalFiles
            ? MessageBarType.success
            : MessageBarType.warning,
      };
      this.notifyOutputChanged();
    } catch (error) {
      console.error("Error in file upload process:", error);
      this.isUploading = false;
      this.uploadMessage = {
        text: `Error uploading files: ${(error as Error).message}`,
        type: MessageBarType.error,
      };
      this.notifyOutputChanged();
    }
  };

  /**
   * Creates a note record with the file as an attachment
   * @param fileWithContent The file with content to attach
   * @param parentId The parent record ID
   * @param parentEntityName The parent entity logical name
   */
  private async createNoteWithAttachment(
    fileWithContent: FileWithContent,
    parentId: string,
    parentEntityName: string
  ): Promise<void> {
    try {
      const fileContent =
        fileWithContent.content ||
        (await this.readFileAsBase64(fileWithContent.file)); // Create annotation (note) record
      const entity = {
        subject: fileWithContent.file.name,
        notetext: `Attachment: ${fileWithContent.file.name}`,
        filename: fileWithContent.file.name,
        mimetype: fileWithContent.file.type || "application/octet-stream",
        documentbody: fileContent,
        objecttypecode: parentEntityName,
        [`objectid_${parentEntityName}@odata.bind`]: `/${parentEntityName}s(${parentId})`,
      };

      // Use Web API to create note
      await this.context.webAPI.createRecord("annotation", entity);
    } catch (error) {
      console.error("Error creating note with attachment:", error);
      throw error;
    }
  }

  /**
   * Reads a file and returns its content as base64
   * @param file The file to read
   * @returns Promise with base64 content
   */
  private readFileAsBase64(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = () => {
        if (reader.result) {
          // Extract base64 data (remove data:*/*;base64, prefix)
          const base64Content = reader.result.toString();
          const base64Data = base64Content.split(",")[1] || base64Content;
          resolve(base64Data);
        } else {
          reject(new Error("Failed to read file"));
        }
      };

      reader.onerror = () => {
        reject(reader.error);
      };

      reader.readAsDataURL(file);
    });
  }

  /**
   * It is called by the framework prior to a control receiving new data.
   * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
   */
  public getOutputs(): IOutputs {
    return {
      fileData: this.fileDataValue,
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
