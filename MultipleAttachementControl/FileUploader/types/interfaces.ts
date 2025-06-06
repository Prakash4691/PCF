import { MessageBarType } from "@fluentui/react/lib/MessageBar";
import { IInputs } from "../generated/ManifestTypes";

export interface FileUploaderComponentProps {
  selectedFiles: File[];
  onFilesSelected: (files: File[], isDragAndDrop?: boolean) => void;
  onFileRemoved: (index: number) => void;
  onSubmitFiles: () => void;
  context: ComponentFramework.Context<IInputs>;
  isUploading: boolean;
  uploadMessage?: { text: string; type: MessageBarType } | null;
  hasExistingFiles: boolean;
  filesUploaded: boolean;
  operationType?: string;
  maxFileSizeForAttachment: number;
  blockedFileExtension: string;
  showDialog?: { title: string; subText: string } | null;
}

export interface FileInfo {
  file: File;
  icon: string;
  sizeMb: string;
  sizeText: string;
  fileType: string;
  isExistingFile: boolean;
}

export interface DialogContent {
  title: string;
  subText: string;
}

export interface FileWithContent {
  file: File;
  content?: string;
  isExisting: boolean;
  notesId?: string | Promise<string>;
  isDragAndDrop?: boolean;
}

export interface UploadMessage {
  text: string;
  type: MessageBarType; // Replace with proper MessageBarType when needed
}

export interface FileOperationResult {
  success: boolean;
  message: string;
  notesId?: string;
}
