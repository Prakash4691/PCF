import { FileInfo } from "../types/interfaces";

export const getFileInfo = (file: File): FileInfo => {
  const fileExtension = file.name.split(".").pop()?.toLowerCase() || "";
  const sizeKb = (file.size / 1024).toFixed(2); // size in KB
  const sizeMb = (file.size / (1024 * 1024)).toFixed(2); // size in MB
  let fileType = file.type;
  let icon = "Document";
  // If file size is very small (1 byte), mark as existing file (dummy file)
  const isExistingFile = file.size <= 1;

  // Map file types to appropriate icons
  if (
    fileType.startsWith("image/") ||
    ["jpg", "jpeg", "png", "gif", "bmp", "webp", "svg", "ico"].includes(
      fileExtension
    )
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
    sizeText = `${sizeKb} KB`;
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
