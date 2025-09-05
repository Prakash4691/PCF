import { FileInfo } from "../types/interfaces";
// Cache to assign stable placeholder names for file objects that have no usable name
const _unnamedCache = new WeakMap<object, string>();
let _unnamedCounter = 0;

/**
 * Returns a safe string file name across browsers and Power Apps mobile.
 * Some mobile webviews/polyfills may expose a non-string `name` or use `fileName`.
 */
interface SafeFileNameLike {
  name?: unknown;
  fileName?: unknown;
  Name?: unknown;
}

export const getSafeFileName = (file: SafeFileNameLike | File): string => {
  const n = (file as SafeFileNameLike)?.name;
  const fn = (file as SafeFileNameLike)?.fileName;
  const N = (file as SafeFileNameLike)?.Name;
  if (typeof n === "string" && n.length) return n;
  if (typeof fn === "string" && fn.length) return fn;
  if (typeof N === "string" && N.length) return N;
  const fallback = (file as File)?.name;
  if (typeof fallback === "string" && fallback.length) return fallback;
  if (_unnamedCache.has(file as object)) {
    return _unnamedCache.get(file as object)!;
  }
  const placeholder = `unnamed-file-${++_unnamedCounter}`;
  _unnamedCache.set(file as object, placeholder);
  return placeholder;
};

export const getFileInfo = (
  file: File,
  additionalInfo?: {
    source?: "fileupload" | "timeline";
    subject?: string;
    noteText?: string;
    createdOn?: Date;
    modifiedOn?: Date;
  }
): FileInfo => {
  const safeName = getSafeFileName(file);
  const fileExtension = safeName.split(".").pop()?.toLowerCase() || "";
  const sizeKb = (
    Number((file as unknown as { size?: number }).size) / 1024
  ).toFixed(2); // size in KB
  const sizeMb = (
    Number((file as unknown as { size?: number }).size) /
    (1024 * 1024)
  ).toFixed(2); // size in MB
  const rawType = (file as unknown as { type?: unknown })?.type;
  let fileType: string = typeof rawType === "string" ? rawType : "";
  let icon = "Document";
  // If file size is very small (1 byte), mark as existing file (dummy file)
  const sizeVal = Number((file as unknown as { size?: number }).size);
  const isExistingFile = sizeVal <= 1;

  // Map file types to appropriate icons
  if (
    (fileType && fileType.startsWith("image/")) ||
    ["jpg", "jpeg", "png", "gif", "bmp", "webp", "svg", "ico"].includes(
      fileExtension
    )
  ) {
    icon = "FileImage";
    fileType = "Image";
  } else if (
    (fileType && fileType.includes("pdf")) ||
    fileExtension === "pdf"
  ) {
    icon = "PDF";
    fileType = "PDF Document";
  } else if (
    (fileType && fileType.includes("spreadsheet")) ||
    ["xls", "xlsx", "csv"].includes(fileExtension)
  ) {
    icon = "ExcelDocument";
    fileType = "Spreadsheet";
  } else if (
    (fileType && fileType.includes("word")) ||
    ["doc", "docx", "rtf"].includes(fileExtension)
  ) {
    icon = "WordDocument";
    fileType = "Word Document";
  } else if (
    (fileType && fileType.includes("presentation")) ||
    ["ppt", "pptx"].includes(fileExtension)
  ) {
    icon = "PowerPointDocument";
    fileType = "Presentation";
  } else if (
    (fileType && fileType.includes("text")) ||
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
  } else if (sizeVal < 1024) {
    sizeText = `${sizeVal} B`;
  } else if (sizeVal < 1024 * 1024) {
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
    source: additionalInfo?.source,
    subject: additionalInfo?.subject,
    noteText: additionalInfo?.noteText,
    createdOn: additionalInfo?.createdOn,
    modifiedOn: additionalInfo?.modifiedOn,
  };
};
