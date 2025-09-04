/**
 * Infer MIME type from file extension
 */
export function inferMimeTypeFromFileName(extension: string): string {
  // Map common extensions to mime types
  if (["jpg", "jpeg"].includes(extension)) return "image/jpeg";
  if (extension === "png") return "image/png";
  if (extension === "gif") return "image/gif";
  if (extension === "bmp") return "image/bmp";
  if (extension === "webp") return "image/webp";
  if (extension === "svg") return "image/svg+xml";
  if (extension === "ico") return "image/x-icon";
  if (extension === "pdf") return "application/pdf";
  if (extension === "doc") return "application/msword";
  if (extension === "docx")
    return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  if (extension === "xls") return "application/vnd.ms-excel";
  if (extension === "xlsx")
    return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  if (extension === "ppt") return "application/vnd.ms-powerpoint";
  if (extension === "pptx")
    return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
  if (extension === "txt") return "text/plain";
  if (extension === "json") return "application/json";
  if (extension === "xml") return "application/xml";
  if (extension === "js") return "application/javascript";
  if (extension === "ts") return "application/typescript";
  if (extension === "html") return "text/html";
  if (extension === "htm") return "text/html";
  if (extension === "css") return "text/css";
  if (extension === "md") return "text/markdown";
  if (extension === "rtf") return "application/rtf";
  if (extension === "csv") return "text/csv";
  if (["zip", "7z"].includes(extension)) return "application/zip";
  if (extension === "rar") return "application/vnd.rar";
  if (extension === "tar") return "application/x-tar";
  if (extension === "gz") return "application/gzip";
  if (["mp4", "mov", "avi", "wmv", "flv", "webm"].includes(extension)) return "video/mp4";
  if (["mp3", "wav", "ogg", "flac", "m4a"].includes(extension)) return "audio/mpeg";

  // Default
  return "application/octet-stream";
}

export const getMimeTypeFromExtension = (fileName: string): string => {
  const extension = fileName.split(".").pop()?.toLowerCase() || "";
  return inferMimeTypeFromFileName(extension);
};
