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

  // Default
  return "application/octet-stream";
}

export const getMimeTypeFromExtension = (fileName: string): string => {
  const extension = fileName.split(".").pop()?.toLowerCase() || "";
  return inferMimeTypeFromFileName(extension);
};
