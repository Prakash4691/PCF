/** Utility helpers for file preview & download */

const objectUrlCache = new Map<string, string>();

export function base64ToBlob(base64: string, mimeType: string): Blob {
  // Remove potential data url prefix
  const cleaned = base64.includes(",") ? base64.split(",").pop() || "" : base64;
  const byteChars = atob(cleaned);
  const byteNumbers = new Array(byteChars.length);
  for (let i = 0; i < byteChars.length; i++) {
    byteNumbers[i] = byteChars.charCodeAt(i);
  }
  const byteArray = new Uint8Array(byteNumbers);
  return new Blob([byteArray], { type: mimeType });
}

export function getObjectUrl(key: string, blob: Blob): string {
  if (objectUrlCache.has(key)) {
    return objectUrlCache.get(key)!;
  }
  const url = URL.createObjectURL(blob);
  objectUrlCache.set(key, url);
  return url;
}

export function revokeObjectUrl(key: string) {
  const existing = objectUrlCache.get(key);
  if (existing) {
    URL.revokeObjectURL(existing);
    objectUrlCache.delete(key);
  }
}

export function clearAllObjectUrls() {
  objectUrlCache.forEach((url) => URL.revokeObjectURL(url));
  objectUrlCache.clear();
}

export function isPreviewable(mimeType: string): boolean {
  return (
    mimeType.startsWith("image/") ||
    mimeType === "application/pdf" ||
    mimeType.startsWith("text/")
  );
}

export function isTextType(mimeType: string): boolean {
  return (
    mimeType.startsWith("text/") ||
    ["application/json", "application/xml", "application/javascript"].includes(
      mimeType
    )
  );
}

export async function readFileAsText(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = reject;
    reader.readAsText(file);
  });
}

export function triggerBrowserDownload(blob: Blob, fileName: string) {
  const url = URL.createObjectURL(blob);
  try {
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  } finally {
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  }
}
