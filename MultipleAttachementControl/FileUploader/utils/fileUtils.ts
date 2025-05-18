/**
 * Helper functions for file operations
 */

/**
 * Properly converts an ArrayBuffer to a base64 string
 * This is crucial for binary files like PDFs and Office documents
 */
export function arrayBufferToBase64(buffer: ArrayBuffer): string {
  let binary = "";
  const bytes = new Uint8Array(buffer);
  const len = bytes.byteLength;

  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i]);
  }

  return binary;
}

/**
 * Reads a file and returns its contents as a base64 string
 */
export async function readFileAsArrayBuffer(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      if (reader.result instanceof ArrayBuffer) {
        resolve(arrayBufferToBase64(reader.result));
      } else if (typeof reader.result === "string") {
        resolve(reader.result.split(",")[1]); // Remove data URL prefix
      } else {
        reject(new Error("Unexpected result type"));
      }
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

// Read a file and return its contents as a base64 string using reader.readAsDataURL
export async function readFileAsDataURL(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      if (typeof reader.result === "string") {
        resolve(reader.result.split(",")[1]); // Remove data URL prefix
      } else {
        reject(new Error("Unexpected result type"));
      }
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}
