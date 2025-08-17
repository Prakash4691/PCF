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
  // Properly base64 encode the binary string
  return btoa(binary);
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

/**
 * Decode base64 (no data URL prefix) into bytes. Useful for context.device.pickFile output
 * when we need to reconstruct a File without inadvertent UTF-16 expansion that corrupts binaries.
 */
export function base64ToUint8Array(base64: string): Uint8Array {
  const cleaned = base64.includes(",") ? base64.split(",").pop() || "" : base64;
  const binary = atob(cleaned);
  const len = binary.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}
