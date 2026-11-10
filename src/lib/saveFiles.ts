export interface OutFile {
  name: string;
  data: string | Uint8Array;
  type: string;
}

type DirectoryPicker = (opts?: { mode?: "readwrite" }) => Promise<FileSystemDirectoryHandle>;

export type SaveResult = { saved: "folder" | "downloads"; count: number } | { saved: "cancelled" };

/**
 * Salva i file in una cartella scelta dall'utente (File System Access API).
 * Se il pannello non la supporta, scarica i file uno per uno nella cartella Download.
 */
export async function saveFiles(files: OutFile[]): Promise<SaveResult> {
  const pick = (window as unknown as { showDirectoryPicker?: DirectoryPicker }).showDirectoryPicker;
  if (pick) {
    let dir: FileSystemDirectoryHandle;
    try {
      dir = await pick.call(window, { mode: "readwrite" });
    } catch (e) {
      if (e instanceof DOMException && e.name === "AbortError") return { saved: "cancelled" };
      throw e;
    }
    for (const f of files) {
      const handle = await dir.getFileHandle(f.name, { create: true });
      const w = await handle.createWritable();
      await w.write(f.data as FileSystemWriteChunkType);
      await w.close();
    }
    return { saved: "folder", count: files.length };
  }
  for (const f of files) {
    const url = URL.createObjectURL(new Blob([f.data as BlobPart], { type: f.type }));
    const a = document.createElement("a");
    a.href = url;
    a.download = f.name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10_000);
    await new Promise((r) => setTimeout(r, 250));
  }
  return { saved: "downloads", count: files.length };
}

export function base64ToBytes(b64: string): Uint8Array {
  const bin = atob(b64);
  return Uint8Array.from(bin, (c) => c.charCodeAt(0));
}
