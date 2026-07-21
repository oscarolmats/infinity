// Håller koll på fristående IFC-filer valda en och en (oavsett vilken mapp de
// ligger i) via webbläsarens File System Access API, så att appen kan
// återuppta samma filer nästa gång den öppnas istället för att de behöver
// väljas om.

export interface SyncedFile {
  /** Stabilt id (oberoende av filnamn, som kan krocka mellan olika mappar). */
  id: string;
  handle: FileSystemFileHandle;
  name: string;
  lastModified: number;
  size: number;
}

export async function readFileMeta(
  handle: FileSystemFileHandle,
): Promise<{ lastModified: number; size: number }> {
  const file = await handle.getFile();
  return { lastModified: file.lastModified, size: file.size };
}

// IndexedDB-lagring av källfilernas handtag, så de återupptas nästa gång
// appen öppnas istället för att väljas om.
const DB_NAME = "ifcinfinity-file-sync";
const STORE_NAME = "handles";
const SOURCE_FILES_KEY = "source-files";

interface StoredSourceFile {
  id: string;
  handle: FileSystemFileHandle;
}

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = () => {
      request.result.createObjectStore(STORE_NAME);
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

/**
 * Sparar/läser handtag är enbart en bekvämlighet för nästa session - om
 * IndexedDB är otillgängligt (t.ex. privat läge) ska den pågående sessionens
 * synk fortsätta fungera ändå, så alla fel fångas och loggas istället för
 * att kastas vidare.
 */
export async function saveSourceFiles(files: StoredSourceFile[]): Promise<void> {
  try {
    const db = await openDb();
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, "readwrite");
      tx.objectStore(STORE_NAME).put(files, SOURCE_FILES_KEY);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
    db.close();
  } catch (error) {
    console.warn("Kunde inte spara källfilerna för nästa session.", error);
  }
}

export async function loadSourceFiles(): Promise<StoredSourceFile[]> {
  try {
    const db = await openDb();
    const files = await new Promise<StoredSourceFile[]>((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, "readonly");
      const request = tx.objectStore(STORE_NAME).get(SOURCE_FILES_KEY);
      request.onsuccess = () => resolve((request.result as StoredSourceFile[]) ?? []);
      request.onerror = () => reject(request.error);
    });
    db.close();
    return files;
  } catch (error) {
    console.warn("Kunde inte läsa tidigare sparade källfiler.", error);
    return [];
  }
}
