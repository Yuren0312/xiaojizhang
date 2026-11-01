import { openDB, type DBSchema, type IDBPDatabase } from 'idb';

interface LobsterBooksDB extends DBSchema {
  files: {
    key: string;
    value: {
      id: string;
      data: Blob;
      name: string;
      type: string;
      size: number;
      createdAt: string;
    };
  };
  results: {
    key: string;
    value: {
      id: string;
      data: Blob;
      name: string;
      type: string;
      size: number;
      createdAt: string;
    };
  };
}

let dbInstance: IDBPDatabase<LobsterBooksDB> | null = null;

async function getDB(): Promise<IDBPDatabase<LobsterBooksDB>> {
  if (dbInstance) return dbInstance;
  dbInstance = await openDB<LobsterBooksDB>('lobster-books-db', 1, {
    upgrade(db) {
      if (!db.objectStoreNames.contains('files')) {
        db.createObjectStore('files', { keyPath: 'id' });
      }
      if (!db.objectStoreNames.contains('results')) {
        db.createObjectStore('results', { keyPath: 'id' });
      }
    },
  });
  return dbInstance;
}

export function isIndexedDBAvailable(): boolean {
  try {
    return !!window.indexedDB;
  } catch {
    return false;
  }
}

export async function storeFile(
  id: string,
  file: File | Blob,
  category: 'files' | 'results' = 'files'
): Promise<void> {
  const db = await getDB();
  await db.put(category, {
    id,
    data: file,
    name: file instanceof File ? file.name : 'result.xlsx',
    type: file.type,
    size: file.size,
    createdAt: new Date().toISOString(),
  });
}

export async function getFile(
  id: string,
  category: 'files' | 'results' = 'files'
): Promise<{ id: string; data: Blob; name: string; type: string; size: number; createdAt: string } | undefined> {
  const db = await getDB();
  return db.get(category, id);
}

export async function deleteFile(
  id: string,
  category: 'files' | 'results' = 'files'
): Promise<void> {
  const db = await getDB();
  await db.delete(category, id);
}

export async function getAllFileKeys(
  category: 'files' | 'results' = 'files'
): Promise<string[]> {
  const db = await getDB();
  return db.getAllKeys(category);
}

export async function clearAllFiles(
  category: 'files' | 'results' = 'files'
): Promise<void> {
  const db = await getDB();
  await db.clear(category);
}

export function generateFileKey(period: string, module: string, fileName: string): string {
  const timestamp = Date.now();
  const random = Math.random().toString(36).substring(2, 8);
  return `${period}/${module}/${timestamp}-${random}-${fileName}`;
}
