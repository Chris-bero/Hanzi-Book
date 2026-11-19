import {
  DictionaryCharRecord,
  ExportRecord,
  HanziCharStrokeData,
  WorkbookProject,
  WorkbookRepository,
} from '../types/workbook';

const DB_NAME = 'hanzi-book-studio-db';
const DB_VERSION = 1;

const STORES = {
  WORKBOOKS: 'workbooks',
  STROKES: 'strokeCache',
  DICTIONARY: 'dictionary',
  EXPORTS: 'exports',
  META: 'meta',
} as const;

class IndexedDBWorkbookRepository implements WorkbookRepository {
  private dbPromise: Promise<IDBDatabase> | null = null;

  private openDB(): Promise<IDBDatabase> {
    if (this.dbPromise) return this.dbPromise;

    this.dbPromise = new Promise((resolve, reject) => {
      const request = indexedDB.open(DB_NAME, DB_VERSION);

      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(STORES.WORKBOOKS)) {
          db.createObjectStore(STORES.WORKBOOKS, { keyPath: 'id' });
        }
        if (!db.objectStoreNames.contains(STORES.STROKES)) {
          db.createObjectStore(STORES.STROKES, { keyPath: 'char' });
        }
        if (!db.objectStoreNames.contains(STORES.DICTIONARY)) {
          db.createObjectStore(STORES.DICTIONARY, { keyPath: 'character' });
        }
        if (!db.objectStoreNames.contains(STORES.EXPORTS)) {
          db.createObjectStore(STORES.EXPORTS, { keyPath: 'id' });
        }
        if (!db.objectStoreNames.contains(STORES.META)) {
          db.createObjectStore(STORES.META, { keyPath: 'key' });
        }
      };

      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });

    return this.dbPromise;
  }

  private async withStore<T>(
    storeName: string,
    mode: IDBTransactionMode,
    run: (store: IDBObjectStore) => IDBRequest<T>
  ): Promise<T> {
    const db = await this.openDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(storeName, mode);
      const store = tx.objectStore(storeName);
      const req = run(store);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  async listWorkbooks(): Promise<WorkbookProject[]> {
    const items = await this.withStore<WorkbookProject[]>(
      STORES.WORKBOOKS,
      'readonly',
      (store) => store.getAll()
    );
    return items.sort((a, b) => b.updatedAt - a.updatedAt);
  }

  async getWorkbook(id: string): Promise<WorkbookProject | null> {
    const result = await this.withStore<WorkbookProject | undefined>(
      STORES.WORKBOOKS,
      'readonly',
      (store) => store.get(id)
    );
    return result ?? null;
  }

  async saveWorkbook(workbook: WorkbookProject): Promise<void> {
    await this.withStore(STORES.WORKBOOKS, 'readwrite', (store) => store.put(workbook));
  }

  async deleteWorkbook(id: string): Promise<void> {
    await this.withStore(STORES.WORKBOOKS, 'readwrite', (store) => store.delete(id));
  }

  async getCachedStrokeData(char: string): Promise<HanziCharStrokeData | null> {
    const result = await this.withStore<HanziCharStrokeData | undefined>(
      STORES.STROKES,
      'readonly',
      (store) => store.get(char)
    );
    return result ?? null;
  }

  async saveCachedStrokeData(data: HanziCharStrokeData): Promise<void> {
    await this.withStore(STORES.STROKES, 'readwrite', (store) => store.put(data));
  }

  async countCachedStrokeChars(): Promise<number> {
    return this.withStore<number>(STORES.STROKES, 'readonly', (store) => store.count());
  }

  async getDictionaryChar(char: string): Promise<DictionaryCharRecord | null> {
    const result = await this.withStore<DictionaryCharRecord | undefined>(
      STORES.DICTIONARY,
      'readonly',
      (store) => store.get(char)
    );
    return result ?? null;
  }

  async bulkSaveDictionaryChars(records: DictionaryCharRecord[]): Promise<void> {
    const db = await this.openDB();
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction(STORES.DICTIONARY, 'readwrite');
      const store = tx.objectStore(STORES.DICTIONARY);
      for (const rec of records) {
        store.put(rec);
      }
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  }

  async getMeta(key: string): Promise<string | null> {
    const row = await this.withStore<{ key: string; value: string } | undefined>(
      STORES.META,
      'readonly',
      (store) => store.get(key)
    );
    return row?.value ?? null;
  }

  async setMeta(key: string, value: string): Promise<void> {
    await this.withStore(STORES.META, 'readwrite', (store) => store.put({ key, value }));
  }

  async deleteMeta(key: string): Promise<void> {
    await this.withStore(STORES.META, 'readwrite', (store) => store.delete(key));
  }

  async listExports(): Promise<ExportRecord[]> {
    const rows = await this.withStore<ExportRecord[]>(
      STORES.EXPORTS,
      'readonly',
      (store) => store.getAll()
    );
    return rows.sort((a, b) => b.exportedAt - a.exportedAt);
  }

  async saveExport(record: ExportRecord): Promise<void> {
    await this.withStore(STORES.EXPORTS, 'readwrite', (store) => store.put(record));
  }

  async deleteExport(id: string): Promise<void> {
    await this.withStore(STORES.EXPORTS, 'readwrite', (store) => store.delete(id));
  }
}

export const repository: WorkbookRepository = new IndexedDBWorkbookRepository();
