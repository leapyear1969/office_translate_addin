const DATABASE = 'word-translation-backups';
const STORE = 'documents';

function openDatabase(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    if (typeof indexedDB === 'undefined') {
      reject(new Error('当前浏览器不支持 IndexedDB，无法保存原文备份。'));
      return;
    }
    const request = indexedDB.open(DATABASE, 1);
    let failed = false;
    const timer = setTimeout(() => {
      failed = true;
      reject(new Error('打开本地备份数据库超时，请关闭其他插件窗口后重试。'));
    }, 10000);
    request.onupgradeneeded = () => {
      if (!request.result.objectStoreNames.contains(STORE)) request.result.createObjectStore(STORE);
    };
    request.onerror = () => { clearTimeout(timer); reject(request.error); };
    request.onblocked = () => {
      clearTimeout(timer);
      failed = true;
      reject(new Error('本地备份数据库被占用，请关闭其他插件窗口后重试。'));
    };
    request.onsuccess = () => {
      clearTimeout(timer);
      if (failed) request.result.close();
      else {
        request.result.onversionchange = () => request.result.close();
        resolve(request.result);
      }
    };
  });
}

/** Resolve only after commit, never on the individual put request's success.
 * Read/merge/write in one transaction prevents concurrent tabs losing records. */
export async function accessLocalBackup<T>(id: string, update?: (current: T | undefined) => T): Promise<T | undefined> {
  const db = await openDatabase();
  try {
    return await new Promise<T | undefined>((resolve, reject) => {
      const transaction = db.transaction(STORE, update ? 'readwrite' : 'readonly');
      const store = transaction.objectStore(STORE);
      let value: T | undefined;
      let failure: unknown;
      transaction.oncomplete = () => resolve(value);
      transaction.onabort = () => reject(failure || transaction.error || new Error('本地备份事务已取消。'));
      transaction.onerror = () => { /* Abort is the final failure signal. */ };
      const request = store.get(id);
      request.onsuccess = () => {
        try {
          value = request.result;
          if (update) { value = update(value); store.put(value, id); }
        } catch (error) { failure = error; transaction.abort(); }
      };
    });
  } finally { db.close(); }
}
