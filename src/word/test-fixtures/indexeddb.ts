import { IDBFactory } from 'fake-indexeddb';
import { serialize, deserialize } from 'v8';
import { randomUUID } from 'crypto';

export function resetIndexedDB() {
  Object.defineProperty(globalThis, 'indexedDB', { configurable: true, writable: true, value: new IDBFactory() });
  Object.defineProperty(globalThis, 'structuredClone', {
    configurable: true, value: (value: unknown) => deserialize(serialize(value)),
  });
  Object.defineProperty(globalThis, 'crypto', { configurable: true, value: { randomUUID } });
}
