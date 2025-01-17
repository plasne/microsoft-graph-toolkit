/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IDBPObjectStore, openDB } from 'idb';
import { Providers } from '../providers/Providers';
import { CacheItem, CacheSchema, Index, dbListKey } from './CacheService';

/**
 * Represents a store in the cache
 *
 * @class CacheStore
 * @template T
 */

export class CacheStore<T extends CacheItem> {
  private readonly schema: CacheSchema;
  private readonly store: string;

  public constructor(schema: CacheSchema, store: string) {
    if (!(store in schema.stores)) {
      throw Error('"store" must be defined in the "schema"');
    }

    this.schema = schema;
    this.store = store;
  }

  /**
   * gets value from cache for the given key
   *
   * @param {string} key
   * @returns {Promise<T>}
   * @memberof Cache
   */
  public async getValue(key: string): Promise<T | null> {
    if (!window.indexedDB) {
      return null;
    }
    try {
      const db = await this.getDb();
      return db.get(this.store, key) as unknown as T;
    } catch (e) {
      return null;
    }
  }

  /**
   * removes a value from the cache for the given key
   *
   * @param {string} key
   * @returns {Promise<void>}
   * @memberof Cache
   */
  public async delete(key: string): Promise<void> {
    if (!window.indexedDB) {
      return;
    }
    try {
      const db = await this.getDb();
      return db.delete(this.store, key);
    } catch (e) {
      return;
    }
  }

  /**
   * inserts value into cache for the given key
   *
   * @param {string} key
   * @param {T} item
   * @returns
   * @memberof Cache
   */
  public async putValue(key: string, item: T) {
    if (!window.indexedDB) {
      return;
    }
    try {
      await (await this.getDb()).put(this.store, { ...item, timeCached: Date.now() }, key);
    } catch (e) {
      return;
    }
  }

  /**
   * Clears the store of all stored values
   *
   * @returns
   * @memberof Cache
   */
  public async clearStore() {
    if (!window.indexedDB) {
      return;
    }
    try {
      await (await this.getDb()).clear(this.store);
    } catch (e) {
      return;
    }
  }

  /**
   * Returns the name of the parent DB that the cache store belongs to
   */
  public async getDBName() {
    const id = await Providers.getCacheId();
    if (id) {
      return `mgt-${this.schema.name}` + `-${id}`;
    }
  }

  private async getDb() {
    const dbName = await this.getDBName();
    if (dbName) {
      return openDB(dbName, this.schema.version, {
        // eslint-disable-next-line @typescript-eslint/no-unused-vars
        upgrade: (db, _oldVersion, _newVersion, transaction) => {
          const dbArray: string[] = (JSON.parse(localStorage.getItem(dbListKey)) as string[]) || [];
          if (!dbArray.includes(dbName)) {
            dbArray.push(dbName);
          }
          localStorage.setItem(dbListKey, JSON.stringify(dbArray));
          for (const storeName in this.schema.stores) {
            if (Object.prototype.hasOwnProperty.call(this.schema.stores, storeName)) {
              const indexes: Index[] = this.schema.indexes?.[storeName] ?? [];
              if (!db.objectStoreNames.contains(storeName)) {
                const objectStore = db.createObjectStore(storeName);
                indexes.forEach(i => {
                  objectStore.createIndex(i.name, i.field);
                });
              } else {
                const store = transaction.objectStore(storeName);
                indexes.forEach(i => {
                  if (store && !store.indexNames.contains(i.name)) {
                    store.createIndex(i.name, i.field);
                  }
                });
              }
            }
          }
        }
      });
    }
  }

  public async queryDb(indexName: string, query: IDBKeyRange | IDBValidKey): Promise<T[]> {
    const db = await this.getDb();
    return (await db.getAllFromIndex(this.store, indexName, query)) as T[];
  }

  /**
   * Helper function to get a wrapping transaction for an action function
   * @param action a function that takes an object store uses it to make changes to the cache
   */
  public async transaction(action: (store: IDBPObjectStore<unknown, [string], string, 'readwrite'>) => Promise<void>) {
    const db = await this.getDb();
    const tx = db.transaction(this.store, 'readwrite');
    const store = tx.objectStore(this.store);
    await action(store);
    await tx.done;
  }
}
