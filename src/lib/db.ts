import Dexie, { type EntityTable } from "dexie";

export interface ManifestItem {
  barcode: string;
  name: string;
  expectedQty: number;
  scannedQty: number;
}

export interface ScanEvent {
  id?: number;
  barcode: string;
  name: string;
  matched: boolean;
  scannedAt: number;
}

interface Setting {
  key: string;
  value: string;
}

class BatchScanDB extends Dexie {
  manifestItems!: EntityTable<ManifestItem, "barcode">;
  scanEvents!: EntityTable<ScanEvent, "id">;
  settings!: EntityTable<Setting, "key">;

  constructor() {
    super("batchscan-local");
    this.version(1).stores({
      manifestItems: "&barcode, name",
      scanEvents: "++id, barcode, scannedAt, matched",
      settings: "&key",
    });
  }
}

export const db = new BatchScanDB();
