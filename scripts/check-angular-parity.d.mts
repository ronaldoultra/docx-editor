export declare const LEDGER: string;
export declare const VUE_API: string;
export declare const VUE_TESTS: string;
export declare const ANGULAR_ENTRY: string;
export declare const STATUSES: readonly string[];

export interface ApiReportExport {
  name: string;
  kind: string;
  tags: string[];
}

export interface LedgerRow {
  key: string;
  angular: string;
  status: string;
  notes: string;
}

export interface Ledger {
  exports: Map<string, LedgerRow>;
  tests: Map<string, LedgerRow>;
  duplicates: string[];
}

export declare function parseApiReport(text: string): ApiReportExport[];
export declare function ledgerExports(exports: ApiReportExport[]): ApiReportExport[];
export declare function testFiles(directory: string, root?: string): string[];
export declare function parseLedger(text: string): Ledger;
export declare function checkLedger(input: {
  exports: ApiReportExport[];
  tests: string[];
  ledger: Ledger;
  noTodo?: boolean;
}): string[];
export declare function untaggedAngularExports(entry?: string): string[];
