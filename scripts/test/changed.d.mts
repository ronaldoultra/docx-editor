export declare function resolveSpecifier(from: string, specifier: string): string | null;
export declare function importers(
  files: string[],
  read?: (file: string) => string
): Map<string, Set<string>>;
export declare function selectTests(
  changed: string[],
  files: string[],
  reverse: Map<string, Set<string>>,
  read: (file: string) => string
): { all: true; reason: string } | { all: false; tests: string[] };
export declare function angularReminder(changed: string[]): string | null;
