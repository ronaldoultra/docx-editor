// Test files that another runner owns. Bun cannot compile Angular components, so the specs in
// `packages/angular` run with Vitest through the Angular unit-test builder
// (`bun run test:angular`). `bun run test` and `bun run test:changed` skip them.

/** Repository-relative path prefixes whose test files `bun test` does not run. */
export const OTHER_RUNNER_PREFIXES = ['packages/angular/'];

/** Changes that can break the Angular specs, so `test:changed` names `test:angular`. */
const ANGULAR_AFFECTING = ['packages/angular/', 'packages/core/src/editor/'];

/** True when another runner owns `file`. Accepts `\` separators from Windows paths. */
export function ownedByOtherRunner(file) {
  const path = file.replaceAll('\\', '/');
  return OTHER_RUNNER_PREFIXES.some((prefix) => path.startsWith(prefix));
}

/** True when a changed file can affect the Angular specs. */
export function affectsAngularTests(file) {
  const path = file.replaceAll('\\', '/');
  return ANGULAR_AFFECTING.some((prefix) => path.startsWith(prefix));
}
