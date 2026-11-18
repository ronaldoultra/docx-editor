// Setup for the Angular unit tests (`bun run test:angular`). The builder runs the specs in
// jsdom, which has no canvas. Its `getContext` logs a "not implemented" error on each call.
// A browser without a 2D context returns `null`, and the engine already handles that case,
// so return `null` here too.
HTMLCanvasElement.prototype.getContext = function getContext() {
  return null;
} as typeof HTMLCanvasElement.prototype.getContext;
