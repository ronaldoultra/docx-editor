// Font origins and document sources for the `fonts` and `document` inputs.
//
// There are two font paths. An eager origin (a configuration, a fragment, a promise, or a loader
// with no argument) can settle before the document opens, so `injectDocxSource` holds the
// document back until it settles. An on-demand resolver needs the parsed document, so nothing
// can wait for it, and `injectDocxSource` gives it through at once behind one stable identity.
import { computed, effect, isSignal, signal, untracked, type Signal } from '@angular/core';
import {
  composeFontConfiguration,
  composeFontOrigins,
  defineFontResolver,
  isFontResolver,
  type FontConfigurationFragment,
  type FontOrigin,
  type FontResolutionRequest,
  type FontResolver,
  type MarkedFontResolver,
} from '@docx-editor.dev/core/editor';
import type { FontConfiguration } from '@docx-editor.dev/core/contracts/editor';

/** A font origin, or a signal of one. @public */
export type FontOriginInput = FontOrigin | FontResolver | Signal<FontOrigin | FontResolver>;

/**
 * Read one origin. A signal gives its current value. Any other function is a resolver and is
 * never called here: a resolver called with no request fails on `request.defaultFamily`.
 */
const readOrigin = (origin: FontOriginInput): FontOrigin =>
  (isSignal(origin) ? origin() : origin) as FontOrigin;

/**
 * Compose font origins into one resolver for the `fonts` input. The first origin wins for a
 * face that two origins supply. Signals are read again for each document, so a changed origin
 * applies on the next resolve. The result is marked, so it can be an origin of another list.
 *
 * To build an origin lazily, pass a `computed`. A plain function is a resolver: it receives
 * the request, and a function that returns another origin is reported, not composed.
 *
 * @example
 * ```ts
 * readonly fonts = injectFonts(packagedFonts(), googleFonts());
 * ```
 *
 * @public
 */
export function injectFonts(...origins: readonly FontOriginInput[]): MarkedFontResolver {
  return defineFontResolver((request: FontResolutionRequest) =>
    composeFontOrigins(origins.map(readOrigin), request)
  );
}

/** The font input that {@link injectDocxSource} composes. @public */
export type DocxFontsInput = FontConfiguration | FontConfigurationFragment;

/** One font origin of {@link DocxSourceOptions}: an origin, or a loader with no argument. @public */
export type DocxFontOrigin = FontOrigin | (() => DocxFontsInput | Promise<DocxFontsInput>);

/** One font origin or a list of them. @public */
export type DocxFontsSource = DocxFontOrigin | readonly DocxFontOrigin[];

/** A document: bytes, or a URL that the host chose. @public */
export type DocxSource = string | URL | Uint8Array | ArrayBuffer;

/** Options of {@link injectDocxSource}. @public */
export interface DocxSourceOptions {
  /** Font origins. Eager origins hold the document until they settle. */
  readonly fonts?: DocxFontsSource;
  /** Options for the `fetch` of a URL source. */
  readonly fetchOptions?: RequestInit;
}

/** What {@link injectDocxSource} returns. Bind `document` and `fonts` to the editor. @public */
export interface DocxSourceState {
  /** The document bytes, or `undefined` while the document or eager fonts load. */
  readonly document: Signal<Uint8Array | undefined>;
  /** The fonts for the `fonts` input. */
  readonly fonts: Signal<FontConfiguration | FontResolver | undefined>;
  /** The error of the last document fetch, or `null`. */
  readonly error: Signal<Error | null>;
  /** True while the document or eager fonts load. */
  readonly isLoading: Signal<boolean>;
}

function bytesOf(source: DocxSource): Uint8Array | null {
  if (source instanceof Uint8Array) return source;
  if (source instanceof ArrayBuffer) return new Uint8Array(source);
  return null;
}

function fontOrigins(source: DocxFontsSource): readonly DocxFontOrigin[] {
  return Array.isArray(source) ? source : [source as DocxFontOrigin];
}

/** A loader in an on-demand list stays a loader: it is called with no argument. */
function asFontOrigin(origin: DocxFontOrigin): FontOrigin {
  if (typeof origin !== 'function' || isFontResolver(origin)) return origin as FontOrigin;
  const loader = origin as () => DocxFontsInput | Promise<DocxFontsInput>;
  return defineFontResolver(async () => loader());
}

/** True when any origin resolves for each document, so nothing can be waited for. */
function isOnDemand(source: DocxFontsSource | undefined): boolean {
  return source !== undefined && fontOrigins(source).some((origin) => isFontResolver(origin));
}

/** On the eager path, a loader is called with no argument; other origins are awaited. */
async function resolveEagerOrigin(origin: DocxFontOrigin): Promise<DocxFontsInput | undefined> {
  if (typeof origin !== 'function') return (await origin) as DocxFontsInput | undefined;
  return await (origin as () => DocxFontsInput | Promise<DocxFontsInput>)();
}

const valueOf = <T>(input: T | (() => T)): T =>
  typeof input === 'function' ? (input as () => T)() : input;

/**
 * Load a document and its fonts for the editor inputs. `source` and `options` can be values,
 * signals, or functions that read signals. A URL source is fetched, and a new source cancels
 * the fetch before it. Call it in an injection context.
 *
 * @example
 * ```ts
 * readonly source = injectDocxSource(() => this.url(), { fonts: packagedFonts() });
 * ```
 * ```html
 * <docx-editor [document]="source.document()" [fonts]="source.fonts()" />
 * ```
 *
 * @public
 */
export function injectDocxSource(
  source: DocxSource | null | undefined | (() => DocxSource | null | undefined),
  options: DocxSourceOptions | (() => DocxSourceOptions) = {}
): DocxSourceState {
  const bytes = signal<Uint8Array | undefined>(undefined);
  const fonts = signal<FontConfiguration | FontResolver | undefined>(undefined);
  const error = signal<Error | null>(null);
  const documentLoading = signal(untracked(() => valueOf(source)) != null);
  const fontsSettled = signal(untracked(() => valueOf(options).fonts) === undefined);

  // One resolver for the life of the caller, which reads the current options at resolve time.
  // The editor replaces its fonts when the `fonts` identity changes, so a stable identity keeps
  // an inline origin from resolving again.
  const resolver: FontResolver = defineFontResolver((request: FontResolutionRequest) =>
    composeFontOrigins(
      fontOrigins(untracked(() => valueOf(options).fonts) ?? []).map(asFontOrigin),
      request
    )
  );

  effect((onCleanup) => {
    const fontsSource = valueOf(options).fonts;
    untracked(() => {
      if (fontsSource === undefined) {
        fonts.set(undefined);
        fontsSettled.set(true);
        return;
      }
      if (isOnDemand(fontsSource)) {
        fonts.set(resolver);
        fontsSettled.set(true);
        return;
      }
      let live = true;
      onCleanup(() => {
        live = false;
      });
      fonts.set(undefined);
      fontsSettled.set(false);
      void loadEagerFonts(fontsSource).then((composed) => {
        if (!live) return;
        if (composed) fonts.set(composed);
        fontsSettled.set(true);
      });
    });
  });

  let generation = 0;
  effect((onCleanup) => {
    const next = valueOf(source);
    untracked(() => {
      generation += 1;
      const current = generation;
      error.set(null);
      if (next == null) {
        bytes.set(undefined);
        documentLoading.set(false);
        return;
      }
      const immediate = bytesOf(next);
      if (immediate) {
        bytes.set(immediate);
        documentLoading.set(false);
        return;
      }
      const controller = new AbortController();
      onCleanup(() => controller.abort());
      bytes.set(undefined);
      documentLoading.set(true);
      // `bytesOf` took the byte sources, so only a URL is left.
      void fetchDocument(
        next as string | URL,
        valueOf(options).fetchOptions,
        controller.signal
      ).then(
        (loaded) => {
          if (current !== generation) return;
          bytes.set(loaded);
          documentLoading.set(false);
        },
        (cause: unknown) => {
          if (current !== generation) return;
          if (cause instanceof Error && cause.name === 'AbortError') return;
          bytes.set(undefined);
          error.set(cause instanceof Error ? cause : new Error('could not open the document'));
          documentLoading.set(false);
        }
      );
    });
  });

  return {
    document: computed(() => (fontsSettled() ? bytes() : undefined)),
    fonts: fonts.asReadonly(),
    error: error.asReadonly(),
    isLoading: computed(() => documentLoading() || !fontsSettled()),
  };
}

/**
 * Settle every eager origin and compose the answers, first wins. One failing origin does not
 * discard the others, and a failure is reported, never fatal.
 */
async function loadEagerFonts(source: DocxFontsSource): Promise<FontConfiguration | undefined> {
  try {
    const settled = await Promise.allSettled(fontOrigins(source).map(resolveEagerOrigin));
    const resolved: DocxFontsInput[] = [];
    for (const outcome of settled) {
      if (outcome.status === 'rejected') {
        console.warn(
          '[fonts] font loading failed for one origin; the others still compose',
          outcome.reason
        );
        continue;
      }
      if (outcome.value !== undefined) resolved.push(outcome.value);
    }
    if (resolved.length === 0) return undefined;
    return composeFontConfiguration(
      resolved[0] as FontConfigurationFragment,
      ...(resolved.slice(1) as readonly FontConfigurationFragment[])
    );
  } catch (cause) {
    console.warn('[fonts] font loading failed; the fixed measurer stays in effect', cause);
    return undefined;
  }
}

async function fetchDocument(
  source: string | URL,
  init: RequestInit | undefined,
  signal: AbortSignal
): Promise<Uint8Array> {
  const response = await fetch(String(source), { ...init, signal });
  if (!response.ok) {
    throw new Error(`could not open the document: ${response.status} ${response.statusText}`);
  }
  return new Uint8Array(await response.arrayBuffer());
}
