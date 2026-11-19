// The claims of packages/vue/test/font-origins.test.ts.
//
// An eager origin holds the document until it settles. An on-demand resolver cannot be waited
// for, so it goes through at once behind one stable identity. In Angular a signal is also a
// function, so `injectFonts` reads a signal and treats any other function as a resolver.
import { computed, signal } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import type { FontSource } from '@docx-editor.dev/core/contracts/editor';
import {
  defineFontResolver,
  isFontResolver,
  type FontResolutionRequest,
} from '@docx-editor.dev/core/editor';
import { configureZoneless } from '../testing/harness';
import { injectDocxSource, injectFonts, type DocxSourceOptions } from './fonts';

const BYTES = new Uint8Array([1, 2, 3]);
const REQUEST: FontResolutionRequest = { families: ['Calibri'], defaultFamily: 'Calibri' };

function source(family: string, id: string): FontSource {
  return {
    request: { family, weight: 400, style: 'normal' },
    id,
    bytes: new Uint8Array([0]),
    hash: `sha256:${id}`,
    faceIndex: 0,
  };
}

function deferred<T>(): { promise: Promise<T>; resolve: (value: T) => void } {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((settle) => {
    resolve = settle;
  });
  return { promise, resolve };
}

/** Run the effects, then let the promises that they started settle. */
async function flush(): Promise<void> {
  TestBed.tick();
  for (let index = 0; index < 4; index += 1) await Promise.resolve();
  TestBed.tick();
}

type Merged = { sources: FontSource[] };
const ids = (merged: unknown) => (merged as Merged).sources.map((entry) => entry.id);

function docxSource(options: DocxSourceOptions | (() => DocxSourceOptions)) {
  return TestBed.runInInjectionContext(() => injectDocxSource(BYTES, options));
}

/** Collect console warnings while `body` runs. */
async function warnings(body: () => Promise<void>): Promise<unknown[][]> {
  const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
  try {
    await body();
    return warn.mock.calls;
  } finally {
    warn.mockRestore();
  }
}

describe('injectDocxSource: the eager path holds the document', () => {
  beforeEach(() => configureZoneless());

  it('keeps the document undefined until a loader with no argument settles', async () => {
    const gate = deferred<Merged>();
    const result = docxSource({ fonts: () => gate.promise });
    await flush();
    expect(result.document()).toBeUndefined();
    expect(result.isLoading()).toBe(true);
    gate.resolve({ sources: [source('Calibri', 'eager')] });
    await flush();
    expect(result.document()).toBe(BYTES);
    expect(result.isLoading()).toBe(false);
  });

  it('releases the document when a loader fails, and reports the failure', async () => {
    let result!: ReturnType<typeof docxSource>;
    const calls = await warnings(async () => {
      result = docxSource({ fonts: () => Promise.reject(new Error('no faces today')) });
      await flush();
    });
    expect(result.document()).toBe(BYTES);
    expect(calls).toHaveLength(1);
    expect(String(calls[0]![0])).toContain('font loading failed');
  });

  it('composes a list of eager origins, first wins', async () => {
    const result = docxSource({
      fonts: [{ sources: [source('Calibri', 'first')] }, { sources: [source('Calibri', 'last')] }],
    });
    await flush();
    expect(ids(result.fonts())).toEqual(['first']);
  });

  it('keeps the other origins when one eager origin rejects, and warns once', async () => {
    let result!: ReturnType<typeof docxSource>;
    const calls = await warnings(async () => {
      result = docxSource({
        fonts: [
          { sources: [source('Calibri', 'before')] },
          () => Promise.reject(new Error('boom')),
          { sources: [source('Cambria', 'after')] },
        ],
      });
      await flush();
    });
    expect(ids(result.fonts())).toEqual(['before', 'after']);
    expect(calls).toHaveLength(1);
  });
});

describe('injectDocxSource: the on-demand path does not wait', () => {
  beforeEach(() => configureZoneless());

  it('gives the document at once, and `fonts` is a resolver', async () => {
    const lazy = defineFontResolver(async () => ({ sources: [source('Calibri', 'lazy')] }));
    const result = docxSource({ fonts: lazy });
    TestBed.tick();
    expect(result.document()).toBe(BYTES);
    expect(result.isLoading()).toBe(false);
    expect(typeof result.fonts()).toBe('function');
  });

  it('keeps one resolver identity across origin changes, and reads the latest origins', async () => {
    const origin = signal(defineFontResolver(async () => ({ sources: [source('Calibri', 'a')] })));
    const result = docxSource(() => ({ fonts: origin() }));
    await flush();
    const first = result.fonts();
    origin.set(defineFontResolver(async () => ({ sources: [source('Calibri', 'b')] })));
    await flush();
    expect(result.fonts()).toBe(first);
    const merged = await (first as (request: FontResolutionRequest) => Promise<unknown>)(REQUEST);
    expect(ids(merged)).toEqual(['b']);
  });

  it('makes the whole list on demand for one resolver in it', async () => {
    const result = docxSource({
      fonts: [
        { sources: [source('Cambria', 'brand')] },
        defineFontResolver(async () => ({ sources: [source('Calibri', 'lazy')] })),
      ],
    });
    await flush();
    expect(result.document()).toBe(BYTES);
    const resolve = result.fonts() as (request: FontResolutionRequest) => Promise<unknown>;
    expect(ids(await resolve(REQUEST))).toEqual(['brand', 'lazy']);
  });

  it('calls a loader in an on-demand list with no argument', async () => {
    let seen: unknown;
    const result = docxSource({
      fonts: [
        (...args: unknown[]) => {
          seen = args;
          return { sources: [source('Cambria', 'loader')] };
        },
        defineFontResolver(async () => ({ sources: [source('Calibri', 'resolver')] })),
      ] as never,
    });
    await flush();
    const resolve = result.fonts() as (request: FontResolutionRequest) => Promise<unknown>;
    expect(ids(await resolve(REQUEST))).toEqual(['loader', 'resolver']);
    expect(seen).toEqual([]);
  });
});

describe('injectDocxSource follows its source', () => {
  beforeEach(() => configureZoneless());

  it('follows a signal of bytes', async () => {
    const other = new Uint8Array([4, 5, 6]);
    const bytes = signal<Uint8Array>(BYTES);
    const result = TestBed.runInInjectionContext(() => injectDocxSource(bytes));
    await flush();
    expect(result.document()).toBe(BYTES);
    bytes.set(other);
    await flush();
    expect(result.document()).toBe(other);
  });

  it('aborts the fetch of an old URL, and ignores its late answer', async () => {
    let answerFirst!: (response: Response) => void;
    let firstSignal: AbortSignal | undefined;
    const fetch = vi.spyOn(globalThis, 'fetch').mockImplementation((input, init) => {
      if (String(input).endsWith('first.docx')) {
        firstSignal = init?.signal ?? undefined;
        return new Promise<Response>((resolve) => {
          answerFirst = resolve;
        });
      }
      return Promise.resolve(new Response(new Uint8Array(BYTES), { status: 200 }));
    });
    try {
      const url = signal('/first.docx');
      const result = TestBed.runInInjectionContext(() => injectDocxSource(url));
      await flush();
      url.set('/second.docx');
      await flush();
      await new Promise((resolve) => setTimeout(resolve, 0));
      await flush();
      expect(firstSignal?.aborted).toBe(true);
      expect(result.document()).toEqual(BYTES);
      answerFirst(new Response(new Uint8Array([9, 9, 9]), { status: 200 }));
      await flush();
      expect(result.document()).toEqual(BYTES);
      expect(result.error()).toBeNull();
    } finally {
      fetch.mockRestore();
    }
  });
});

describe('injectFonts takes each origin in the same shape', () => {
  it('calls a marked resolver once, with the request', async () => {
    let calls = 0;
    const lazy = defineFontResolver(async (request: FontResolutionRequest) => {
      calls += 1;
      return { sources: [source(request.defaultFamily, 'lazy')] };
    });
    expect(ids(await injectFonts(lazy)(REQUEST))).toEqual(['lazy']);
    expect(calls).toBe(1);
  });

  it('composes two resolvers, the packaged-plus-Google shape', async () => {
    const compose = injectFonts(
      defineFontResolver(async () => ({ sources: [source('Calibri', 'packaged')] })),
      defineFontResolver(async () => ({ sources: [source('Montserrat', 'google')] }))
    );
    expect(ids(await compose(REQUEST))).toEqual(['packaged', 'google']);
  });

  it('returns a marked resolver, so it can be an origin of another list', async () => {
    const inner = injectFonts(
      defineFontResolver(async () => ({ sources: [source('Calibri', 'a')] }))
    );
    expect(isFontResolver(inner)).toBe(true);
    expect(ids(await injectFonts(inner)(REQUEST))).toEqual(['a']);
  });

  // Two of these shapes report `length === 0`, so an arity test would read them as getters.
  const unmarked: readonly [string, (request: FontResolutionRequest) => unknown][] = [
    ['one declared argument', (request) => ({ sources: [source(request.defaultFamily, 'r')] })],
    [
      'a defaulted argument, length 0',
      ((request = { families: [], defaultFamily: 'Fallback' }) => ({
        sources: [source(request.defaultFamily, 'r')],
      })) as (request: FontResolutionRequest) => unknown,
    ],
    [
      'rest arguments, length 0',
      ((...args: FontResolutionRequest[]) => ({
        sources: [source(args[0]!.defaultFamily, 'r')],
      })) as (request: FontResolutionRequest) => unknown,
    ],
  ];

  it.each(unmarked)('gives the request to an unmarked resolver: %s', async (_shape, resolve) => {
    const merged = (await injectFonts(resolve as never)(REQUEST)) as Merged;
    expect(merged.sources.map((entry) => entry.request.family)).toEqual(['Calibri']);
  });

  it('reports a function that returns another origin, and does not compose it', async () => {
    const getterOfOrigin = () =>
      defineFontResolver(async () => ({ sources: [source('Calibri', 'inner')] }));
    let merged: unknown = 'unset';
    const calls = await warnings(async () => {
      merged = await injectFonts(getterOfOrigin as never)(REQUEST);
    });
    expect(merged).toBeUndefined();
    expect(calls).toHaveLength(1);
    expect(String(calls[0]![0])).toContain('answered with a function');
  });

  it('reads a computed origin, the way to build an origin lazily', async () => {
    const origin = computed(() =>
      defineFontResolver(async () => ({ sources: [source('Calibri', 'from-computed')] }))
    );
    expect(ids(await injectFonts(origin)(REQUEST))).toEqual(['from-computed']);
  });

  it('composes a plain function that ignores the request', async () => {
    const getter = () => ({ sources: [source('Calibri', 'from-getter')] });
    expect(ids(await injectFonts(getter)(REQUEST))).toEqual(['from-getter']);
  });

  it('reads a signal origin again on each resolve', async () => {
    const fragment = signal<Merged | undefined>({ sources: [source('Calibri', 'brand')] });
    const compose = injectFonts(fragment);
    expect(await compose(REQUEST)).toBeDefined();
    fragment.set(undefined);
    expect(await compose(REQUEST)).toBeUndefined();
  });
});
