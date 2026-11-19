import {
  Directive,
  Injectable,
  computed,
  inject,
  input,
  signal,
  type Provider,
  type Signal,
} from '@angular/core';
import {
  catalogDirection,
  chromeDirection,
  createT,
  deepMerge,
  en,
  type LocaleStrings,
  type TFunction,
  type Translations,
} from '@docx-editor.dev/i18n';

/** A writing direction of the editor chrome. @public */
export type DocxEditorDirection = 'ltr' | 'rtl';

/**
 * The active translations of the editor chrome. Each provider merges its catalog over the
 * catalog of the provider above it, so a partial catalog falls back to the outer one, then to
 * English.
 *
 * Provide it with `provideDocxEditorI18n()` for fixed translations, or with the
 * `docxEditorI18n` directive for translations that change. Read it with `injectTranslation()`.
 *
 * @public
 */
@Injectable()
export class DocxEditorI18n {
  private readonly parent = inject(DocxEditorI18n, { optional: true, skipSelf: true });
  private readonly source = signal<{
    readonly i18n: () => Translations | undefined;
    readonly dir: () => DocxEditorDirection | undefined;
  }>({ i18n: () => undefined, dir: () => undefined });

  /**
   * The direction that a provider decided for the chrome: its `dir`, the catalog's `_dir`, or a
   * right-to-left `_lang`. `undefined` when no provider decided, so the chrome keeps the
   * direction of the host page.
   */
  readonly chromeDir: Signal<DocxEditorDirection | undefined> = computed(() =>
    chromeDirection(this.source().dir(), this.source().i18n(), this.parent?.chromeDir())
  );

  /** The direction of the translations: the decided direction, otherwise `'ltr'`. */
  readonly dir: Signal<DocxEditorDirection> = computed(
    () => this.chromeDir() ?? catalogDirection(this.source().i18n()) ?? this.parent?.dir() ?? 'ltr'
  );

  /** The language tag of the catalog (`_lang`), `en` by default. */
  readonly lang: Signal<string> = computed(() => {
    const lang = this.source().i18n()?._lang;
    return typeof lang === 'string' ? lang : (this.parent?.lang() ?? 'en');
  });

  /** The merged catalog. */
  readonly catalogue: Signal<LocaleStrings> = computed(
    () =>
      deepMerge(
        (this.parent?.catalogue() ?? en) as Record<string, unknown>,
        this.source().i18n() as Record<string, unknown> | undefined
      ) as LocaleStrings
  );

  /** The translation function for the merged catalog and its language. */
  readonly t: Signal<TFunction> = computed(() => createT(this.catalogue(), this.lang()));

  /**
   * Read the catalog and the direction from these functions. They can read signals, so a
   * change applies without a new provider.
   */
  use(
    i18n: () => Translations | undefined,
    dir: () => DocxEditorDirection | undefined = () => undefined
  ): void {
    this.source.set({ i18n, dir });
  }
}

/** Options for {@link provideDocxEditorI18n}. @public */
export interface DocxEditorI18nOptions {
  /**
   * The writing direction of the chrome. It decides over the catalog's `_dir` and `_lang`.
   * `'ltr'` keeps a left-to-right chrome on a right-to-left host page.
   */
  readonly dir?: DocxEditorDirection;
}

/**
 * Provide fixed translations for the editor chrome below this injector.
 *
 * @example
 * ```ts
 * import { ptBR } from '@docx-editor.dev/i18n';
 *
 * @Component({ providers: [provideDocxEditorI18n(ptBR)], … })
 * ```
 *
 * @public
 */
export function provideDocxEditorI18n(
  i18n?: Translations,
  options: DocxEditorI18nOptions = {}
): Provider[] {
  return [
    {
      provide: DocxEditorI18n,
      useFactory: () => {
        const service = new DocxEditorI18n();
        service.use(
          () => i18n,
          () => options.dir
        );
        return service;
      },
    },
  ];
}

/**
 * Provide translations for the editor chrome inside this element, from inputs that can change.
 *
 * @example
 * ```html
 * <section [docxEditorI18n]="catalog()" docxEditorDir="rtl">
 *   <div docxEditorRoot [document]="bytes()">…</div>
 * </section>
 * ```
 *
 * @public
 */
@Directive({
  selector: '[docxEditorI18n]',
  providers: [DocxEditorI18n],
})
export class DocxEditorI18nDirective {
  /** The catalog. Missing keys fall back to the outer provider, then to English. */
  readonly docxEditorI18n = input<Translations | undefined>(undefined);

  /** The writing direction of the chrome. See {@link DocxEditorI18nOptions.dir}. */
  readonly docxEditorDir = input<DocxEditorDirection | undefined>(undefined);

  constructor() {
    inject(DocxEditorI18n, { self: true }).use(
      () => this.docxEditorI18n(),
      () => this.docxEditorDir()
    );
  }
}

/** What {@link injectTranslation} returns. @public */
export interface DocxEditorTranslation {
  /** The translation function. */
  readonly t: Signal<TFunction>;
  /** The merged catalog. */
  readonly catalogue: Signal<LocaleStrings>;
  /** The language tag of the catalog, `en` by default. */
  readonly lang: Signal<string>;
  /** The direction of the translations. Use `injectChromeDir()` for the `dir` attribute. */
  readonly dir: Signal<DocxEditorDirection>;
}

const DEFAULT_I18N: DocxEditorTranslation = {
  t: signal(createT(en, 'en')).asReadonly(),
  catalogue: signal(en as LocaleStrings).asReadonly(),
  lang: signal('en').asReadonly(),
  dir: signal<DocxEditorDirection>('ltr').asReadonly(),
};

/**
 * The active translations, as signals. Without a provider, the English catalog. Call it in an
 * injection context.
 *
 * @public
 */
export function injectTranslation(): DocxEditorTranslation {
  const service = inject(DocxEditorI18n, { optional: true });
  if (!service) return DEFAULT_I18N;
  return { t: service.t, catalogue: service.catalogue, lang: service.lang, dir: service.dir };
}

const NO_DIRECTION: Signal<DocxEditorDirection | undefined> = signal(undefined).asReadonly();

/**
 * The `dir` attribute for a chrome part: the direction a provider decided, or `undefined`, so
 * that the part keeps the direction of the host page. Call it in an injection context.
 *
 * @public
 */
export function injectChromeDir(): Signal<DocxEditorDirection | undefined> {
  return inject(DocxEditorI18n, { optional: true })?.chromeDir ?? NO_DIRECTION;
}
