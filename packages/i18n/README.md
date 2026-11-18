<p align="center">
  <a href="https://www.docx-editor.dev/">
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/eigenpal/docx-editor/main/.github/assets/readme-logo-dark.svg" />
      <source media="(prefers-color-scheme: light)" srcset="https://raw.githubusercontent.com/eigenpal/docx-editor/main/.github/assets/readme-logo-light.svg" />
      <img src="https://raw.githubusercontent.com/eigenpal/docx-editor/main/.github/assets/readme-logo-light.svg" alt="DOCX Editor by EigenPal" width="320" height="90" />
    </picture>
  </a>
</p>

<p align="center">
  <a href="https://www.npmjs.com/package/@docx-editor.dev/i18n"><img src="https://img.shields.io/npm/v/@docx-editor.dev/i18n.svg?style=flat-square&color=3B5BDB" alt="npm version" /></a>
  <a href="https://www.npmjs.com/package/@docx-editor.dev/i18n"><img src="https://img.shields.io/npm/dm/@docx-editor.dev/i18n.svg?style=flat-square&color=3B5BDB" alt="npm downloads" /></a>
  <a href="https://github.com/eigenpal/docx-editor/blob/main/LICENSE"><img src="https://img.shields.io/badge/license-Apache_2.0-blue.svg?style=flat-square&color=3B5BDB" alt="license" /></a>
  <a href="https://docx-editor.dev/editor"><img src="https://img.shields.io/badge/Live_Demo-3B5BDB?style=flat-square&logo=vercel&logoColor=white" alt="Demo" /></a>
  <a href="https://www.docx-editor.dev/docs"><img src="https://img.shields.io/badge/Docs-3B5BDB?style=flat-square&logo=readthedocs&logoColor=white" alt="Documentation" /></a>
</p>

# @docx-editor.dev/i18n

Translate [docx-editor.dev](https://docx-editor.dev) controls with locale strings, types, and runtime helpers. The package provides 15 languages and falls back to English for missing translations.

## Quick start

Install the locale package:

```bash
npm install @docx-editor.dev/i18n
```

Pass a locale to the editor. This example requires the React adapter and its engine peer:

```tsx
import { DocxEditor } from '@docx-editor.dev/react';
import { ptBR } from '@docx-editor.dev/i18n';

<DocxEditor document={bytes} i18n={ptBR} locale="pt-BR" />;
```

The `i18n` prop translates controls. Set `locale` separately for regional date input and generated document labels. It defaults to `en-US`.

To share a locale across editors and custom controls, use `LocaleProvider`:

```tsx
import { DocxEditor, LocaleProvider } from '@docx-editor.dev/react';
import { ptBR } from '@docx-editor.dev/i18n';

<LocaleProvider i18n={ptBR}>
  <DocxEditor document={bytes} />
</LocaleProvider>;
```

Custom controls read the same catalog through `useTranslation()`.

Mix a community locale with custom overrides:

```ts
import { ptBR } from '@docx-editor.dev/i18n';

const myLocale = {
  ...ptBR,
  formattingBar: { ...ptBR.formattingBar, bold: 'Em negrito' },
};
```

Missing or `null` keys keep the inherited catalog value. Without an outer `LocaleProvider`, they fall back to English.

## Right-to-left interfaces

The shipped catalogs are left to right. A catalog that declares `_dir: 'rtl'` mirrors editor controls. Interface direction does not change the document's paragraph direction.

For custom translations, set `_dir: 'rtl'` in your catalog or pass `dir="rtl"` to `LocaleProvider`. For setup and limits, see [Right-to-left support](https://www.docx-editor.dev/docs/2.x/guides/right-to-left).

## Available locales

| Code    | Export | Language            |
| ------- | ------ | ------------------- |
| `en`    | `en`   | English (source)    |
| `pt-BR` | `ptBR` | Portuguese (Brazil) |

The BCP-47 code `pt-BR` uses the camelCase JS identifier `ptBR`. For runtime lookup by tag:

```tsx
import { locales } from '@docx-editor.dev/i18n';

<LocaleProvider i18n={locales[userPreferredLocale]}>
  <DocxEditor document={bytes} />
</LocaleProvider>;
```

Importing `locales` includes every locale in your bundle. Import individual locales to reduce the bundle size.

## Per-locale subpaths

If you choose a locale at runtime, import its subpath to load only that locale. Static imports include it in the bundle. Dynamic imports let the bundler create a separate chunk:

```ts
// Static import: include only this locale's strings.
import ptBR from '@docx-editor.dev/i18n/pt-BR';
```

For on-demand loading, use a dynamic import instead:

```ts
const ptBR = (await import('@docx-editor.dev/i18n/pt-BR')).default;
```

Subpaths ship for every locale: `/en` and `/pt-BR`. Each also exports its locale as a named binding (`import { ptBR } from '@docx-editor.dev/i18n/pt-BR'`) for callers that prefer non-default imports.

## Types

Import types to describe locale data and translation functions:

```ts
import type {
  LocaleStrings, // shape of `en`, the full source of truth
  PartialLocaleStrings, // shape of a community partial (null falls back)
  Translations, // alias for PartialLocaleStrings
  TranslationKey, // 'formattingBar.bold' | 'navigation.find.counter' | ...
  LocaleCode, // 'en' | 'pt-BR'
  TFunction, // signature of the `t()` callback
} from '@docx-editor.dev/i18n';
```

## Outside the editor adapters

Build a typed `t()` outside the adapter packages:

```ts
import { createT, deepMerge, en, ptBR, type LocaleStrings } from '@docx-editor.dev/i18n';

const merged = deepMerge(en, ptBR) as LocaleStrings;
const t = createT(merged, 'pt-BR');
t('formattingBar.bold'); // 'Negrito'
t('navigation.find.total', { total: 15 }); // ICU plurals
```

`en.json` is the source of truth. Add keys there, then run `bun run i18n:fix` from the repo root to sync community locales (new keys land as `null`). Full guide: [docs/i18n.md](https://github.com/eigenpal/docx-editor/blob/main/docs/i18n.md).

## Contributing

The catalogs contain editor controls. Examples keep English text in `examples/shared/example-text.ts`.

If you use `collaborationDemo`, `documentRefresh`, `anchorNavigation`, or `writerAgent` keys, move those messages into your application's text catalog. Also move `toolbar.exportPdfServerUnavailable`. These keys are removed from `LocaleStrings` and `TranslationKey`.

## Commercial support

For commercial support or custom features, [email the support team](mailto:docx-editor@eigenpal.com).
