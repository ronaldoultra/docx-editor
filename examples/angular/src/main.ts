import {
  provideZoneChangeDetection,
  provideZonelessChangeDetection,
  type EnvironmentProviders,
  type Provider,
} from '@angular/core';
import { bootstrapApplication } from '@angular/platform-browser';
import { setHarfBuzzWasmUrl } from '@docx-editor.dev/core/layout';
import { AppComponent } from './app/app.component';

// The Angular builder does not emit `new URL('harfbuzz.wasm', import.meta.url)` as an
// asset. `angular.json` copies the binary to the site root, and this line points the
// text shaper at it. Call it before the first editor is created.
setHarfBuzzWasmUrl(new URL('harfbuzz.wasm', document.baseURI));

// The app starts without zone.js. Add `?zone` to the URL to start it with zone.js, so you
// can check both change detection modes with one build.
const withZone = new URLSearchParams(location.search).has('zone');

async function start() {
  let changeDetection: Provider | EnvironmentProviders;
  if (withZone) {
    await import('zone.js');
    changeDetection = provideZoneChangeDetection({ eventCoalescing: true });
  } else {
    changeDetection = provideZonelessChangeDetection();
  }
  await bootstrapApplication(AppComponent, { providers: [changeDetection] });
}

start().catch((error: unknown) => console.error(error));
