// Task 6.2: the claims of packages/vue/test/hyperlink-as-child.test.ts, and the `javascript:`
// claim of security.test.ts.
import { ChangeDetectionStrategy, Component, Directive, signal, viewChild } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import type { DocxEditorInstance, SurfaceHyperlink } from '@docx-editor.dev/core/editor';
import { sanitizeHref } from '@docx-editor.dev/core/store';
import { strToU8, zipSync } from 'fflate';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import { configureZoneless, engineOf, render, settle } from '../../testing/harness';
import { DocxEditorContentComponent } from '../content.component';
import { DocxEditorRootDirective } from '../root.directive';
import { DocxEditorViewportComponent } from '../viewport.component';
import { DocxHyperlinkPartDirective, DocxHyperlinkUnlinkDirective } from './hyperlink-parts';
import {
  DocxHyperlinkPopupRegistry,
  injectHyperlinkPopup,
  injectHyperlinkPopupInstance,
  type HyperlinkPopup,
  type HyperlinkPopupState,
} from './hyperlink-state';
import { DocxEditorHyperlinkComponent } from './hyperlink.component';

const LINK: SurfaceHyperlink = {
  id: 'link-1',
  paragraphId: 'p',
  kind: 'external',
  href: 'https://example.com/',
  authored: 'https://example.com/',
  text: 'example',
} as SurfaceHyperlink;

/** A popup in reading mode over `LINK`, so no engine or document is needed. */
function stubPopup(): { popup: HyperlinkPopup; unlinked: () => number } {
  let unlinked = 0;
  const state = signal<HyperlinkPopupState>({
    mode: 'reading',
    link: LINK,
    anchor: null,
    text: 'example',
    url: LINK.authored,
    copied: false,
    error: false,
    canEdit: true,
  });
  const noop = () => {};
  const popup: HyperlinkPopup = {
    state,
    open: noop,
    openAtCaret: noop,
    close: noop,
    copy: async () => true,
    beginEdit: noop,
    setText: noop,
    setUrl: noop,
    commitEdit: () => true,
    unlink: () => {
      unlinked += 1;
      return true;
    },
    openTarget: () => true,
  };
  return { popup, unlinked: () => unlinked };
}

const PART_IMPORTS = [
  DocxEditorHyperlinkComponent,
  DocxHyperlinkPartDirective,
  DocxHyperlinkUnlinkDirective,
];

@Component({
  selector: 'as-child-paired-host',
  imports: PART_IMPORTS,
  template: `<docx-editor-hyperlink>
    <ng-template docxHyperlinkPart="unlink" asChild className="paired-class">
      <button docxHyperlinkUnlink class="brand-button">Remove</button>
    </ng-template>
  </docx-editor-hyperlink>`,
})
class AsChildPairedHostComponent {}

@Component({
  selector: 'as-child-plain-host',
  imports: PART_IMPORTS,
  template: `<docx-editor-hyperlink>
    <ng-template docxHyperlinkPart="unlink" asChild>
      <button docxHyperlinkUnlink class="brand-button">Remove</button>
    </ng-template>
  </docx-editor-hyperlink>`,
})
class AsChildPlainHostComponent {}

@Component({
  selector: 'packaged-part-host',
  imports: PART_IMPORTS,
  template: `<docx-editor-hyperlink>
    <ng-template docxHyperlinkPart="unlink" keepContent className="paired-class" />
  </docx-editor-hyperlink>`,
})
class PackagedPartHostComponent {}

async function mountPart(component: new () => unknown) {
  const stub = stubPopup();
  configureZoneless([
    {
      provide: DocxHyperlinkPopupRegistry,
      useValue: { popup: stub.popup, attach: () => () => {} },
    },
  ]);
  const fixture = TestBed.createComponent(component);
  document.body.append(fixture.nativeElement);
  await settle(fixture, 1);
  const button = (fixture.nativeElement as HTMLElement).querySelector<HTMLButtonElement>(
    '[data-testid="hyperlink-popup-unlink"]'
  )!;
  return { fixture, button, unlinked: stub.unlinked };
}

describe('hyperlink popup state', () => {
  beforeEach(() => configureZoneless());

  // The claim of packages/vue/test/lifecycle-regression.test.ts.
  test('keeps text and address updates of the same task', () => {
    const popup = TestBed.runInInjectionContext(() => injectHyperlinkPopupInstance(false));
    popup.open();
    popup.setText('Link text');
    popup.setUrl('https://example.com');
    expect(popup.state()).toMatchObject({
      mode: 'editing',
      text: 'Link text',
      url: 'https://example.com',
    });
  });
});

describe('hyperlink parts', () => {
  afterEach(() => document.body.replaceChildren());

  test('asChild keeps the packaged class off the host element and the wiring on it', async () => {
    const { fixture, button, unlinked } = await mountPart(AsChildPairedHostComponent);
    expect(button.classList.contains('brand-button')).toBe(true);
    expect(button.classList.contains('paired-class')).toBe(true);
    expect(button.classList.contains('docx-hyperlink-popup__action')).toBe(false);
    button.click();
    await settle(fixture, 1);
    expect(unlinked()).toBe(1);
  });

  test("asChild with no className leaves only the host element's own class", async () => {
    const { button } = await mountPart(AsChildPlainHostComponent);
    expect(button.className).toBe('brand-button');
  });

  test('without asChild the packaged part keeps its class beside className', async () => {
    const { button } = await mountPart(PackagedPartHostComponent);
    expect(button.classList.contains('docx-hyperlink-popup__action')).toBe(true);
    expect(button.classList.contains('paired-class')).toBe(true);
    // `keepContent` keeps the packaged icon.
    expect(button.querySelector('svg')).not.toBeNull();
  });
});

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
/** One paragraph with an external link to `target`. */
function linkedDocx(target: string): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLink" Type="${R}/hyperlink" Target="${target}" TargetMode="External"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body><w:p><w:r><w:t xml:space="preserve">Visit </w:t></w:r>` +
        '<w:hyperlink r:id="rIdLink"><w:r><w:t>example</w:t></w:r></w:hyperlink></w:p></w:body></w:document>'
    ),
  });
}

/** A host part that exposes the shared popup. */
@Component({ selector: 'popup-probe', template: '' })
class PopupProbeComponent {
  readonly popup = injectHyperlinkPopup();
}

@Directive()
abstract class Host {
  readonly bytes = signal<Uint8Array>(linkedDocx('javascript:alert(1)'));
  readonly root = viewChild.required(DocxEditorRootDirective);
  readonly probe = viewChild.required(PopupProbeComponent);

  editor(): DocxEditorInstance {
    return engineOf(this.root().context.editor());
  }
}

@Component({
  selector: 'security-host',
  imports: [
    DocxEditorRootDirective,
    DocxEditorViewportComponent,
    DocxEditorContentComponent,
    DocxEditorHyperlinkComponent,
    PopupProbeComponent,
  ],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <docx-editor-viewport>
      <docx-editor-content />
      <docx-editor-hyperlink />
    </docx-editor-viewport>
    <popup-probe />
  </ng-container>`,
})
class SecurityHostComponent extends Host {}

describe('link targets', () => {
  beforeEach(() => configureZoneless());
  afterEach(() => {
    vi.restoreAllMocks();
    document.body.replaceChildren();
  });

  test('sanitizeHref refuses javascript targets', () => {
    expect(sanitizeHref('javascript:alert(1)').ok).toBe(false);
    expect(sanitizeHref('java\nscript:alert(1)').ok).toBe(false);
  });

  async function openLink(target: string) {
    const fixture = await render(SecurityHostComponent);
    const root: HTMLElement = fixture.nativeElement;
    document.body.append(root);
    const host = fixture.componentInstance;
    host.bytes.set(linkedDocx(target));
    await settle(fixture);
    const paragraphId = host.editor().surface!.session.paragraphIds()[0]!;
    host.editor().surface!.setSelection({
      anchor: { paragraphId, offset: 8 },
      head: { paragraphId, offset: 8 },
    });
    await settle(fixture);
    const link = host.editor().surface!.hyperlinks.linkAtCaret();
    expect(link).not.toBeNull();
    host.probe().popup.open(link);
    await settle(fixture);
    return { root, host };
  }

  test('a javascript: link shows no active target, and opening it does nothing', async () => {
    const open = vi.spyOn(window, 'open').mockReturnValue(null);
    const { root, host } = await openLink('javascript:alert(1)');
    const url = root.querySelector<HTMLElement>('[data-testid="hyperlink-popup-url"]')!;
    expect(url.tagName).toBe('SPAN');
    expect(url.hasAttribute('data-inert')).toBe(true);
    expect(url.title).toBe('This link points somewhere the editor will not open');
    expect(root.querySelector('[data-testid="hyperlink-popup-copy"]')).toBeNull();
    url.click();
    expect(host.probe().popup.openTarget()).toBe(false);
    expect(open).not.toHaveBeenCalled();
  });

  test('an http link is a button that opens its address', async () => {
    const open = vi.spyOn(window, 'open').mockReturnValue(null);
    const { root } = await openLink('https://example.com/');
    const url = root.querySelector<HTMLElement>('[data-testid="hyperlink-popup-url"]')!;
    expect(url.tagName).toBe('BUTTON');
    expect(url.hasAttribute('data-inert')).toBe(false);
    expect(root.querySelector('[data-testid="hyperlink-popup-copy"]')).not.toBeNull();
    url.click();
    expect(open).toHaveBeenCalledTimes(1);
    expect(open.mock.calls[0]![0]).toBe('https://example.com/');
  });
});
