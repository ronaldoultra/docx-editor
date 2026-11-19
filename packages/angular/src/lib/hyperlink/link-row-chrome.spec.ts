// Task 6.2: the claims of packages/vue/test/link-row-chrome.test.ts, and the popup claims of
// context-menu-link.test.ts and context-menu-target.test.ts.
import { ChangeDetectionStrategy, Component, Directive, signal, viewChild } from '@angular/core';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import { contextMenuPanel, rightClick, surfaceOf } from '../../testing/context-menu';
import { docx } from '../../testing/docx';
import { configureZoneless, engineOf, render, settle } from '../../testing/harness';
import { DocxEditorContentComponent } from '../content.component';
import { DocxEditorContextMenuComponent } from '../context-menu/context-menu.component';
import { DocxEditorRootDirective } from '../root.directive';
import { DocxToolbarButtonComponent } from '../toolbar/toolbar-button.component';
import { DocxEditorToolbarComponent } from '../toolbar/toolbar.component';
import { DocxEditorViewportComponent } from '../viewport.component';
import type { DocxEditorPopups } from '../popups/popup-config';
import { injectHyperlinkPopup } from './hyperlink-state';
import { DocxEditorHyperlinkComponent } from './hyperlink.component';

const SOURCE = docx('<w:p><w:r><w:t>Visit example today</w:t></w:r></w:p>');
const LINKED = docx(
  '<w:p><w:r><w:t xml:space="preserve">Visit </w:t></w:r>' +
    '<w:hyperlink w:anchor="target"><w:r><w:t>example</w:t></w:r></w:hyperlink>' +
    '<w:r><w:t xml:space="preserve"> today</w:t></w:r></w:p>'
);
const PID = '/word/document.xml#0.0.0';
const PARTS = [
  DocxEditorRootDirective,
  DocxEditorViewportComponent,
  DocxEditorContentComponent,
  DocxEditorToolbarComponent,
  DocxToolbarButtonComponent,
  DocxEditorContextMenuComponent,
  DocxEditorHyperlinkComponent,
];

/** A host part that shows the shared popup's mode. */
@Component({
  selector: 'host-link-panel',
  template: `<output data-testid="host-link-panel">{{ popup.state().mode }}</output>`,
})
class HostLinkPanelComponent {
  readonly popup = injectHyperlinkPopup();
}

@Directive()
abstract class Host {
  readonly bytes = signal<Uint8Array>(SOURCE);
  readonly popup = signal(true);
  readonly popups = signal<DocxEditorPopups | undefined>(undefined);
  readonly root = viewChild.required(DocxEditorRootDirective);

  editor(): DocxEditorInstance {
    return engineOf(this.root().context.editor());
  }
}

@Component({
  selector: 'link-row-host',
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()" [popups]="popups()">
    <fieldset docxEditorToolbar [preset]="false" [overflow]="false">
      <button docxToolbarButton="text.link"></button>
    </fieldset>
    <docx-editor-viewport>
      <docx-editor-content />
      <div docxEditorContextMenu></div>
      @if (popup()) {
        <docx-editor-hyperlink />
      }
    </docx-editor-viewport>
  </ng-container>`,
})
class LinkRowHostComponent extends Host {}

@Component({
  selector: 'link-panel-host',
  imports: [...PARTS, HostLinkPanelComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <fieldset docxEditorToolbar [preset]="false" [overflow]="false">
      <button docxToolbarButton="text.link"></button>
    </fieldset>
    <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
    <host-link-panel />
  </ng-container>`,
})
class HostPanelHostComponent extends Host {}

async function mount<T extends Host>(
  component: new () => T,
  setup: (host: T) => void = () => {},
  selection: [number, number] = [6, 13]
) {
  const fixture = await render(component);
  const root: HTMLElement = fixture.nativeElement;
  document.body.append(root);
  setup(fixture.componentInstance);
  await settle(fixture);
  const host = fixture.componentInstance;
  host.editor().surface!.setSelection({
    anchor: { paragraphId: PID, offset: selection[0] },
    head: { paragraphId: PID, offset: selection[1] },
  });
  await settle(fixture);
  return { root, host, fixture };
}

const toolbarLink = (root: Element) =>
  root.querySelector<HTMLButtonElement>('fieldset button[data-slot="text.link"]')!;
const popup = (root: Element) => root.querySelector<HTMLElement>('[data-testid="hyperlink-popup"]');
const byTestId = <T extends HTMLElement = HTMLElement>(root: Element, id: string) =>
  root.querySelector<T>(`[data-testid="${id}"]`)!;

async function openContextMenu(root: HTMLElement): Promise<HTMLElement> {
  await rightClick(surfaceOf(root));
  return contextMenuPanel(root)!.querySelector<HTMLElement>('[data-slot="text.link"]')!;
}

async function typeInto(input: HTMLInputElement, value: string): Promise<void> {
  input.value = value;
  input.dispatchEvent(new Event('input', { bubbles: true }));
  await settle(undefined, 1);
}

const key = (target: HTMLElement, name: string) =>
  target.dispatchEvent(
    new KeyboardEvent('keydown', { key: name, bubbles: true, cancelable: true })
  );

beforeEach(() => configureZoneless());
afterEach(() => document.body.replaceChildren());

describe('link controls', () => {
  test('with the popup, the row and the button are enabled', async () => {
    const { root } = await mount(LinkRowHostComponent);
    expect(toolbarLink(root).disabled).toBe(false);
    expect((await openContextMenu(root)).getAttribute('aria-disabled')).toBeNull();
  });

  test('without a link UI, the row and the button are disabled with a reason', async () => {
    const { root } = await mount(LinkRowHostComponent, (host) => host.popup.set(false));
    const button = toolbarLink(root);
    expect(button.disabled).toBe(true);
    expect(button.title).toBe("Links can't be added here");
    const row = await openContextMenu(root);
    expect(row.getAttribute('aria-disabled')).toBe('true');
    expect(row.title).toBe("Links can't be added here");
    row.click();
    await settle();
    expect(popup(root)).toBeNull();
  });

  test("a host's own link UI enables both controls, and they open it", async () => {
    const { root, host } = await mount(LinkRowHostComponent, (h) => h.popup.set(false));
    let requests = 0;
    const unregister = host.editor().setHyperlinkChrome({ onRequest: () => (requests += 1) });
    await settle();
    expect(toolbarLink(root).disabled).toBe(false);
    toolbarLink(root).click();
    expect(requests).toBe(1);
    const row = await openContextMenu(root);
    expect(row.getAttribute('aria-disabled')).toBeNull();
    row.click();
    await settle();
    expect(requests).toBe(2);
    expect(contextMenuPanel(root)).toBeNull();
    expect(popup(root)).toBeNull();
    unregister();
    await settle();
    expect(toolbarLink(root).disabled).toBe(true);
  });

  test('a host panel that calls injectHyperlinkPopup() keeps the controls working', async () => {
    const { root } = await mount(HostPanelHostComponent);
    expect(toolbarLink(root).disabled).toBe(false);
    toolbarLink(root).click();
    await settle();
    expect(byTestId(root, 'host-link-panel').textContent).toBe('editing');
  });

  test('a caret inside a word opens the popup with that word, and Apply links it', async () => {
    const { root, host } = await mount(LinkRowHostComponent, undefined, [8, 8]);
    toolbarLink(root).click();
    await settle();
    expect(byTestId<HTMLInputElement>(root, 'hyperlink-popup-text').value).toBe('example');
    await typeInto(byTestId(root, 'hyperlink-popup-url-input'), 'https://example.com');
    byTestId<HTMLButtonElement>(root, 'hyperlink-popup-apply').click();
    await settle();
    const [link] = host.editor().surface!.hyperlinks.linksInCaretParagraph();
    expect(link?.text).toBe('example');
    expect(host.editor().surface!.session.bodyText()).toBe('Visit example today');
    expect(popup(root)).toBeNull();
  });
});

describe('the popup from the context menu', () => {
  test('on an existing link the row says Edit link and opens the edit fields', async () => {
    const { root } = await mount(LinkRowHostComponent, (host) => host.bytes.set(LINKED), [8, 8]);
    const row = await openContextMenu(root);
    expect(row.textContent).toContain('Edit link');
    row.click();
    await settle();
    expect(popup(root)?.dataset['mode']).toBe('editing');
    expect(byTestId<HTMLInputElement>(root, 'hyperlink-popup-url-input').value).toBe('#target');
  });

  async function openFromMenu() {
    const view = await mount(LinkRowHostComponent);
    (await openContextMenu(view.root)).click();
    await settle();
    const url = byTestId<HTMLInputElement>(view.root, 'hyperlink-popup-url-input');
    expect(document.activeElement).toBe(url);
    // jsdom cannot focus the editable pages, so count the engine's focus calls.
    const focus = vi.spyOn(view.host.editor(), 'focus');
    return { ...view, url, focus };
  }

  test('Escape closes it, gives the focus back to the document, and keeps the selection', async () => {
    const { root, host, url, focus } = await openFromMenu();
    const before = host.editor().surface!.state().selection;
    key(url, 'Escape');
    await settle();
    expect(popup(root)).toBeNull();
    expect(focus).toHaveBeenCalled();
    expect(host.editor().surface!.state().selection).toEqual(before);
  });

  test('Enter applies the link and gives the focus back to the document', async () => {
    const { root, host, url, focus } = await openFromMenu();
    await typeInto(url, 'https://example.com/a');
    key(url, 'Enter');
    await settle();
    expect(popup(root)).toBeNull();
    expect(focus).toHaveBeenCalled();
    expect(host.editor().surface!.hyperlinks.linkAtCaret()?.href).toBe('https://example.com/a');
  });

  test('Cancel closes it and gives the focus back to the document', async () => {
    const { root, focus } = await openFromMenu();
    byTestId(root, 'hyperlink-popup-cancel').click();
    await settle();
    expect(popup(root)).toBeNull();
    expect(focus).toHaveBeenCalled();
  });

  test('an empty address refuses Apply, and a press outside closes the popup', async () => {
    const { root, url } = await openFromMenu();
    expect(byTestId<HTMLButtonElement>(root, 'hyperlink-popup-apply').disabled).toBe(true);
    key(url, 'Enter');
    await settle();
    expect(byTestId(root, 'hyperlink-popup-error').getAttribute('role')).toBe('alert');
    document.body.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true, button: 0 }));
    await settle();
    expect(popup(root)).toBeNull();
  });
});

describe('the popup as a registry entry', () => {
  async function openRegistered() {
    const view = await mount(LinkRowHostComponent, (host) => {
      host.popup.set(false);
      host.popups.set({ hyperlink: DocxEditorHyperlinkComponent });
    });
    toolbarLink(view.root).click();
    await settle();
    const url = byTestId<HTMLInputElement>(view.root, 'hyperlink-popup-url-input');
    await typeInto(url, 'https://example.com');
    url.focus();
    url.setSelectionRange(0, 5);
    return { ...view, url };
  }

  test('it renders inside the content mount, beside the pages', async () => {
    const { url } = await openRegistered();
    expect(url.closest('.docx-content-mount')).not.toBeNull();
    expect(url.closest('.docx-paginated-surface')).toBeNull();
  });

  test('a right-click in the address field keeps the browser menu, the focus, and the draft', async () => {
    const { root, host, url } = await openRegistered();
    const selection = host.editor().surface!.state().selection;
    const event = await rightClick(url);
    expect(event.defaultPrevented).toBe(false);
    expect(contextMenuPanel(root)).toBeNull();
    expect(document.activeElement).toBe(url);
    expect(url.value).toBe('https://example.com');
    expect([url.selectionStart, url.selectionEnd]).toEqual([0, 5]);
    expect(popup(root)?.dataset['mode']).toBe('editing');
    expect(host.editor().surface!.state().selection).toEqual(selection);
  });

  test('the panel and its buttons keep the browser menu', async () => {
    const { root, url } = await openRegistered();
    for (const id of ['hyperlink-popup', 'hyperlink-popup-cancel']) {
      expect((await rightClick(byTestId(root, id))).defaultPrevented).toBe(false);
      expect(contextMenuPanel(root)).toBeNull();
      expect(document.activeElement).toBe(url);
    }
  });
});
