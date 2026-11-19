// Task 5.2: the claims of packages/vue/test/context-menu-target.test.ts and
// context-menu-start.test.ts.
import { ChangeDetectionStrategy, Component, Directive, signal, viewChild } from '@angular/core';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import {
  contextMenuOutline,
  contextMenuPanel,
  pressKey,
  rightClick,
  scrollerOf,
} from '../../testing/context-menu';
import { docx } from '../../testing/docx';
import { configureZoneless, engineOf, render } from '../../testing/harness';
import { DocxEditorContentComponent } from '../content.component';
import { DocxEditorRootDirective } from '../root.directive';
import { DocxEditorViewportComponent } from '../viewport.component';
import { injectContextMenuAnchor, injectContextMenuTarget } from './context-menu-state';
import {
  DocxContextMenuItemComponent,
  DocxContextMenuStartComponent,
} from './context-menu-item.component';
import { DocxEditorContextMenuComponent } from './context-menu.component';

const SOURCE = docx(
  '<w:p><w:r><w:t>Visit example today</w:t></w:r></w:p>' +
    '<w:p><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Color"/><w:enabled/><w:ddList><w:result w:val="1"/><w:listEntry w:val="Red"/><w:listEntry w:val="Green"/><w:listEntry w:val="Blue"/></w:ddList></w:ffData></w:fldChar></w:r><w:r><w:instrText> FORMDROPDOWN </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Green</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
);
const PARTS = [
  DocxEditorRootDirective,
  DocxEditorViewportComponent,
  DocxEditorContentComponent,
  DocxEditorContextMenuComponent,
  DocxContextMenuItemComponent,
  DocxContextMenuStartComponent,
];

/** A host row that shows what the menu opened on. */
@Component({
  selector: 'opened-on',
  imports: [DocxContextMenuItemComponent],
  template: `<docx-context-menu-item [label]="text()" />`,
})
class OpenedOnComponent {
  private readonly anchor = injectContextMenuAnchor();
  private readonly target = injectContextMenuTarget();
  text(): string {
    return `${this.anchor()?.trigger ?? 'closed'}:${this.target()?.tagName ?? 'none'}`;
  }
}

/** A host component that renders a start: it is not a direct child of the menu. */
@Component({
  selector: 'host-suggestions',
  imports: [DocxContextMenuStartComponent, DocxContextMenuItemComponent],
  template: `<docx-context-menu-start>
    <docx-context-menu-item label="suggestion" />
  </docx-context-menu-start>`,
})
class SuggestionsComponent {}

@Directive()
abstract class Host {
  readonly bytes = signal<Uint8Array>(SOURCE);
  readonly root = viewChild.required(DocxEditorRootDirective);

  editor(): DocxEditorInstance {
    return engineOf(this.root().context.editor());
  }
}

@Component({
  selector: 'target-host',
  imports: [...PARTS, OpenedOnComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <docx-editor-viewport>
      <docx-editor-content />
      <div docxEditorContextMenu><opened-on /></div>
    </docx-editor-viewport>
  </ng-container>`,
})
class TargetHostComponent extends Host {}

@Component({
  selector: 'start-host',
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <docx-editor-viewport>
      <docx-editor-content />
      <div docxEditorContextMenu>
        <docx-context-menu-start><docx-context-menu-item label="first" /></docx-context-menu-start>
        <docx-context-menu-item label="tail" />
        <docx-context-menu-start><docx-context-menu-item label="second" /></docx-context-menu-start>
      </div>
    </docx-editor-viewport>
  </ng-container>`,
})
class StartHostComponent extends Host {}

@Component({
  selector: 'empty-start-host',
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <docx-editor-viewport>
      <docx-editor-content />
      <div docxEditorContextMenu><docx-context-menu-start /></div>
    </docx-editor-viewport>
  </ng-container>`,
})
class EmptyStartHostComponent extends Host {}

@Component({
  selector: 'nested-start-host',
  imports: [...PARTS, SuggestionsComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <docx-editor-viewport>
      <docx-editor-content />
      <div docxEditorContextMenu>
        <docx-context-menu-start><host-suggestions /></docx-context-menu-start>
      </div>
    </docx-editor-viewport>
  </ng-container>`,
})
class NestedStartHostComponent extends Host {}

@Component({
  selector: 'verbatim-start-host',
  imports: [...PARTS, SuggestionsComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <docx-editor-viewport>
      <docx-editor-content />
      <div docxEditorContextMenu [preset]="false"><host-suggestions /></div>
    </docx-editor-viewport>
  </ng-container>`,
})
class VerbatimStartHostComponent extends Host {}

@Component({
  selector: 'indirect-start-host',
  imports: [...PARTS, SuggestionsComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <docx-editor-viewport>
      <docx-editor-content />
      <div docxEditorContextMenu><host-suggestions /></div>
    </docx-editor-viewport>
  </ng-container>`,
})
class IndirectStartHostComponent extends Host {}

async function mount<T extends Host>(component: new () => T) {
  const fixture = await render(component);
  const root: HTMLElement = fixture.nativeElement;
  document.body.append(root);
  return {
    root,
    host: fixture.componentInstance,
    destroy: () => {
      fixture.destroy();
      root.remove();
    },
  };
}

const pages = (root: Element) => root.querySelector<HTMLElement>('.docx-pages')!;
const escape = () => pressKey(document.body, { key: 'Escape' });

describe('which right-clicks open the context menu', () => {
  beforeEach(() => configureZoneless());
  afterEach(() => {
    document.body.replaceChildren();
  });

  test('the pages and the canvas around them open the document menu', async () => {
    const view = await mount(TargetHostComponent);
    expect((await rightClick(pages(view.root))).defaultPrevented).toBe(true);
    expect(contextMenuPanel(view.root)).not.toBeNull();
    await escape();
    expect(contextMenuPanel(view.root)).toBeNull();
    expect((await rightClick(scrollerOf(view.root))).defaultPrevented).toBe(true);
    expect(contextMenuPanel(view.root)).not.toBeNull();
    view.destroy();
  });

  test('painted text, a painted drop-down field, and the mount beside the page open it', async () => {
    const view = await mount(TargetHostComponent);
    const select = view.root.querySelector('.docx-pages select[data-docx-form-dropdown]');
    expect(select).not.toBeNull();
    const targets = [
      view.root.querySelector('.docx-pages [data-start]')!,
      select!,
      view.root.querySelector('.docx-content-mount')!,
    ];
    for (const target of targets) {
      expect((await rightClick(target)).defaultPrevented).toBe(true);
      expect(contextMenuPanel(view.root)).not.toBeNull();
      await escape();
    }
    view.destroy();
  });

  test('a form control beside the pages keeps the browser menu, its focus and selection', async () => {
    // A stand-in for the link popup's address field, which mounts beside the pages.
    const view = await mount(TargetHostComponent);
    const field = document.createElement('input');
    field.value = 'https://example.com';
    view.root.querySelector('.docx-content-mount')!.append(field);
    field.focus();
    field.setSelectionRange(0, 5);
    const before = JSON.stringify(view.host.editor().surface!.session.part().root);
    for (const keyboard of [false, true]) {
      const event = await rightClick(field, undefined, keyboard);
      expect(event.defaultPrevented).toBe(false);
      expect(contextMenuPanel(view.root)).toBeNull();
      expect(document.activeElement).toBe(field);
      expect([field.selectionStart, field.selectionEnd]).toEqual([0, 5]);
    }
    expect(JSON.stringify(view.host.editor().surface!.session.part().root)).toBe(before);
    view.destroy();
  });

  test('a right-click on the open menu keeps it in place and suppresses the browser menu', async () => {
    const view = await mount(TargetHostComponent);
    await rightClick(pages(view.root));
    const panel = contextMenuPanel(view.root)!;
    const { left, top } = panel.style;
    const event = await rightClick(panel, { x: 400, y: 300 });
    expect(event.defaultPrevented).toBe(true);
    expect(contextMenuPanel(view.root)).toBe(panel);
    expect([panel.style.left, panel.style.top]).toEqual([left, top]);
    view.destroy();
  });

  test('a host row reads the trigger and the target of the open', async () => {
    const view = await mount(TargetHostComponent);
    const text = view.root.querySelector('.docx-pages [data-start]')!;
    await rightClick(text);
    expect(contextMenuOutline(view.root).at(-1)).toBe(`pointer:${text.tagName}`);
    await escape();
    await rightClick(pages(view.root), undefined, true);
    expect(contextMenuOutline(view.root).at(-1)).toBe('keyboard:none');
    view.destroy();
  });
});

describe('docx-context-menu-start (context-menu-start.test.ts)', () => {
  beforeEach(() => configureZoneless());
  afterEach(() => {
    document.body.replaceChildren();
    vi.restoreAllMocks();
  });

  async function outlineOf<T extends Host>(component: new () => T): Promise<string[]> {
    const view = await mount(component);
    await rightClick(pages(view.root));
    const outline = contextMenuOutline(view.root);
    view.destroy();
    return outline;
  }

  const notDirect = (messages: string[]) =>
    messages.some((message) => message.includes('not a direct child of the context menu'));

  function warnings(): string[] {
    const messages: string[] = [];
    vi.spyOn(console, 'warn').mockImplementation((message: unknown) => {
      messages.push(String(message));
    });
    return messages;
  }

  test('every start moves up, and the host rows beside it stay below', async () => {
    const outline = await outlineOf(StartHostComponent);
    expect(outline.slice(0, 3)).toEqual(['first', 'second', '---']);
    expect(outline.at(-1)).toBe('tail');
  });

  test('an empty start adds no separator', async () => {
    const outline = await outlineOf(EmptyStartHostComponent);
    expect(outline[0]).not.toBe('---');
  });

  test('a start in a start and a start with preset false do not warn', async () => {
    const messages = warnings();
    const nested = await outlineOf(NestedStartHostComponent);
    expect(nested.slice(0, 2)).toEqual(['suggestion', '---']);
    await outlineOf(VerbatimStartHostComponent);
    expect(notDirect(messages)).toBe(false);
  });

  test('a start that a component renders warns, and goes after the packaged rows', async () => {
    const messages = warnings();
    const outline = await outlineOf(IndirectStartHostComponent);
    expect(notDirect(messages)).toBe(true);
    expect(outline.at(-1)).toBe('suggestion');
  });
});
