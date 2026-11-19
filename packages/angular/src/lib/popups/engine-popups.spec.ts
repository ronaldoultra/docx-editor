// Task 6.1: the routing claims of packages/vue/test/engine-popup-renderers.test.ts. The claims
// about the packaged widget's own UI port with task 9.3, and the invalid-field acknowledgement
// with task 6.4.
import {
  ChangeDetectionStrategy,
  Component,
  Directive,
  effect,
  input,
  signal,
  viewChild,
  type OnInit,
} from '@angular/core';
import type { ContentControlWidgetSession, DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { strToU8, zipSync } from 'fflate';
import { afterEach, beforeEach, describe, expect, test } from 'vitest';
import { docx } from '../../testing/docx';
import { configureZoneless, engineOf, render, settle } from '../../testing/harness';
import { DocxEditorContentComponent } from '../content.component';
import { injectDocxEditor } from '../context';
import { DocxEditorRootDirective } from '../root.directive';
import { DocxEditorViewportComponent } from '../viewport.component';
import type { DocxEditorPopups } from './popup-config';

const DROPDOWN = docx(
  '<w:sdt><w:sdtPr><w:dropDownList><w:listItem w:displayText="One" w:value="1"/><w:listItem w:displayText="Two" w:value="2"/></w:dropDownList></w:sdtPr><w:sdtContent><w:p><w:r><w:t>One</w:t></w:r></w:p></w:sdtContent></w:sdt>'
);
const CHECKBOX = docx(
  '<w:p><w:r><w:t xml:space="preserve">Done: </w:t></w:r><w:sdt><w:sdtPr>' +
    '<w14:checkbox xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w14:checked w14:val="0"/>' +
    '<w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>' +
    '</w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t xml:space="preserve"> yes</w:t></w:r></w:p>'
);

const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
};
const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
  ),
  (c) => c.charCodeAt(0)
);
/** A picture content control around a 1x1 PNG. */
const PICTURE = zipSync({
  '[Content_Types].xml': strToU8(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  ),
  '_rels/.rels': strToU8(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${NS.r}/officeDocument" Target="word/document.xml"/></Relationships>`
  ),
  'word/_rels/document.xml.rels': strToU8(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImg" Type="${NS.r}/image" Target="media/image1.png"/></Relationships>`
  ),
  'word/media/image1.png': PNG,
  'word/document.xml': strToU8(
    `<w:document xmlns:w="${NS.w}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}" xmlns:r="${NS.r}"><w:body>` +
      '<w:p><w:r><w:t xml:space="preserve">Photo: </w:t></w:r><w:sdt><w:sdtPr><w:alias w:val="Photo"/><w:picture/></w:sdtPr>' +
      '<w:sdtContent><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
      '<wp:extent cx="457200" cy="457200"/><wp:docPr id="3" name="pic"/>' +
      `<a:graphic><a:graphicData uri="${NS.pic}"><pic:pic>` +
      '<pic:nvPicPr><pic:cNvPr id="3" name="pic"/><pic:cNvPicPr/></pic:nvPicPr>' +
      '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
      '<pic:spPr><a:xfrm><a:ext cx="457200" cy="457200"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr>' +
      '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:sdtContent></w:sdt></w:p>' +
      '</w:body></w:document>'
  ),
});

/** The kinds of the sessions that the spec widget received. */
let seen: ContentControlWidgetSession['kind'][] = [];

/** A widget renderer: a list of options, Escape cancels, and a checkbox toggles at once. */
@Component({
  selector: 'spec-widget',
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<div
    data-docx-popup="contentControlWidget"
    [attr.data-kind]="session().kind"
    (keydown.escape)="session().cancel()"
  >
    @for (item of session().items; track item.value) {
      <button role="option" (click)="session().apply(item.value)">{{ item.displayText }}</button>
    }
  </div>`,
})
class SpecWidgetComponent implements OnInit {
  readonly session = input.required<ContentControlWidgetSession>();

  ngOnInit(): void {
    const session = this.session();
    seen.push(session.kind);
    if (session.kind === 'checkbox') session.apply(session.value === 'true' ? 'false' : 'true');
  }
}

/** A host part that registers its own widget renderer, and records the sessions it gets. */
@Component({ selector: 'spec-manual-widget', template: '' })
class ManualWidgetComponent {
  static last: ContentControlWidgetSession | undefined;
  constructor() {
    const editor = injectDocxEditor();
    effect((onCleanup) => {
      const instance = editor();
      if (!instance) return;
      onCleanup(
        instance.setContentControlWidgetChrome({
          onRequest: (session) => {
            ManualWidgetComponent.last = session;
          },
        })
      );
    });
  }
}

@Directive()
abstract class Host {
  readonly bytes = signal<Uint8Array>(DROPDOWN);
  readonly popups = signal<DocxEditorPopups | undefined>(undefined);
  readonly manual = signal(false);
  readonly root = viewChild.required(DocxEditorRootDirective);

  editor(): DocxEditorInstance {
    return engineOf(this.root().context.editor());
  }
}

@Component({
  selector: 'engine-popups-host',
  imports: [
    DocxEditorRootDirective,
    DocxEditorViewportComponent,
    DocxEditorContentComponent,
    ManualWidgetComponent,
  ],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()" [popups]="popups()">
    <docx-editor-viewport>
      <docx-editor-content />
      @if (manual()) {
        <spec-manual-widget />
      }
    </docx-editor-viewport>
  </ng-container>`,
})
class EngineHostComponent extends Host {}

async function mount(bytes: Uint8Array, popups: DocxEditorPopups | undefined, manual = false) {
  const fixture = await render(EngineHostComponent);
  fixture.componentInstance.bytes.set(bytes);
  fixture.componentInstance.popups.set(popups);
  fixture.componentInstance.manual.set(manual);
  await settle(fixture);
  return fixture;
}

/** A pointerdown on a painted widget. jsdom has no PointerEvent. */
function press(container: HTMLElement, kind: string): void {
  const widget = container.querySelector(`[data-docx-cc-widget="${kind}"]`);
  expect(widget).not.toBeNull();
  widget!.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true, button: 0 }));
}

describe('engine popups', () => {
  beforeEach(() => {
    configureZoneless();
    seen = [];
    ManualWidgetComponent.last = undefined;
  });
  afterEach(() => document.body.replaceChildren());

  for (const mode of ['native', 'custom', 'manual'] as const) {
    test(`value widget ownership: ${mode}`, async () => {
      const fixture = await mount(
        DROPDOWN,
        {
          contentControlWidget:
            mode === 'native' ? undefined : mode === 'manual' ? false : SpecWidgetComponent,
        },
        mode === 'manual'
      );
      const container = fixture.nativeElement as HTMLElement;
      document.body.append(container);
      press(container, 'dropdown');
      await settle(fixture);
      expect(container.querySelectorAll('.docx-content-control-menu').length).toBe(
        mode === 'native' ? 1 : 0
      );
      const popups = container.querySelectorAll('[data-docx-popup="contentControlWidget"]');
      expect(popups.length).toBe(mode === 'custom' ? 1 : 0);
      if (mode === 'custom') {
        // The popup renders inside the content, not on the painted pages.
        expect(popups[0]!.closest('.docx-content-mount')).not.toBeNull();
        expect(popups[0]!.closest('.docx-paginated-surface')).toBeNull();
        // Escape cancels without a write and removes the popup.
        popups[0]!.dispatchEvent(
          new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: 'Escape' })
        );
        await settle(fixture);
        expect(container.querySelector('[data-docx-popup="contentControlWidget"]')).toBeNull();
        press(container, 'dropdown');
        await settle(fixture);
        const options = [...container.querySelectorAll<HTMLButtonElement>('[role="option"]')];
        expect(options.map((option) => option.textContent)).toEqual(['One', 'Two']);
        options[1]!.click();
        await settle(fixture);
        expect(fixture.componentInstance.editor().surface!.session.bodyText()).toContain('Two');
        expect(container.querySelector('[data-docx-popup="contentControlWidget"]')).toBeNull();
      }
      if (mode === 'manual') {
        // A `false` entry leaves the session to the host's own registration.
        const session = ManualWidgetComponent.last;
        expect(session).toBeDefined();
        expect(session!.signal.aborted).toBe(false);
        expect(session!.apply('2')).toBe(true);
        await settle(fixture);
        expect(fixture.componentInstance.editor().surface!.session.bodyText()).toContain('Two');
      }
    });
  }

  for (const configured of [false, true]) {
    test(`checkbox presses reach a renderer only through contentControlCheckbox (${configured})`, async () => {
      const fixture = await mount(CHECKBOX, {
        contentControlWidget: SpecWidgetComponent,
        ...(configured ? { contentControlCheckbox: SpecWidgetComponent } : {}),
      });
      const container = fixture.nativeElement as HTMLElement;
      press(container, 'checkbox');
      await settle(fixture);
      // Either way the box flips: the engine toggles it when no renderer takes the press, and
      // the renderer applies the toggle when one does.
      expect(
        container.querySelector('[data-docx-cc-widget="checkbox"]')?.getAttribute('data-checked')
      ).toBe('true');
      // A renderer for drop-down lists never sees the checkbox session.
      expect(seen).toEqual(configured ? ['checkbox'] : []);
    });
  }

  for (const configured of [false, true]) {
    test(`picture presses reach a renderer only through contentControlPicture (${configured})`, async () => {
      const fixture = await mount(PICTURE, {
        contentControlWidget: SpecWidgetComponent,
        ...(configured ? { contentControlPicture: SpecWidgetComponent } : {}),
      });
      const container = fixture.nativeElement as HTMLElement;
      press(container, 'picture');
      await settle(fixture);
      // Without the entry, the engine arms its own picker on the pages layer.
      expect(container.querySelector('[data-docx-popup="contentControlWidget"]') !== null).toBe(
        configured
      );
      expect(container.querySelector('.docx-content-control-picture-picker') !== null).toBe(
        !configured
      );
      expect(seen).toEqual(configured ? ['picture'] : []);
    });
  }

  test('a live false entry cancels the next session, and a removed entry gives presses back', async () => {
    const fixture = await mount(DROPDOWN, { contentControlWidget: SpecWidgetComponent });
    const container = fixture.nativeElement as HTMLElement;
    fixture.componentInstance.popups.set({ contentControlWidget: false });
    await settle(fixture);
    press(container, 'dropdown');
    await settle(fixture);
    expect(container.querySelector('[data-docx-popup="contentControlWidget"]')).toBeNull();
    expect(container.querySelector('.docx-content-control-menu')).toBeNull();
    fixture.componentInstance.popups.set({});
    await settle(fixture);
    press(container, 'dropdown');
    await settle(fixture);
    expect(container.querySelectorAll('.docx-content-control-menu').length).toBe(1);
  });
});
