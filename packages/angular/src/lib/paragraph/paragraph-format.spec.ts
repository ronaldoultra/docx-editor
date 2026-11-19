// Task 6.3: the paragraph format claims of packages/vue/test/paragraph-hooks-lock-state.test.ts.
// The ruler claims (`useParagraphIndent`) port with the rulers in phase 9.
import { ChangeDetectionStrategy, Component, viewChild } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { afterEach, beforeEach, expect, test } from 'vitest';
import { docx } from '../../testing/docx';
import { configureZoneless, engineOf, settle } from '../../testing/harness';
import { DocxEditorContentComponent } from '../content.component';
import { DocxEditorRootDirective } from '../root.directive';
import { DocxEditorViewportComponent } from '../viewport.component';
import { injectParagraphFormat } from './paragraph-format';

const LOCKED =
  '<w:sdt><w:sdtPr><w:tag w:val="t"/><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent>' +
  '<w:p><w:r><w:t>Locked title</w:t></w:r></w:p></w:sdtContent></w:sdt>';
const BOUND =
  '<w:sdt><w:sdtPr><w:tag w:val="t"/><w:dataBinding w:storeItemID="{1}" w:xpath="/a"/></w:sdtPr>' +
  '<w:sdtContent><w:p><w:r><w:t>Bound title</w:t></w:r></w:p></w:sdtContent></w:sdt>';
const BOUND_LOCKED =
  '<w:sdt><w:sdtPr><w:tag w:val="t"/><w:lock w:val="contentLocked"/>' +
  '<w:dataBinding w:storeItemID="{1}" w:xpath="/a"/></w:sdtPr>' +
  '<w:sdtContent><w:p><w:r><w:t>Bound title</w:t></w:r></w:p></w:sdtContent></w:sdt>';
const FREE = '<w:p><w:r><w:t>Free text</w:t></w:r></w:p>';

@Component({
  selector: 'spec-format-probe',
  template: `<output data-format>{{ format.isEnabled() }}</output>`,
})
class FormatProbeComponent {
  readonly format = injectParagraphFormat();
}

@Component({
  selector: 'spec-format-host',
  imports: [
    DocxEditorRootDirective,
    DocxEditorViewportComponent,
    DocxEditorContentComponent,
    FormatProbeComponent,
  ],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes">
    <docx-editor-viewport>
      <spec-format-probe />
      <docx-editor-content />
    </docx-editor-viewport>
  </ng-container>`,
})
class FormatHostComponent {
  bytes!: Uint8Array;
  readonly root = viewChild.required(DocxEditorRootDirective);
  readonly probe = viewChild.required(FormatProbeComponent);
}

beforeEach(() => configureZoneless());
afterEach(() => document.body.replaceChildren());

// Indents and alignment pass a content lock, and a data binding mirrors only the text, so the
// dialog stays available in both, as in a binding that also locks its content.
for (const [name, first] of [
  ['a content lock', LOCKED],
  ['a data binding', BOUND],
  ['a data binding with a content lock', BOUND_LOCKED],
] as const) {
  test(`the paragraph format state follows the caret out of and back into ${name}`, async () => {
    const fixture = TestBed.createComponent(FormatHostComponent);
    fixture.componentInstance.bytes = docx(first + FREE);
    document.body.append(fixture.nativeElement);
    await settle(fixture);
    const editor = engineOf(fixture.componentInstance.root().context.editor());
    const ids = editor.surface!.session.paragraphIds();
    const put = async (index: number) => {
      const at = { paragraphId: ids[index]!, offset: 1 };
      editor.surface!.setSelection({ anchor: at, head: at });
      await settle(fixture, 2);
      return fixture.componentInstance.probe().format.isEnabled();
    };
    expect(await put(0)).toBe(true);
    expect(await put(1)).toBe(true);
    expect(await put(0)).toBe(true);
    // A real refusal: the document open for viewing takes it away, in and out of the control.
    expect(editor.exec({ type: 'setEditingMode', mode: 'viewing' }).ok).toBe(true);
    await settle(fixture, 2);
    expect(await put(0)).toBe(false);
    expect(await put(1)).toBe(false);
    fixture.destroy();
  });
}
