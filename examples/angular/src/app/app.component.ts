import { ChangeDetectionStrategy, Component, computed, signal, viewChild } from '@angular/core';
import {
  DocxEditorComponent,
  type DocumentChange,
  type DocxEditorMode,
} from '@docx-editor.dev/angular';

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

@Component({
  selector: 'app-root',
  imports: [DocxEditorComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="toolbar">
      <input
        type="file"
        accept=".docx,${DOCX_MIME}"
        aria-label="Abrir arquivo .docx"
        (change)="open($event)"
      />
      <label>
        <input type="checkbox" [checked]="mode() === 'view'" (change)="toggleView($event)" />
        Somente leitura
      </label>
      <button type="button" [disabled]="!bytes() || saving()" (click)="save()">Salvar</button>
      <button type="button" [disabled]="!bytes()" (click)="close()">Fechar</button>
      <span class="status" data-testid="status">{{ status() }}</span>
    </div>
    @if (bytes(); as document) {
      <docx-editor [document]="document" [mode]="mode()" (change)="onChange($event)" />
    } @else {
      <div class="empty">Escolha um arquivo .docx.</div>
    }
  `,
})
export class AppComponent {
  protected readonly bytes = signal<ArrayBuffer | null>(null);
  protected readonly fileName = signal('documento.docx');
  protected readonly mode = signal<DocxEditorMode>('edit');
  protected readonly saving = signal(false);
  protected readonly changes = signal(0);
  protected readonly lastSource = signal('');

  private readonly editor = viewChild(DocxEditorComponent);
  private readonly zone = 'Zone' in globalThis ? 'com zone.js' : 'TESTE';

  protected readonly status = computed(() => {
    const parts = [this.zone, `${this.changes()} alterações`];
    if (this.lastSource()) parts.push(`última: ${this.lastSource()}`);
    return parts.join(' · ');
  });

  protected async open(event: Event) {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;
    this.fileName.set(file.name);
    this.changes.set(0);
    this.bytes.set(await file.arrayBuffer());
    input.value = '';
  }

  /** Removes the editor from the page, which destroys the component and its engine. */
  protected close() {
    this.bytes.set(null);
    this.changes.set(0);
    this.lastSource.set('');
  }

  protected toggleView(event: Event) {
    this.mode.set((event.target as HTMLInputElement).checked ? 'view' : 'edit');
  }

  protected onChange(change: DocumentChange) {
    this.changes.update((count) => count + 1);
    this.lastSource.set(change.source ?? 'edição');
  }

  protected async save() {
    const editor = this.editor();
    if (!editor) return;
    this.saving.set(true);
    try {
      const buffer = await editor.save();
      const url = URL.createObjectURL(new Blob([buffer], { type: DOCX_MIME }));
      const link = document.createElement('a');
      link.href = url;
      link.download = this.fileName();
      document.body.append(link);
      link.click();
      link.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } finally {
      this.saving.set(false);
    }
  }
}
