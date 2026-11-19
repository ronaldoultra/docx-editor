import {
  ChangeDetectionStrategy,
  Component,
  NgZone,
  computed,
  inject,
  input,
  output,
  signal,
  viewChild,
} from '@angular/core';
import {
  DocxEditorContentComponent,
  DocxEditorI18nDirective,
  DocxEditorRootDirective,
  DocxEditorViewportComponent,
  DocxEditorWindowComponent,
  type DocumentChange,
  type DocxEditorColorMode,
  type DocxEditorDirection,
  type DocxEditorRootMode,
} from '@docx-editor.dev/angular';
import { en, ptBR, type Translations } from '@docx-editor.dev/core/i18n';
import { EditorStatusComponent } from './editor-status.component';
import { FormatBarComponent } from './format-bar.component';

type Language = 'en' | 'pt-BR';

const CATALOGS: Record<Language, Translations> = { en, 'pt-BR': ptBR };
const LOCALES: Record<Language, string> = { en: 'en-US', 'pt-BR': 'pt-BR' };

/**
 * An editor made from the composition API: the root, a custom formatting bar, the viewport and
 * content, a pop-out window, and a status bar. The options above it set the root inputs.
 */
@Component({
  selector: 'app-composed-editor',
  imports: [
    DocxEditorI18nDirective,
    DocxEditorRootDirective,
    DocxEditorViewportComponent,
    DocxEditorContentComponent,
    DocxEditorWindowComponent,
    FormatBarComponent,
    EditorStatusComponent,
  ],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="options">
      <label>
        Modo
        <select data-testid="composed-mode" (change)="mode.set(pick($event))">
          <option value="edit" selected>Edição</option>
          <option value="view">Somente leitura</option>
          <option value="suggesting">Sugestões</option>
        </select>
      </label>
      <label>
        Autor
        <input type="text" [value]="author()" (change)="author.set(text($event))" />
      </label>
      <label>
        Tema
        <select data-testid="composed-theme" (change)="colorMode.set(pick($event))">
          <option value="dark">Escuro</option>
          <option value="light">Claro</option>
          <option value="system" selected>Sistema</option>
        </select>
      </label>
      <label>
        Idioma
        <select data-testid="composed-language" (change)="language.set(pick($event))">
          <option value="pt-BR" selected>Português (Brasil)</option>
          <option value="en">English</option>
        </select>
      </label>
      <label>
        Direção
        <select data-testid="composed-direction" (change)="direction.set(pick($event))">
          <option value="" selected>Do idioma</option>
          <option value="ltr">Esquerda para a direita</option>
          <option value="rtl">Direita para a esquerda</option>
        </select>
      </label>
      <label>
        <input type="checkbox" [checked]="popout()" (change)="popout.set(checked($event))" />
        Páginas em outra janela
      </label>
    </div>
    <div class="composed" [docxEditorI18n]="catalog()" [docxEditorDir]="dir()">
      <ng-container
        docxEditorRoot
        [document]="document()"
        [mode]="mode()"
        [colorMode]="colorMode()"
        [locale]="locale()"
        [author]="author()"
        (ready)="ready.emit()"
        (change)="change.emit($event)"
      >
        <app-format-bar />
        <docx-editor-window [(open)]="popout" title="docx-editor">
          <ng-template>
            <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
          </ng-template>
        </docx-editor-window>
        <app-editor-status />
      </ng-container>
    </div>
  `,
})
export class ComposedEditorComponent {
  /** The DOCX bytes to open. */
  readonly document = input.required<ArrayBuffer>();
  /** Emits once, when the first document is open. */
  readonly ready = output<void>();
  /** Emits for each change that the engine commits. */
  readonly change = output<DocumentChange>();

  protected readonly mode = signal<DocxEditorRootMode>('edit');
  protected readonly author = signal('Exemplo Angular');
  protected readonly colorMode = signal<DocxEditorColorMode>('system');
  protected readonly language = signal<Language>('pt-BR');
  protected readonly direction = signal<DocxEditorDirection | ''>('');
  protected readonly popout = signal(false);

  protected readonly catalog = computed(() => CATALOGS[this.language()]);
  protected readonly locale = computed(() => LOCALES[this.language()]);
  protected readonly dir = computed(() => this.direction() || undefined);

  private readonly root = viewChild.required(DocxEditorRootDirective);
  private readonly zone = inject(NgZone);

  /** Write the current document to DOCX bytes. */
  save(): Promise<ArrayBuffer> {
    const editor = this.root().context.editor();
    if (!editor) return Promise.reject(new Error('The editor has no document.'));
    return this.zone.runOutsideAngular(() => editor.save());
  }

  protected pick<T extends string>(event: Event): T {
    return (event.target as HTMLSelectElement).value as T;
  }

  protected text(event: Event): string {
    return (event.target as HTMLInputElement).value;
  }

  protected checked(event: Event): boolean {
    return (event.target as HTMLInputElement).checked;
  }
}
