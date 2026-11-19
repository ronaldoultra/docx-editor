// The menu rows that are more than one command: the insert-table grid, the chart types, the
// reviewers, and Help › Report issue.

import {
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  NgZone,
  booleanAttribute,
  computed,
  inject,
  input,
  untracked,
} from '@angular/core';
import type { EditorSnapshot } from '@docx-editor.dev/core/contracts/editor';
import {
  CHART_INSERT_CHOICES,
  chartTypeIconPaths,
  chartTypeLabelKey,
  hostWindowOf,
  type ChartType,
  type ReviewAuthorInfo,
} from '@docx-editor.dev/core/editor';
import { injectDocxEditor } from '../context';
import {
  injectEditorCommand,
  injectEditorValueCommand,
  type EditorCommandState,
} from '../editor-command';
import { injectEditorState } from '../editor-state';
import { chromeControlForSlot } from '../toolbar/toolbar-registry';
import { DocxTableSizeGridComponent } from '../toolbar/toolbar-table-insert.component';
import { DocxMenuItemComponent } from './menu-item.component';
import { DocxMenuRowComponent, DocxMenuSeparatorComponent } from './menu-row.component';
import { DocxMenuBarState, injectMenuLabel, reportIssueUrl } from './menu-state';
import { DocxMenuSubmenuComponent } from './menu-submenu.component';

/**
 * Insert › Table: the size grid behind a submenu when the engine can insert a table, and a
 * disabled row with the engine's reason when it cannot. A submenu that opens onto a grid where
 * nothing can be picked is worse than a disabled row.
 *
 * @internal
 */
@Component({
  selector: 'docx-menu-table-picker',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxMenuItemComponent, DocxMenuSubmenuComponent, DocxTableSizeGridComponent],
  host: { style: 'display: contents' },
  template: `@if (command.isEnabled()) {
      <docx-menu-submenu [labelKey]="labelKey() ?? controlKey" [paths]="paths">
        <docx-table-size-grid [label]="label()('toolbar.insertTable')" (inserted)="close()" />
      </docx-menu-submenu>
    } @else {
      <docx-menu-item slotId="table.insert" [labelKey]="labelKey()" />
    }`,
})
export class DocxMenuTablePickerComponent {
  /** The plain label key of the row. */
  readonly labelKey = input<string | undefined>(undefined);

  private readonly bar = inject(DocxMenuBarState, { optional: true });

  /** @internal */
  protected readonly command: EditorCommandState = injectEditorCommand('table.insert');
  /** @internal */
  protected readonly label = injectMenuLabel();
  /** @internal */
  protected readonly paths = chromeControlForSlot('table.insert')?.paths ?? null;
  /** @internal */
  protected readonly controlKey = chromeControlForSlot('table.insert')?.labelKey ?? 'table.insert';

  /** @internal */
  protected close(): void {
    this.bar?.setOpenMenu(null);
  }
}

/**
 * Insert › Chart: the chart types behind a submenu when the engine can insert a chart, and a
 * disabled row with the engine's reason when it cannot.
 *
 * @internal
 */
@Component({
  selector: 'docx-menu-chart-picker',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxMenuItemComponent, DocxMenuRowComponent, DocxMenuSubmenuComponent],
  host: { style: 'display: contents' },
  template: `@if (command.isEnabled()) {
      <docx-menu-submenu [labelKey]="labelKey() ?? controlKey" [paths]="paths">
        @for (type of choices; track type) {
          <button
            docxMenuRow="insert.chart"
            [icon]="iconOf(type)"
            [attr.data-value]="type"
            (activate)="insert(type)"
          >
            {{ label()(labelOf(type)) }}
          </button>
        }
      </docx-menu-submenu>
    } @else {
      <docx-menu-item slotId="insert.chart" [labelKey]="labelKey()" />
    }`,
})
export class DocxMenuChartPickerComponent {
  /** The plain label key of the row. */
  readonly labelKey = input<string | undefined>(undefined);

  private readonly bar = inject(DocxMenuBarState, { optional: true });

  /** @internal */
  protected readonly command = injectEditorValueCommand('insert.chart');
  /** @internal */
  protected readonly label = injectMenuLabel();
  /** @internal */
  protected readonly choices = CHART_INSERT_CHOICES;
  /** @internal */
  protected readonly paths = chromeControlForSlot('insert.chart')?.paths ?? null;
  /** @internal */
  protected readonly controlKey = chromeControlForSlot('insert.chart')?.labelKey ?? 'insert.chart';
  /** @internal */
  protected readonly iconOf = chartTypeIconPaths;
  /** @internal */
  protected readonly labelOf = chartTypeLabelKey;

  /** @internal */
  protected insert(type: ChartType): void {
    this.command.execute(type);
    this.bar?.setOpenMenu(null);
  }
}

const selectHiddenAuthors = (snapshot: EditorSnapshot): readonly string[] =>
  snapshot.hiddenReviewAuthors ?? [];
const sameList = (a: readonly string[], b: readonly string[]) =>
  a.length === b.length && a.every((value, index) => value === b[index]);
const EMPTY_AUTHORS: readonly ReviewAuthorInfo[] = [];

/**
 * The Reviewers submenu: one checked row for each review author, and a row for all of them. It
 * changes the view only, not the document. Disabled with the engine's reason when the
 * document has no review authors to show.
 *
 * @example
 * ```html
 * <docx-menu menu="review">
 *   <docx-menu-reviewers />
 * </docx-menu>
 * ```
 *
 * @public
 */
@Component({
  selector: 'docx-menu-reviewers',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxMenuRowComponent, DocxMenuSeparatorComponent, DocxMenuSubmenuComponent],
  host: { style: 'display: contents' },
  template: `@if (!hidden()) {
    @if (command.isEnabled()) {
      <docx-menu-submenu labelKey="reviewers.label" [className]="className()">
        <button
          docxMenuRow="reviewer-all"
          rowClass="docx-menubar__reviewer-item"
          [active]="allVisible()"
          (activate)="toggleAll()"
        >
          <span docxMenuRowIcon class="docx-menubar__reviewer-check">{{
            allVisible() ? '✓' : ''
          }}</span>
          {{ label()('reviewers.all') }}
        </button>
        <div docxMenuSeparator></div>
        @for (info of authors(); track info.author) {
          <button
            rowClass="docx-menubar__reviewer-item"
            [docxMenuRow]="'reviewer-' + info.slot"
            [active]="!hiddenSet().has(info.author)"
            (activate)="toggle(info.author)"
          >
            <span docxMenuRowIcon class="docx-menubar__reviewer-check">{{
              hiddenSet().has(info.author) ? '' : '✓'
            }}</span>
            {{ info.author || '—' }}
          </button>
        }
      </docx-menu-submenu>
    } @else {
      <button docxMenuRow="review.authors" disabled [reason]="command.disabledReason()">
        {{ label()('reviewers.label') }}
      </button>
    }
  }`,
})
export class DocxMenuReviewersComponent {
  /** More classes for the submenu root. */
  readonly className = input<string | undefined>(undefined);
  /** Render nothing. */
  readonly hidden = input(false, { transform: booleanAttribute });

  private readonly editor = injectDocxEditor();
  private readonly zone = inject(NgZone);
  private readonly hiddenAuthors = injectEditorState(selectHiddenAuthors, sameList);

  /** @internal */
  protected readonly command: EditorCommandState = injectEditorCommand('review.authors');
  /** @internal */
  protected readonly label = injectMenuLabel();
  /** @internal */
  protected readonly authors = injectEditorState(() => {
    const editor = this.editor();
    return untracked(() => editor?.getReviewAuthors() ?? EMPTY_AUTHORS);
  });
  /** @internal */
  protected readonly hiddenSet = computed(() => new Set(this.hiddenAuthors()));
  /** @internal */
  protected readonly allVisible = computed(() => {
    const hidden = this.hiddenSet();
    return this.authors().every((info) => !hidden.has(info.author));
  });

  /** @internal */
  protected toggleAll(): void {
    const visible = untracked(this.allVisible);
    this.zone.runOutsideAngular(() => untracked(this.editor)?.setAllReviewAuthorsVisible(!visible));
  }

  /** @internal */
  protected toggle(author: string): void {
    const hidden = untracked(this.hiddenSet).has(author);
    this.zone.runOutsideAngular(() =>
      untracked(this.editor)?.setReviewAuthorVisible(author, hidden)
    );
  }
}

/**
 * Help › Report issue. The bar's `reportIssueHandler` replaces the packaged action, which opens
 * this project's issue form in a new tab. The bar's `reportIssue` input set to `false` removes
 * the row.
 *
 * @internal
 */
@Component({
  selector: 'docx-menu-report-issue',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxMenuRowComponent],
  host: { style: 'display: contents' },
  template: `@if (bar?.handlers()?.reportIssueRow !== false) {
    <button docxMenuRow="help.reportIssue" (activate)="run()">
      {{ label()('toolbar.reportIssue') }}
    </button>
  }`,
})
export class DocxMenuReportIssueComponent {
  /** @internal */
  protected readonly bar = inject(DocxMenuBarState, { optional: true });
  /** @internal */
  protected readonly label = injectMenuLabel();
  private readonly host: HTMLElement = inject(ElementRef).nativeElement;

  /** @internal */
  protected run(): void {
    const own = this.bar?.handlers().reportIssue;
    this.bar?.setOpenMenu(null);
    if (own) {
      own();
      return;
    }
    const view = hostWindowOf(this.host);
    view.open(reportIssueUrl(view), '_blank', 'noopener,noreferrer');
  }
}
