import {
  DestroyRef,
  Injectable,
  NgZone,
  computed,
  effect,
  inject,
  signal,
  untracked,
  type Signal,
} from '@angular/core';
import type { Editor, SelectionPin } from '@docx-editor.dev/core/contracts/editor';
import {
  hostDocumentOf,
  hostWindowOf,
  type DocxEditorInstance,
  type SurfaceHyperlink,
} from '@docx-editor.dev/core/editor';
import { injectDocxEditorContext } from '../context';

/** Where the popup opens, in viewport coordinates. @public */
export interface HyperlinkPopupAnchor {
  readonly left: number;
  readonly top: number;
}

/** `'reading'` shows a link; `'editing'` shows the text and address fields. @public */
export type HyperlinkPopupMode = 'closed' | 'reading' | 'editing';

/** The state of the hyperlink popup. @public */
export interface HyperlinkPopupState {
  readonly mode: HyperlinkPopupMode;
  /** The link that the popup shows or edits, or `null` for a new link. */
  readonly link: SurfaceHyperlink | null;
  readonly anchor: HyperlinkPopupAnchor | null;
  /** The display text draft. */
  readonly text: string;
  /** The address draft. A `#` prefix names a bookmark. */
  readonly url: string;
  /** True after a copy of the address. */
  readonly copied: boolean;
  /** True after the engine refused the draft. */
  readonly error: boolean;
  /** True when the document takes edits at the caret. */
  readonly canEdit: boolean;
}

/**
 * The behavior of the hyperlink popup. `injectHyperlinkPopup()` returns the instance that an
 * editor root shares.
 *
 * @public
 */
export interface HyperlinkPopup {
  readonly state: Signal<HyperlinkPopupState>;
  /** Show `link`, or open an empty link draft. */
  open(link?: SurfaceHyperlink | null, anchor?: HyperlinkPopupAnchor | null): void;
  /** Edit the link at the caret, or draft a link over the selection or the word at the caret. */
  openAtCaret(): void;
  close(): void;
  /** Copy the link's address to the clipboard. */
  copy(): Promise<boolean>;
  beginEdit(): void;
  setText(text: string): void;
  setUrl(url: string): void;
  /** Apply the draft. A refused draft keeps the popup open with `error`. */
  commitEdit(): boolean;
  /** Remove the link. */
  unlink(): boolean;
  /** Follow the link: go to its bookmark, or open its address in a new window. */
  openTarget(): boolean;
}

/**
 * True for a link that a `HYPERLINK` field makes. Its text belongs to the field, so the popup
 * does not offer to edit or remove it.
 *
 * @public
 */
export function isFieldLink(link: SurfaceHyperlink): boolean {
  return link.paragraphId === '';
}

/**
 * The longest target that the check on each update reads. `selectedText` builds the whole
 * selected string, so a long selection is checked once, at Apply.
 */
const TARGET_CHECK_LIMIT = 4096;

const CLOSED: HyperlinkPopupState = Object.freeze({
  mode: 'closed' as const,
  link: null,
  anchor: null,
  text: '',
  url: '',
  copied: false,
  error: false,
  canEdit: true,
});

function linkUrl(link: SurfaceHyperlink): string {
  return link.kind === 'internal' ? `#${link.anchor ?? ''}` : link.authored;
}

function caretViewportAnchor(doc: Document | null): HyperlinkPopupAnchor | null {
  const usable = (rect: DOMRect | undefined): rect is DOMRect =>
    !!rect && (rect.width > 0 || rect.height > 0);
  if (!doc) return null;
  const rect = doc.querySelector('[data-docx-caret]')?.getBoundingClientRect();
  if (usable(rect)) return { left: rect.left, top: rect.bottom };
  const selection = doc.getSelection();
  if (selection && selection.rangeCount > 0) {
    const range = selection.getRangeAt(0);
    // Some test DOMs have no range geometry.
    const box =
      typeof range.getBoundingClientRect === 'function' ? range.getBoundingClientRect() : undefined;
    if (usable(box)) return { left: box.left, top: box.bottom };
    const node = range.startContainer;
    const element = node.nodeType === 1 ? (node as Element) : node.parentElement;
    const fallback = element?.getBoundingClientRect();
    if (usable(fallback)) return { left: fallback.left, top: fallback.bottom };
  }
  return null;
}

/**
 * Create a hyperlink popup instance. It registers with the engine as the link UI while
 * `active()` is true. Call it in an injection context.
 */
function createHyperlinkPopup(active: () => boolean): HyperlinkPopup {
  const context = injectDocxEditorContext();
  const zone = inject(NgZone);
  const editorOf = (): DocxEditorInstance | null => (context ? untracked(context.editor) : null);
  const state = signal<HyperlinkPopupState>(CLOSED);
  const set = (next: HyperlinkPopupState) => zone.run(() => state.set(next));
  let retained: { editor: Editor; pin: SelectionPin } | null = null;
  // The words a new-link draft opened over; empty for a draft at a bare caret.
  let targetText = '';
  // Whether those words were already a pending deletion when the draft opened.
  let struckAtOpen = false;

  const canEditNow = (): boolean => editorOf()?.can({ type: 'insertText', text: '' }).ok ?? false;
  const canEdit = computed(() => {
    context?.revision();
    const editor = context?.editor();
    return editor ? untracked(() => editor.can({ type: 'insertText', text: '' }).ok) : false;
  });

  const releaseSelectionPin = () => {
    targetText = '';
    struckAtOpen = false;
    if (!retained) return;
    retained.editor.releaseSelection(retained.pin);
    retained = null;
  };
  inject(DestroyRef).onDestroy(releaseSelectionPin);

  // Whether the words that a new-link draft will link are gone: a peer deleted or struck them,
  // or an undo took them out. Apply would otherwise insert the draft at the collapsed caret.
  const targetGone = (everyUpdate = false): boolean => {
    const editor = editorOf();
    if (targetText === '' || !editor) return false;
    if (!struckAtOpen && editor.surface?.hyperlinks.selectionIsPendingDeletion() === true) {
      return true;
    }
    if (everyUpdate && targetText.length > TARGET_CHECK_LIMIT) return false;
    return editor.query({ type: 'selectedText' }) !== targetText;
  };

  const open = (link?: SurfaceHyperlink | null, anchor?: HyperlinkPopupAnchor | null) => {
    releaseSelectionPin();
    set({
      mode: link ? 'reading' : 'editing',
      link: link ?? null,
      anchor: anchor ?? null,
      text: link?.text ?? '',
      url: link ? linkUrl(link) : '',
      copied: false,
      error: false,
      canEdit: canEditNow(),
    });
  };

  const close = () => {
    releaseSelectionPin();
    set(CLOSED);
  };

  const openAtCaret = () => {
    const editor = editorOf();
    const element = context ? untracked(context.contentElement) : null;
    const anchor = caretViewportAnchor(element ? hostDocumentOf(element) : null);
    releaseSelectionPin();
    const hyperlinks = editor?.surface?.hyperlinks;
    const link = hyperlinks?.linkAtCaret() ?? hyperlinks?.fieldLinkAtCaret() ?? null;
    if (link && isFieldLink(link)) {
      open(link, anchor);
      return;
    }
    if (link) {
      set({
        ...CLOSED,
        mode: 'editing',
        link,
        anchor,
        text: link.text,
        url: linkUrl(link),
        canEdit: canEditNow(),
      });
      return;
    }
    // A caret inside a word links that word: it becomes the selection and the display text.
    const word = hyperlinks?.wordAtCaret() ?? null;
    if (word) editor?.surface?.setSelection({ anchor: word.from, head: word.to });
    const selected = editor?.query({ type: 'selectedText' }) ?? '';
    const pin = editor?.retainSelection() ?? null;
    if (editor && pin) retained = { editor, pin };
    targetText = selected;
    struckAtOpen = hyperlinks?.selectionIsPendingDeletion() === true;
    set({ ...CLOSED, mode: 'editing', anchor, text: selected, canEdit: canEditNow() });
  };

  effect((onCleanup) => {
    const editor = context?.editor();
    if (!editor || !active()) return;
    const off = untracked(() =>
      editor.setHyperlinkChrome({
        onPopover: (activation) =>
          open(activation.link, { left: activation.rect.left, top: activation.rect.bottom }),
        onRequest: openAtCaret,
      })
    );
    onCleanup(off);
  });

  // A reading popup closes when the caret leaves its link; a new-link draft closes when its
  // words are gone.
  effect(() => {
    context?.revision();
    const editor = context?.editor() ?? null;
    untracked(() => {
      const current = state();
      if (current.mode === 'reading' && current.link) {
        const hyperlinks = editor?.surface?.hyperlinks;
        const atCaret = isFieldLink(current.link)
          ? (hyperlinks?.fieldLinkAtCaret() ?? null)
          : (hyperlinks?.linkAtCaret() ?? null);
        if (atCaret && atCaret.id === current.link.id) return;
        set(CLOSED);
        return;
      }
      if (current.mode === 'editing' && !current.link) {
        const surface = editor?.surface;
        if (!surface) return;
        if (surface.retainedSelection() && !targetGone(true)) return;
        close();
      }
    });
  });

  const copy = async (): Promise<boolean> => {
    const href = untracked(state).link?.href;
    if (!href) return false;
    const element = context ? untracked(context.contentElement) : null;
    try {
      await hostWindowOf(element).navigator.clipboard.writeText(href);
    } catch {
      return false;
    }
    set({ ...untracked(state), copied: true });
    return true;
  };

  const update = (change: Partial<HyperlinkPopupState>) =>
    set({ ...untracked(state), copied: false, error: false, ...change });

  const commitEdit = (): boolean => {
    const current = untracked(state);
    const hyperlinks = editorOf()?.surface?.hyperlinks;
    if (!hyperlinks) return false;
    if (!current.link && targetGone()) {
      close();
      return false;
    }
    const url = current.url.trim();
    if (url.length === 0) {
      set({ ...current, error: true });
      return false;
    }
    const internal = url.startsWith('#');
    const applied = hyperlinks.applyHyperlink({
      ...(internal ? { anchor: url.slice(1) } : { url }),
      ...(current.text.trim().length > 0 ? { text: current.text } : {}),
    });
    if (!applied) {
      set({ ...current, error: true });
      return false;
    }
    close();
    return true;
  };

  const unlink = (): boolean => {
    const hyperlinks = editorOf()?.surface?.hyperlinks;
    if (!hyperlinks) return false;
    const removed = hyperlinks.removeHyperlink(untracked(state).link?.id);
    if (removed) close();
    return removed;
  };

  const openTarget = (): boolean => {
    const navigation = editorOf()?.surface?.navigation;
    const link = untracked(state).link;
    if (!navigation || !link) return false;
    const opened =
      link.kind === 'internal'
        ? !!link.anchor && navigation.goToBookmark(link.anchor)
        : navigation.openExternal(link.href);
    if (opened) set(CLOSED);
    return opened;
  };

  return {
    state: computed(() => {
      const current = state();
      return current.mode === 'closed' ? CLOSED : { ...current, canEdit: canEdit() };
    }),
    open,
    openAtCaret,
    close,
    copy,
    beginEdit: () => update({ mode: 'editing' }),
    setText: (text) => update({ text }),
    setUrl: (url) => update({ url }),
    commitEdit,
    unlink,
    openTarget,
  };
}

/**
 * The popup state that an editor root shares. It registers with the engine as the link UI only
 * while a view shows it: `<docx-editor-hyperlink>`, or a host part that calls
 * `injectHyperlinkPopup()`. With no view, the link controls are disabled, so they never open a
 * popup that no one sees.
 *
 * @internal
 */
@Injectable()
export class DocxHyperlinkPopupRegistry {
  private readonly views = signal(0);
  /** The shared popup. */
  readonly popup: HyperlinkPopup = createHyperlinkPopup(() => this.views() > 0);

  /** Count one view. Returns the function that removes it. */
  attach(): () => void {
    this.views.update((n) => n + 1);
    let attached = true;
    return () => {
      if (!attached) return;
      attached = false;
      this.views.update((n) => n - 1);
    };
  }
}

/**
 * The hyperlink popup of the enclosing editor root. Calling it marks the caller as a view of the
 * shared popup until Angular destroys the caller. Outside a root, it returns an instance of its
 * own that registers with the engine itself. Call it in an injection context.
 *
 * @public
 */
export function injectHyperlinkPopup(): HyperlinkPopup {
  const registry = inject(DocxHyperlinkPopupRegistry, { optional: true });
  if (!registry) return createHyperlinkPopup(() => true);
  inject(DestroyRef).onDestroy(registry.attach());
  return registry.popup;
}

/**
 * A hyperlink popup instance of its own, not the one the root shares. With `active` false, it
 * does not register with the engine. Call it in an injection context.
 *
 * @public
 */
export function injectHyperlinkPopupInstance(active = true): HyperlinkPopup {
  return createHyperlinkPopup(() => active);
}
