import {
  ChangeDetectionStrategy,
  Component,
  DestroyRef,
  ViewContainerRef,
  computed,
  effect,
  inject,
  input,
  untracked,
  type ComponentRef,
  type Type,
} from '@angular/core';
import { popupComponent, popupInputs, type DocxEditorPopup } from './popup-config';

const sessionKeys = new WeakMap<object, number>();
let nextSessionKey = 0;

/**
 * A stable key for each form session. The outlet renders a new instance for a new key, so no
 * draft, focus return, or late progress of the earlier session carries over.
 *
 * @internal
 */
export function sessionKey(session: object): number {
  let key = sessionKeys.get(session);
  if (key === undefined) {
    key = ++nextSessionKey;
    sessionKeys.set(session, key);
  }
  return key;
}

/**
 * Render one popup registry entry as a sibling of this element. An omitted or `false` entry
 * renders nothing. The same component type keeps its instance when the registry object or the
 * inputs change, and gets only the inputs that changed. A new `session` renders a new instance.
 *
 * @internal
 */
@Component({
  selector: 'docx-popup-outlet',
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: { style: 'display: none' },
  template: '',
})
export class DocxPopupOutletComponent {
  /** The registry entry to render. */
  readonly entry = input<DocxEditorPopup<object> | undefined>(undefined);
  /** The inputs that the editor sets on the popup. */
  readonly inputs = input<Readonly<Record<string, unknown>>>({});
  /** The form session. A new session renders a new instance. */
  readonly session = input<object | null>(null);

  private readonly container = inject(ViewContainerRef);
  private current: { type: Type<unknown>; key: number; ref: ComponentRef<unknown> } | null = null;
  private applied: Record<string, unknown> = {};

  constructor() {
    const component = computed(() => popupComponent(this.entry()));
    const key = computed(() => {
      const session = this.session();
      return session ? sessionKey(session) : 0;
    });
    const inputs = computed(() => popupInputs(this.entry(), this.inputs()));
    effect(() => {
      const type = component();
      const next = key();
      const values = inputs();
      untracked(() => this.render(type, next, values));
    });
    inject(DestroyRef).onDestroy(() => this.clear());
  }

  private render(type: Type<unknown> | null, key: number, inputs: Record<string, unknown>): void {
    if (!type) {
      this.clear();
      return;
    }
    if (this.current?.type !== type || this.current.key !== key) {
      this.clear();
      this.current = { type, key, ref: this.container.createComponent(type) };
    }
    const ref = this.current.ref;
    for (const [name, value] of Object.entries(inputs)) {
      if (!(name in this.applied) || this.applied[name] !== value) ref.setInput(name, value);
    }
    this.applied = inputs;
  }

  private clear(): void {
    this.current?.ref.destroy();
    this.current = null;
    this.applied = {};
  }
}
