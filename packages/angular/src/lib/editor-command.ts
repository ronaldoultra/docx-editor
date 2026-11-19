import {
  DestroyRef,
  ElementRef,
  NgZone,
  computed,
  effect,
  inject,
  untracked,
  type Signal,
} from '@angular/core';
import type {
  EditorCommand,
  EditorExecOptions,
  ExecResult,
} from '@docx-editor.dev/core/contracts/editor';
import {
  IMAGE_WRAP_TARGETS,
  bindHistoryGroup,
  commandExecOptions,
  runToolbarCommand,
  toolbarCommandState,
  type ChromeSlotId,
  type HistoryGroupBinding,
  type HistoryGroupBindingOptions,
  type ObjectSlotValueMap,
  type ToolbarValueMap,
} from '@docx-editor.dev/core/editor';
import { localizeDisabledReason } from '@docx-editor.dev/core/i18n';
import { injectDocxEditor } from './context';
import { injectEditorState } from './editor-state';
import { injectTranslation } from './i18n';

/** A chrome slot or an engine command, as {@link injectEditorCommand} takes it. @public */
export type EditorCommandTarget = ChromeSlotId | EditorCommand;

/** What {@link injectEditorCommand} returns. @public */
export interface EditorCommandState {
  /**
   * Run the command. Returns `true` when the engine applied it, and `false` when the command is
   * disabled or the engine refused it. Pass the options of `injectHistoryGroup` to group a
   * gesture into one undo step.
   */
  readonly execute: (options?: unknown) => boolean;
  /** True when the command's state is on, such as bold text at the caret. */
  readonly isActive: Signal<boolean>;
  /** True when the command can run. The engine's toolbar command state decides it. */
  readonly isEnabled: Signal<boolean>;
  /** The translated reason when the command is disabled, otherwise `null`. */
  readonly disabledReason: Signal<string | null>;
  /**
   * What the control shows, for slots whose state is a value and not on or off, such as the
   * editing mode and the format painter's `off`, `once`, or `locked`. `null` for other slots
   * and for engine commands.
   */
  readonly value: Signal<string | null>;
}

interface CommandSlice {
  readonly active: boolean;
  readonly enabled: boolean;
  readonly disabledReason: string | null;
  readonly value: string | null;
}

/**
 * Read command state from the engine outside the reactive context. A state read can flush
 * pending input, and the engine then publishes `selectionChange` at once. A listener that
 * writes a signal must not run inside a `computed` (NG0600), because the error stops the flush
 * and the engine applies the same input again.
 */
const readEngine = <T>(read: () => T): T => untracked(read);

const sameSlice = (a: CommandSlice, b: CommandSlice): boolean =>
  a.active === b.active &&
  a.enabled === b.enabled &&
  a.disabledReason === b.disabledReason &&
  a.value === b.value;

/**
 * The state of a chrome slot or an engine command, and a function that runs it. Pass a function
 * to change the target later: the state follows the signals that the function reads. Outside an
 * editor root, the command is disabled and `execute` returns `false`. Call it in an injection
 * context.
 *
 * @example
 * ```ts
 * readonly bold = injectEditorCommand('text.bold');
 * ```
 *
 * @public
 */
export function injectEditorCommand(
  target: EditorCommandTarget | (() => EditorCommandTarget)
): EditorCommandState {
  const editor = injectDocxEditor();
  const zone = inject(NgZone);
  const { t } = injectTranslation();
  const current = typeof target === 'function' ? target : () => target;
  const slice = injectEditorState<CommandSlice>(() => {
    const goal = current();
    const instance = editor();
    const translate = t();
    if (typeof goal === 'string') {
      const state = readEngine(() => toolbarCommandState(instance, goal));
      return {
        active: state.active,
        enabled: state.enabled,
        disabledReason: localizeDisabledReason(state.disabledReason, translate),
        value: state.value ?? null,
      };
    }
    if (!instance) return { active: false, enabled: false, disabledReason: null, value: null };
    const { allowed, active } = readEngine(() => ({
      allowed: instance.can(goal),
      active: instance.isActive(goal),
    }));
    return {
      active,
      enabled: allowed.ok,
      disabledReason: allowed.ok ? null : localizeDisabledReason(allowed.reason ?? null, translate),
      value: null,
    };
  }, sameSlice);
  const execute = (input?: unknown): boolean => {
    const options = commandExecOptions(input);
    const goal = untracked(current);
    const instance = untracked(editor);
    return zone.runOutsideAngular(() => {
      if (typeof goal === 'string') {
        return runToolbarCommand(instance, goal, undefined, options ?? {}).ok;
      }
      return instance ? instance.exec(goal, options).ok : false;
    });
  };
  return {
    execute,
    isActive: computed(() => slice().active),
    isEnabled: computed(() => slice().enabled),
    disabledReason: computed(() => slice().disabledReason),
    value: computed(() => slice().value),
  };
}

/** A value slot that {@link injectEditorValueCommand} binds. @public */
export type EditorValueSlot = Exclude<keyof ToolbarValueMap, keyof ObjectSlotValueMap>;

/** What {@link injectEditorValueCommand} returns. @public */
export interface EditorValueCommandState<T extends string | number> {
  /** Apply a value through the engine. */
  readonly execute: (value: T, options?: EditorExecOptions) => ExecResult;
  /** The current value from the snapshot, or `null` when the selection has no single value. */
  readonly value: Signal<T | null>;
  /** The values the slot offers, for slots with a fixed set. */
  readonly options: Signal<readonly T[]>;
  /** True when the slot can apply a value. */
  readonly isEnabled: Signal<boolean>;
  /** The translated reason when the slot is disabled, otherwise `null`. */
  readonly disabledReason: Signal<string | null>;
}

/**
 * Bind a value slot, such as `font.size` or `font.family`. Font sizes are in half-points.
 * Colors are six hexadecimal digits without `#`. Call it in an injection context.
 *
 * @public
 */
export function injectEditorValueCommand<K extends EditorValueSlot>(
  slotId: K
): EditorValueCommandState<ToolbarValueMap[K]> {
  const editor = injectDocxEditor();
  const zone = inject(NgZone);
  const { t } = injectTranslation();
  const state = injectEditorState(
    () => {
      const instance = editor();
      return readEngine(() => toolbarCommandState(instance, slotId));
    },
    (a, b) =>
      a.value === b.value && a.enabled === b.enabled && a.disabledReason === b.disabledReason
  );
  const numeric = slotId === 'font.size' || slotId === 'list.lineSpacing';
  return {
    execute: (value, options) =>
      zone.runOutsideAngular(() => runToolbarCommand(untracked(editor), slotId, value, options)),
    value: computed(() => {
      const value = state().value;
      return (value === undefined ? null : numeric ? Number(value) : value) as
        ToolbarValueMap[K] | null;
    }),
    options: computed(
      () => (slotId === 'image.wrap' ? IMAGE_WRAP_TARGETS : []) as readonly ToolbarValueMap[K][]
    ),
    isEnabled: computed(() => state().enabled),
    disabledReason: computed(() => localizeDisabledReason(state().disabledReason, t())),
  };
}

/** Options for {@link injectHistoryGroup}. @public */
export interface HistoryGroupOptions extends HistoryGroupBindingOptions {
  /** The control element. You can also give it later with `ref`. */
  readonly element?: Signal<ElementRef<HTMLElement> | HTMLElement | null | undefined>;
}

/** What {@link injectHistoryGroup} returns. @public */
export interface HistoryGroupState {
  /** Set the control element, or `null` to unbind it. */
  readonly ref: (element: HTMLElement | null) => void;
  /**
   * The exec options for the current gesture. Pass them to `execute`. Throws when the control
   * has no element or no engine.
   */
  readonly options: () => EditorExecOptions;
}

/**
 * Group the values of one gesture on a control, such as a held slider or stepper, into one
 * undo step. Apply the values with `injectEditorValueCommand` or `injectEditorCommand`, and pass
 * `options()` with each value. Call it in an injection context.
 *
 * @public
 */
export function injectHistoryGroup(options: HistoryGroupOptions): HistoryGroupState {
  const editor = injectDocxEditor();
  const { kind } = options;
  let element: HTMLElement | null = null;
  let binding: HistoryGroupBinding | null = null;
  let destroyed = false;
  const rebind = () => {
    binding?.dispose();
    binding = null;
    const instance = untracked(editor);
    if (destroyed || !instance || !element) return;
    binding = bindHistoryGroup(instance, element, { kind });
  };
  const ref = (node: HTMLElement | null) => {
    element = node;
    rebind();
  };
  effect(() => {
    editor();
    const given = options.element?.();
    if (options.element)
      element = given instanceof ElementRef ? given.nativeElement : (given ?? null);
    untracked(rebind);
  });
  inject(DestroyRef).onDestroy(() => {
    destroyed = true;
    rebind();
  });
  return {
    ref,
    options() {
      if (!binding) throw new Error('history group control is not mounted');
      return binding.options();
    },
  };
}
