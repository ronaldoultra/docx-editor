// Root inputs that need more than one engine setter: zoom, and the options that the engine reads
// only when it is created.
import { isDevMode } from '@angular/core';
import type { ZoomMode } from '@docx-editor.dev/core/contracts/editor';
import {
  resolveZoomMode,
  sameZoomMode,
  type ContentLockPolicy,
  type DocxEditorInstance,
} from '@docx-editor.dev/core/editor';

/** The zoom input values that the root last gave the engine. @internal */
export interface AppliedZoom {
  zoom: number | undefined;
  mode: ZoomMode | 'auto' | undefined;
}

function sameZoomInput(a: ZoomMode | 'auto', b: ZoomMode | 'auto'): boolean {
  if (a === b) return true;
  const left = resolveZoomMode(a);
  const right = resolveZoomMode(b);
  return left !== null && right !== null && sameZoomMode(left, right);
}

/**
 * Give the engine a changed `zoom` and `zoomMode`. A fixed zoom ends a fit mode in the engine,
 * so a declared fit mode applies again after each new zoom value.
 *
 * @internal
 */
export function applyZoomInputs(
  editor: DocxEditorInstance,
  applied: AppliedZoom,
  zoom: number | undefined,
  zoomMode: ZoomMode | 'auto' | undefined
): void {
  const zoomMoved = zoom !== undefined && zoom !== applied.zoom;
  if (zoomMoved) {
    applied.zoom = zoom;
    editor.setZoom(zoom);
  }
  if (zoomMode === undefined) return;
  const modeMoved = applied.mode === undefined || !sameZoomInput(applied.mode, zoomMode);
  const reassertFit = zoomMoved && resolveZoomMode(zoomMode)?.type === 'fit';
  if (!modeMoved && !reassertFit) return;
  applied.mode = zoomMode;
  editor.setZoomMode(zoomMode);
}

/** An option that the engine reads only when it is created. @internal */
export type CreationOption = 'fieldResults' | 'contentLockPolicy' | 'lockedControlNavigation';

const warned = new Set<CreationOption>();

const normalized: Record<CreationOption, (value: unknown) => unknown> = {
  fieldResults: (value) => value ?? 'atomic',
  // Compared by value, so a new object with the same policy does not warn.
  contentLockPolicy: (value) =>
    (value as ContentLockPolicy | undefined)?.allowFormattingInLockedContent === true,
  lockedControlNavigation: (value) => value ?? 'characters',
};

/**
 * Warn once in development when an option that the engine reads at creation changes later. The
 * engine keeps the value it was created with, so the change has no effect.
 *
 * @internal
 */
export function warnCreationOptionChanged(
  option: CreationOption,
  created: unknown,
  current: unknown
): void {
  const normalize = normalized[option];
  if (normalize(created) === normalize(current) || warned.has(option) || !isDevMode()) return;
  warned.add(option);
  console.warn(
    `[docx-editor] ${option} changed after the editor was created, so the change is ignored. ` +
      `The editor reads ${option} once; create a new editor to change it.`
  );
}
