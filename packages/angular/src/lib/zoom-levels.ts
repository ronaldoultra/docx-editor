// The zoom steps of the zoom shortcuts. 25% is below the fit floor on purpose: a fit that sits
// on the floor must still offer a step down.
export const ZOOM_LEVELS: readonly number[] = [0.25, 0.5, 0.75, 1, 1.25, 1.5, 2];

/** The next zoom step from `zoom`, or `null` at the end of the steps. */
export function stepZoomLevel(zoom: number, direction: 'in' | 'out'): number | null {
  const epsilon = 0.001;
  return direction === 'in'
    ? (ZOOM_LEVELS.find((level) => level > zoom + epsilon) ?? null)
    : ([...ZOOM_LEVELS].reverse().find((level) => level < zoom - epsilon) ?? null);
}

/** The zoom for Ctrl or Cmd with `key`, or `null` when the key is not a zoom shortcut. */
export function zoomLevelForShortcut(key: string, zoom: number): number | null {
  if (key === '0') return 1;
  if (key === '+' || key === '=') return stepZoomLevel(zoom, 'in') ?? zoom;
  if (key === '-') return stepZoomLevel(zoom, 'out') ?? zoom;
  return null;
}
