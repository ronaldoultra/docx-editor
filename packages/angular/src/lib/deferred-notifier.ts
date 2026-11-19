// Coalesced, deferred notification of engine state changes.
//
// The engine publishes its layout after the event that changed the document. A state read in
// the same task can see the snapshot from before layout, so the root raises its tick in a later
// microtask, and one tick serves every change in a burst. Under sustained input, a run of
// microtask waves yields through a task now and then, so input handling keeps its priority.

/** How many consecutive notification waves may take the microtask path before one yields. */
const MAX_MICROTASK_WAVES_BEFORE_YIELD = 16;

let waveStreak = 0;
let waveOpen = false;
let streakResetArmed = false;

let pendingTasks: (() => void)[] = [];
let taskChannel: MessageChannel | null = null;

/**
 * Run `task` in a later event-loop task, through a message and not a timer. Timers belong to the
 * script's window, and a hidden opener window throttles them when the editor is in a pop-out.
 */
function yieldTask(task: () => void): void {
  if (typeof MessageChannel !== 'function') {
    globalThis.setTimeout(task, 0);
    return;
  }
  if (!taskChannel) {
    taskChannel = new MessageChannel();
    taskChannel.port1.onmessage = () => {
      const tasks = pendingTasks;
      pendingTasks = [];
      for (const run of tasks) {
        // One failing subscriber must not stop the notifications of the others.
        try {
          run();
        } catch (error) {
          queueMicrotask(() => {
            throw error;
          });
        }
      }
    };
  }
  pendingTasks.push(task);
  if (pendingTasks.length === 1) taskChannel.port2.postMessage(null);
}

/** True when this notification must yield through a task. */
function notificationYieldsToTask(): boolean {
  const scheduling = (
    globalThis as typeof globalThis & {
      navigator?: {
        scheduling?: { isInputPending?: (options?: { includeContinuous?: boolean }) => boolean };
      };
    }
  ).navigator?.scheduling;
  // Without the Scheduling API there is no input cascade to break. Tests and engines other
  // than Chromium keep the immediate path.
  if (typeof scheduling?.isInputPending !== 'function') return false;
  if (scheduling.isInputPending({ includeContinuous: true })) return true;
  if (!waveOpen) {
    waveOpen = true;
    waveStreak += 1;
    queueMicrotask(() => {
      waveOpen = false;
    });
    if (!streakResetArmed) {
      streakResetArmed = true;
      yieldTask(() => {
        streakResetArmed = false;
        waveStreak = 0;
      });
    }
  }
  return waveStreak > MAX_MICROTASK_WAVES_BEFORE_YIELD;
}

/** A function that schedules one call of `notify`, however often it runs before the call. */
export function deferredTick(notify: () => void): () => void {
  let scheduled = false;
  return () => {
    if (scheduled) return;
    scheduled = true;
    const flush = () => {
      scheduled = false;
      notify();
    };
    if (notificationYieldsToTask()) yieldTask(flush);
    else queueMicrotask(flush);
  };
}
