/**
 * Capability gating helpers for the grid editor.
 *
 * These two predicates centralize the policy that controls when the toolbar,
 * menu, and grid mutators are enabled. Keeping them out of `App.vue` and the
 * composable makes the policy reusable across components and easier to
 * unit-test in isolation.
 */

/**
 * The grid is read-only whenever the host opened the database in
 * time-travel mode or the user is viewing a historical revision.
 */
export function isGridSessionReadOnly(options: { timeTravelDate: number | null; viewingHistorical: boolean }) {
  return options.timeTravelDate != null || options.viewingHistorical;
}

/**
 * Returns true when the UI should allow mutating the grid (formula commits,
 * formatting, structural edits) and exposing the Save action.
 *
 * The `dirty` flag is optional: pass `undefined` when only checking general
 * editability, and `false` when an action should only be enabled while
 * there are unsaved local edits.
 */
export function canMutateGrid(options: {
  canUpdate: boolean;
  hasDocument: boolean;
  dirty?: boolean;
  timeTravelDate: number | null;
  viewingHistorical: boolean;
}) {
  return options.canUpdate
    && options.hasDocument
    && options.dirty !== false
    && !isGridSessionReadOnly(options);
}
