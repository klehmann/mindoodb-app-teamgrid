/**
 * Reading and reordering the workbook's tab order.
 *
 * `worksheetOrder` stays the only source of truth for that order: the document
 * validator requires the array, Vega's sheet documents read the same subtree,
 * and the XLSX export writes its sheets in exactly that sequence. A
 * per-worksheet sort key would be a second ordering that could disagree with
 * the one the export uses, so a move edits the list itself.
 *
 * Automerge lists have no move operation, so a move is a delete plus an
 * insert. Two people moving the same tab at once therefore both insert, and
 * the id can end up in the list twice. Rather than try to prevent that,
 * {@link resolveWorksheetOrder} makes it harmless: every list that a merge can
 * produce reads back as a sound order, so a doubled tab or a sheet without a
 * tab is not a state the UI can reach.
 */
import type { Workbook, WorksheetId } from "@/features/document/lib/teamgridDocument";

/**
 * The stored list as a sound one: every id once, in the stored sequence.
 *
 * Ids the workbook does not know are dropped, and worksheets missing from the
 * list are appended in id order so two clients repairing the same document
 * independently still arrive at the same answer. Tombstones stay in — this is
 * the order of the list, not of the visible tabs, and the indices it yields
 * have to keep lining up with the array being patched.
 */
export function resolveWorksheetOrder(workbook: Workbook): WorksheetId[] {
  const seen = new Set<WorksheetId>();
  const order: WorksheetId[] = [];
  for (const id of workbook.worksheetOrder) {
    if (!seen.has(id) && workbook.worksheetsById[id]) {
      seen.add(id);
      order.push(id);
    }
  }
  const missing = Object.keys(workbook.worksheetsById)
    .filter((id) => !seen.has(id))
    .sort();
  return [...order, ...missing];
}

/**
 * Whether the stored list already reads back unchanged.
 *
 * A move may only be persisted as the minimal delete-plus-insert while this
 * holds, because those two indices address the array as stored.
 */
export function isWorksheetOrderSound(workbook: Workbook): boolean {
  const resolved = resolveWorksheetOrder(workbook);
  return resolved.length === workbook.worksheetOrder.length
    && resolved.every((id, index) => workbook.worksheetOrder[index] === id);
}

export interface WorksheetMovePlan {
  /** The whole list after the move. */
  order: WorksheetId[];
  /** Where the dragged id sits in the list handed in. */
  fromIndex: number;
  /** Where it goes back in once that entry has been taken out. */
  toIndex: number;
}

/**
 * Send a tab to `toIndex`, which is both where it sits in the resulting list
 * and where a patch inserts it once the old entry is gone.
 *
 * The two indices are the pair a delete-then-insert patch needs, and applying
 * them in that order has to reproduce `order` exactly, which is why both come
 * out of the same splice pair rather than being computed separately.
 *
 * A finished drag is one call. Expressing it as several — a step per tab the
 * pointer crossed — would leave the outcome depending on whether the host
 * applies a patch's deletes and inserts in pairs or all deletes first.
 */
export function planWorksheetMoveTo(
  order: readonly WorksheetId[],
  worksheetId: WorksheetId,
  toIndex: number,
): WorksheetMovePlan | null {
  const fromIndex = order.indexOf(worksheetId);
  if (fromIndex < 0 || toIndex < 0 || toIndex >= order.length || fromIndex === toIndex) {
    return null;
  }
  const next = [...order];
  next.splice(fromIndex, 1);
  next.splice(toIndex, 0, worksheetId);
  return { order: next, fromIndex, toIndex };
}

/**
 * Drop `draggedId` onto `targetId`'s tab. Dragging right lands after the
 * target, dragging left before it — the gap the pointer is over in both cases.
 */
export function planWorksheetMove(
  order: readonly WorksheetId[],
  draggedId: WorksheetId,
  targetId: WorksheetId,
): WorksheetMovePlan | null {
  return planWorksheetMoveTo(order, draggedId, order.indexOf(targetId));
}

/**
 * One step left or right, so the order is not drag-only.
 *
 * Phrased as a move onto the neighbouring live tab rather than as an index
 * shift, because a tombstoned entry between the two would otherwise swallow
 * the step and the tab would appear not to move.
 */
export function planWorksheetNudge(
  workbook: Workbook,
  worksheetId: WorksheetId,
  offset: -1 | 1,
): WorksheetMovePlan | null {
  const order = resolveWorksheetOrder(workbook);
  const live = order.filter((id) => !workbook.worksheetsById[id]?.deletedAt);
  const index = live.indexOf(worksheetId);
  if (index < 0) {
    return null;
  }
  const neighbour = live[index + offset];
  return neighbour ? planWorksheetMove(order, worksheetId, neighbour) : null;
}
