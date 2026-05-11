export function isGridSessionReadOnly(options: { timeTravelDate: number | null; viewingHistorical: boolean }) {
  return options.timeTravelDate != null || options.viewingHistorical;
}

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
