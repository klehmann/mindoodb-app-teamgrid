// The workbooks opened in this launch (the toolbar's Window menu). The editor
// shows one workbook at a time; switching saves the shown one and opens the
// other from MindooDB, at the sheet and selection it was left at.

/** Where a workbook was left: the active sheet and its selection (A1, e.g. "B2:D9"). */
export interface WindowView {
  sheetName: string
  range?: string
}

export interface WorkbookWindow {
  id: string
  title: string
  view?: WindowView
}

export class WorkbookWindows {
  /** In the order they were opened (the Window menu's order). */
  private readonly windows: WorkbookWindow[] = []
  /** Most recently shown last. */
  private readonly recency: string[] = []
  private activeId: string | null = null

  get active(): WorkbookWindow | undefined {
    return this.activeId === null ? undefined : this.find(this.activeId)
  }

  list(): readonly WorkbookWindow[] {
    return this.windows
  }

  find(id: string): WorkbookWindow | undefined {
    return this.windows.find((window) => window.id === id)
  }

  /** The workbook now shown: added when it is new, else brought forward (keeps its view). */
  show(id: string, title: string): WorkbookWindow {
    let window = this.find(id)
    if (window) window.title = title
    else this.windows.push((window = { id, title }))
    this.activeId = id
    this.touch(id)
    return window
  }

  /** No workbook is shown (welcome screen); the windows stay open. */
  hide(): void {
    this.activeId = null
  }

  rename(id: string, title: string): void {
    const window = this.find(id)
    if (window) window.title = title
  }

  remember(id: string, view: WindowView | undefined): void {
    const window = this.find(id)
    if (window && view) window.view = view
  }

  /** Closes a window; returns the one to show next (the most recently shown), if any. */
  close(id: string): WorkbookWindow | undefined {
    const index = this.windows.findIndex((window) => window.id === id)
    if (index >= 0) this.windows.splice(index, 1)
    const recent = this.recency.indexOf(id)
    if (recent >= 0) this.recency.splice(recent, 1)
    if (this.activeId === id) this.activeId = null
    const next = this.recency[this.recency.length - 1]
    return next === undefined ? undefined : this.find(next)
  }

  private touch(id: string): void {
    const at = this.recency.indexOf(id)
    if (at >= 0) this.recency.splice(at, 1)
    this.recency.push(id)
  }
}
