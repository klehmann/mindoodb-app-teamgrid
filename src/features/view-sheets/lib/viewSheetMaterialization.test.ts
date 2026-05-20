import { describe, expect, it, vi } from "vitest";
import type {
  MindooDBAppResolvedViewDefinition,
  MindooDBAppViewEntry,
  MindooDBAppViewNavigator,
} from "mindoodb-app-sdk";

import {
  buildViewSheetWorksheet,
  materializeViewSheet,
  parseRootCategoryPath,
  viewValueToCellValue,
  type ViewSheetSettings,
} from "@/features/view-sheets/lib/viewSheetMaterialization";

describe("view sheet materialization", () => {
  it("parses nested category paths", () => {
    expect(parseRootCategoryPath(" Customers \\ ACME \\ 2026 ")).toEqual(["Customers", "ACME", "2026"]);
    expect(parseRootCategoryPath("")).toEqual([]);
  });

  it("converts SDK values into supported TeamGrid cell values", () => {
    expect(viewValueToCellValue(12)).toEqual({ kind: "number", value: 12 });
    expect(viewValueToCellValue("2026-05-20")).toEqual({ kind: "date", isoDate: "2026-05-20T00:00:00.000Z", format: "date" });
    expect(viewValueToCellValue(["a", "b"])).toEqual({ kind: "string", text: "[\"a\",\"b\"]" });
    expect(viewValueToCellValue(null)).toEqual({ kind: "empty" });
  });

  it("builds category and document rows with stable row and column IDs", () => {
    const worksheet = buildViewSheetWorksheet({
      settings: settings(),
      view: viewDefinition(),
      rootCategoryPath: ["Customers"],
      entries: entries(),
      lastViewCursor: "cursor_1",
      now: () => new Date("2026-05-20T12:00:00.000Z"),
    });
    const refreshed = buildViewSheetWorksheet({
      settings: settings({ title: "Contacts refreshed" }),
      view: viewDefinition(),
      rootCategoryPath: ["Customers"],
      entries: entries(),
      existingWorksheet: worksheet,
      lastViewCursor: "cursor_2",
      now: () => new Date("2026-05-21T12:00:00.000Z"),
    });

    expect(refreshed.id).toBe(worksheet.id);
    expect(refreshed.rowOrder).toEqual(worksheet.rowOrder);
    expect(refreshed.columnOrder).toEqual(worksheet.columnOrder);
    expect(refreshed.viewBinding).toMatchObject({
      viewId: "contacts_flat",
      showDocuments: true,
      showCategories: true,
      rootCategoryPath: ["Customers"],
      lastViewCursor: "cursor_2",
    });
    expect(Object.values(refreshed.cellsById).map((cell) => cell.value)).toContainEqual({ kind: "string", text: "Alice" });
    expect(refreshed.columnOrder).toHaveLength(2);
    expect(refreshed.cellsById[`${refreshed.rowOrder[1]}:${refreshed.columnOrder[0]}`]?.style).toMatchObject({
      backgroundColor: "#f3f4f6",
    });
    expect(refreshed.cellsById[`${refreshed.rowOrder[2]}:${refreshed.columnOrder[0]}`]).toBeUndefined();
    expect(refreshed.cellsById[`${refreshed.rowOrder[2]}:${refreshed.columnOrder[1]}`]?.value).toEqual({
      kind: "string",
      text: "Alice",
    });
  });

  it("materializes view count formulas from category and document entry metadata", () => {
    const worksheet = buildViewSheetWorksheet({
      settings: settings(),
      view: {
        ...viewDefinition(),
        columns: [
          ...viewDefinition().columns,
          {
            id: "count",
            name: "count",
            title: "Count",
            role: "display",
            expression: {
              mode: "formula",
              expression: { kind: "operation", op: "descendantDocumentCount", args: [] },
            },
            sorting: "none",
            totalMode: "none",
            hidden: false,
        }, {
          id: "siblings",
          name: "siblings",
          title: "Siblings",
          role: "display",
          expression: {
            mode: "formula",
            expression: { kind: "operation", op: "siblingCount", args: [] },
          },
          sorting: "none",
          totalMode: "none",
          hidden: false,
          },
        ],
      },
      rootCategoryPath: ["Customers"],
      entries: entries(),
      lastViewCursor: "cursor_1",
      now: () => new Date("2026-05-20T12:00:00.000Z"),
    });

    const countColumnId = worksheet.columnOrder[2];
    const siblingColumnId = worksheet.columnOrder[3];

    expect(worksheet.cellsById[`${worksheet.rowOrder[1]}:${countColumnId}`]?.value).toEqual({
      kind: "number",
      value: 1,
    });
    expect(worksheet.cellsById[`${worksheet.rowOrder[1]}:${siblingColumnId}`]?.value).toEqual({
      kind: "number",
      value: 1,
    });
    expect(worksheet.cellsById[`${worksheet.rowOrder[2]}:${countColumnId}`]?.value).toEqual({
      kind: "number",
      value: 0,
    });
  });

  it("opens a configured view with selected inclusion options and disposes it", async () => {
    const dispose = vi.fn<() => Promise<void>>().mockResolvedValue();
    const navigator = {
      expandAll: vi.fn().mockResolvedValue(undefined),
      entriesForward: vi.fn().mockResolvedValue({ entries: entries(), nextPosition: null, hasMore: false }),
      getViewCursor: vi.fn().mockResolvedValue("cursor_1"),
      dispose,
    } as unknown as MindooDBAppViewNavigator;
    const openViewNavigator = vi.fn().mockResolvedValue(navigator);

    await materializeViewSheet({
      settings: settings({ rootCategoryPathInput: "Customers\\ACME", showCategories: false }),
      view: viewDefinition(),
      openViewNavigator,
    });

    expect(openViewNavigator).toHaveBeenCalledWith("contacts_flat", {
      includeCategories: false,
      includeDocuments: true,
      hideEmptyCategories: true,
      rootCategoryPath: ["Customers", "ACME"],
    });
    expect(navigator.expandAll).toHaveBeenCalledOnce();
    expect(dispose).toHaveBeenCalledOnce();
  });
});

function settings(overrides: Partial<ViewSheetSettings> = {}): ViewSheetSettings {
  return {
    title: "Contacts",
    viewId: "contacts_flat",
    showDocuments: true,
    showCategories: true,
    rootCategoryPathInput: "",
    ...overrides,
  };
}

function viewDefinition(): MindooDBAppResolvedViewDefinition {
  return {
    id: "contacts_flat",
    description: "Contacts",
    categorizationStyle: "category_then_document",
    previewMode: "tree",
    sources: [],
    filter: { mode: "rules", match: "all", rules: [] },
    columns: [{
      id: "tags",
      name: "tags",
      title: "Tags",
      role: "category",
      expression: { mode: "field", field: "tags" },
      sorting: "ascending",
      totalMode: "none",
      hidden: false,
    }, {
      id: "name",
      name: "name",
      title: "Name",
      role: "display",
      expression: { mode: "field", field: "name" },
      sorting: "ascending",
      totalMode: "none",
      hidden: false,
    }, {
      id: "hidden",
      name: "hidden",
      title: "Hidden",
      role: "display",
      expression: { mode: "field", field: "hidden" },
      sorting: "none",
      totalMode: "none",
      hidden: true,
    }],
  };
}

function entries(): MindooDBAppViewEntry[] {
  return [{
    key: "cat_1",
    kind: "category",
    origin: "contacts",
    docId: null,
    level: 0,
    parentKey: null,
    categoryPath: ["Customers"],
    categoryValue: "Customers",
    columnValues: { tags: "Customers", name: "Customers" },
    descendantDocumentCount: 1,
    childCategoryCount: 0,
    childDocumentCount: 1,
    siblingCount: 1,
    position: "cat_1",
    expanded: true,
    selected: false,
    isVisible: true,
  }, {
    key: "doc_1",
    kind: "document",
    origin: "contacts",
    docId: "contact_1",
    level: 1,
    parentKey: "cat_1",
    categoryPath: ["Customers"],
    categoryValue: null,
    columnValues: { tags: "Customers", name: "Alice" },
    descendantDocumentCount: 0,
    childCategoryCount: 0,
    childDocumentCount: 0,
    siblingCount: 1,
    position: "doc_1",
    expanded: false,
    selected: false,
    isVisible: true,
  }];
}
