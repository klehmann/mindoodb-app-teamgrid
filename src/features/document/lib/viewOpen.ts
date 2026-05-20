/**
 * View-backed model for the File/Open dialog.
 *
 * Instead of pulling a flat list of documents with `documents.list()` and
 * building a tag tree client-side, Teamgrid drives the Open dialog from a
 * MindooDB virtual view. The view declares one category column over the
 * top-level `tags` array, which gives us three useful behaviours for free:
 *
 * - Array fan-out: a document tagged `["Work", "Customer\\ACME"]` appears
 *   both under the `Work` category and under the nested `Customer > ACME`
 *   category, because the view engine iterates array category values.
 * - Backslash hierarchy: a single string `"Work\\Planning"` becomes a
 *   nested `Work > Planning` path. The host splits on backslash before
 *   placing the document into the tree.
 * - Cheap counts: each navigator entry reports `descendantDocumentCount`,
 *   so the tree shows accurate counts per category without us re-counting.
 *
 * The category column is `"tags"`, sorted ascending, and the `subject`
 * display column is sorted ascending too so the document pane orders
 * alphabetically by title regardless of insertion order.
 */
import {
  createViewLanguage,
  type MindooDBAppViewDefinition,
  type MindooDBAppViewEntry,
  type MindooDBAppViewNavigator,
} from "mindoodb-app-sdk";

import { normalizeTags } from "@/features/document/lib/teamgridDocument";

/** Synthetic key for the root "All spreadsheets" node shown above the tag tree. */
export const ALL_SPREADSHEETS_NODE_KEY = "all";

export type OpenSpreadsheetType = "all" | "noTemplates" | "onlyTemplates";

/** One node in the category tree rendered on the left side of File/Open. */
export interface OpenCategoryNode {
  /** Navigator entry key, or {@link ALL_SPREADSHEETS_NODE_KEY} for the root. */
  key: string;
  /** Display label, derived from the last segment of the category path. */
  label: string;
  /** Number of documents in this subtree, used as a badge in the tree. */
  count: number;
  /** Nested categories sorted by the view's ascending sort. */
  children: OpenCategoryNode[];
}

/** One document row rendered on the right side of File/Open. */
export interface OpenDocumentRow {
  /** Navigator entry key; useful when the same document appears in many tags. */
  key: string;
  /** Stable document id used by `openDocument` to load the spreadsheet. */
  id: string;
  /** Display title, falling back to the document id when no subject is set. */
  title: string;
  /** Normalized tag list shown beneath the title. */
  tags: string[];
  /** Comma-joined tag string used as the secondary line, or "Untagged". */
  detail: string;
}

/**
 * Build the view definition that drives the File/Open category tree.
 *
 * Uses the SDK's view-language helper to produce strongly-typed expressions:
 * `v.field("tags")` for the category column and `v.field("subject")` for the
 * display column. The fixed `id` lets the host cache the view's index across
 * dialog opens.
 */
export function createOpenViewDefinition(type: OpenSpreadsheetType = "noTemplates"): MindooDBAppViewDefinition {
  const v = createViewLanguage();
  const definition: MindooDBAppViewDefinition = {
    id: `teamgrid-open-tags-${type}-v1`,
    title: "Teamgrid spreadsheets by tag",
    defaultExpand: "expanded",
    columns: [
      {
        name: "tags",
        title: "Tags",
        role: "category",
        expression: v.field("tags"),
        sorting: "ascending",
      },
      {
        name: "subject",
        title: "Title",
        role: "display",
        expression: v.field("subject"),
        sorting: "ascending",
      },
    ],
  };
  if (type === "noTemplates") {
    definition.filter = {
      mode: "expression",
      expression: v.neq(v.field("istemplate"), true),
    };
  } else if (type === "onlyTemplates") {
    definition.filter = {
      mode: "expression",
      expression: v.eq(v.field("istemplate"), true),
    };
  }
  return definition;
}

/**
 * Drain every entry from a navigator in forward order.
 *
 * The page size is intentionally generous because the dialog needs both the
 * full category tree and the document list to render the synthetic "All
 * spreadsheets" view. For very large databases the dialog should instead
 * lazy-load `childDocuments(key)` per selected category, which this module
 * exposes via {@link mapDocumentEntries}.
 */
export async function collectNavigatorEntries(navigator: MindooDBAppViewNavigator) {
  const entries: MindooDBAppViewEntry[] = [];
  let startPosition: string | null = null;
  do {
    const page = await navigator.entriesForward({ limit: 1000, startPosition });
    entries.push(...page.entries);
    startPosition = page.nextPosition;
  } while (startPosition);
  return entries;
}

/**
 * Re-shape the navigator's flat list of category entries into a tree.
 *
 * Each entry already knows its `parentKey`, so this is a single pass: create
 * a node per entry, then attach it under the matching parent (or the
 * synthetic root). The result preserves the navigator's sort order because
 * we visit entries in the order the navigator returned them.
 */
export function buildOpenCategoryTree(categoryEntries: MindooDBAppViewEntry[], documentCount: number) {
  const root: OpenCategoryNode = {
    key: ALL_SPREADSHEETS_NODE_KEY,
    label: "All spreadsheets",
    count: documentCount,
    children: [],
  };
  const nodesByKey = new Map<string, OpenCategoryNode>([[root.key, root]]);
  for (const entry of categoryEntries) {
    const node: OpenCategoryNode = {
      key: entry.key,
      label: readCategoryLabel(entry),
      count: entry.descendantDocumentCount ?? 0,
      children: [],
    };
    nodesByKey.set(node.key, node);
  }
  for (const entry of categoryEntries) {
    const node = nodesByKey.get(entry.key);
    if (!node) {
      continue;
    }
    const parent = entry.parentKey ? nodesByKey.get(entry.parentKey) : root;
    parent?.children.push(node);
  }
  return { roots: [root], nodesByKey };
}

/**
 * Collapse the navigator's fan-out so each document appears at most once.
 *
 * Documents with multiple tags appear once per category in the raw entry
 * stream. The "All spreadsheets" pane only wants each document once, so we
 * dedupe by document id and keep the first occurrence.
 */
export function dedupeDocumentEntries(entries: MindooDBAppViewEntry[]) {
  const rowsById = new Map<string, OpenDocumentRow>();
  for (const entry of entries) {
    const row = mapDocumentEntry(entry);
    if (row && !rowsById.has(row.id)) {
      rowsById.set(row.id, row);
    }
  }
  return Array.from(rowsById.values());
}

/**
 * Map document-kind navigator entries to UI rows without deduping.
 *
 * Use this when showing the children of a specific category, where each
 * entry is already scoped to one category path and duplicates can't occur.
 */
export function mapDocumentEntries(entries: MindooDBAppViewEntry[]) {
  return entries
    .map(mapDocumentEntry)
    .filter((row): row is OpenDocumentRow => row !== null);
}

function mapDocumentEntry(entry: MindooDBAppViewEntry) {
  if (entry.kind !== "document" || !entry.docId) {
    return null;
  }
  const tags = normalizeTags(entry.columnValues.tags);
  return {
    key: entry.key,
    id: entry.docId,
    title: readTitle(entry),
    tags,
    detail: tags.join(", ") || "Untagged",
  };
}

function readTitle(entry: MindooDBAppViewEntry) {
  const value = entry.columnValues.subject;
  return typeof value === "string" && value.trim() ? value : entry.docId ?? "Untitled spreadsheet";
}

function readCategoryLabel(entry: MindooDBAppViewEntry) {
  const value = entry.categoryPath.at(-1);
  return value == null || value === "" ? "Untagged" : String(value);
}
