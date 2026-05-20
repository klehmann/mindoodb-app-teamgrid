import { describe, expect, it } from "vitest";

import { buildOpenCategoryTree, createOpenViewDefinition } from "@/features/document/lib/viewOpen";
import type { MindooDBAppViewEntry } from "mindoodb-app-sdk";

describe("createOpenViewDefinition", () => {
  it("builds fixed view definitions for each spreadsheet type", () => {
    const allDefinition = createOpenViewDefinition("all");
    expect(allDefinition).toMatchObject({
      id: "teamgrid-open-tags-all-v1",
    });
    expect(allDefinition).not.toHaveProperty("filter");
    expect(createOpenViewDefinition("noTemplates")).toMatchObject({
      id: "teamgrid-open-tags-noTemplates-v1",
      filter: {
        mode: "expression",
        expression: {
          kind: "operation",
          op: "neq",
          args: [
            { kind: "field", path: "istemplate" },
            { kind: "literal", value: true },
          ],
        },
      },
    });
    expect(createOpenViewDefinition("onlyTemplates")).toMatchObject({
      id: "teamgrid-open-tags-onlyTemplates-v1",
      filter: {
        mode: "expression",
        expression: {
          kind: "operation",
          op: "eq",
          args: [
            { kind: "field", path: "istemplate" },
            { kind: "literal", value: true },
          ],
        },
      },
    });
  });

  it("uses descendant document counts for category badges", () => {
    const { roots } = buildOpenCategoryTree([
      {
        key: "category:Work",
        kind: "category",
        categoryPath: ["Work"],
        columnValues: {},
        descendantCount: 8,
        descendantDocumentCount: 3,
      } as MindooDBAppViewEntry,
    ], 3);

    expect(roots[0]?.children[0]?.count).toBe(3);
  });
});
