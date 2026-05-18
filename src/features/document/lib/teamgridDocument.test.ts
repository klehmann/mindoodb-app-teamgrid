import { describe, expect, it } from "vitest";
import {
  createTeamGridDocument,
  getFirstVisibleWorksheet,
  isTeamGridEnvelope,
  migrateTeamGridDocument,
  normalizeTags,
  readTags,
} from "@/features/document/lib/teamgridDocument";

describe("teamgrid document schema", () => {
  it("creates one workbook with one visible worksheet tab", () => {
    const envelope = createTeamGridDocument("Planning");

    expect(envelope.subject).toBe("Planning");
    expect(envelope.tags).toEqual([]);
    expect(envelope.form).toBe("teamgrid");
    expect(isTeamGridEnvelope(envelope)).toBe(true);
    expect(envelope.teamgrid.workbook.worksheetOrder).toHaveLength(1);
    expect("title" in envelope.teamgrid.workbook).toBe(false);
    expect(getFirstVisibleWorksheet(envelope.teamgrid)?.title).toBe("Sheet 1");
  });

  it("migrates unknown document data into a Teamgrid envelope", () => {
    const migrated = migrateTeamGridDocument({ subject: "Legacy title" });

    expect(migrated.subject).toBe("Legacy title");
    expect(migrated.tags).toEqual([]);
    expect(isTeamGridEnvelope(migrated)).toBe(true);
  });

  it("normalizes top-level tags and drops duplicate workbook titles during migration", () => {
    const envelope = createTeamGridDocument("Planning", [" Work\\Q1 ", "Work\\Q1", "", "Personal"]);
    const legacyEnvelope = {
      ...envelope,
      teamgrid: {
        ...envelope.teamgrid,
        workbook: {
          ...envelope.teamgrid.workbook,
          title: "Legacy duplicate title",
        },
      },
    };

    const migrated = migrateTeamGridDocument(legacyEnvelope);

    expect(migrated.subject).toBe("Planning");
    expect(migrated.tags).toEqual(["Work\\Q1", "Personal"]);
    expect("title" in migrated.teamgrid.workbook).toBe(false);
  });

  it("reads and normalizes tags defensively", () => {
    expect(readTags({ tags: ["A", " A ", null, "B\\C", ""] })).toEqual(["A", "B\\C"]);
    expect(normalizeTags("A")).toEqual([]);
  });

  it("uses tombstoned worksheets only as historical state", () => {
    const envelope = createTeamGridDocument();
    const worksheet = getFirstVisibleWorksheet(envelope.teamgrid);
    expect(worksheet).not.toBeNull();

    worksheet!.deletedAt = new Date().toISOString();

    expect(getFirstVisibleWorksheet(envelope.teamgrid)).toBeNull();
  });
});
