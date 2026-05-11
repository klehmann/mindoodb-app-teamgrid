import { describe, expect, it } from "vitest";
import {
  createTeamGridDocument,
  getFirstVisibleWorksheet,
  isTeamGridEnvelope,
  migrateTeamGridDocument,
} from "@/lib/teamgridDocument";

describe("teamgrid document schema", () => {
  it("creates one workbook with one visible worksheet tab", () => {
    const envelope = createTeamGridDocument("Planning");

    expect(envelope.subject).toBe("Planning");
    expect(envelope.form).toBe("teamgrid");
    expect(isTeamGridEnvelope(envelope)).toBe(true);
    expect(envelope.teamgrid.workbook.worksheetOrder).toHaveLength(1);
    expect(getFirstVisibleWorksheet(envelope.teamgrid)?.title).toBe("Sheet 1");
  });

  it("migrates unknown document data into a Teamgrid envelope", () => {
    const migrated = migrateTeamGridDocument({ subject: "Legacy title" });

    expect(migrated.subject).toBe("Legacy title");
    expect(isTeamGridEnvelope(migrated)).toBe(true);
  });

  it("uses tombstoned worksheets only as historical state", () => {
    const envelope = createTeamGridDocument();
    const worksheet = getFirstVisibleWorksheet(envelope.teamgrid);
    expect(worksheet).not.toBeNull();

    worksheet!.deletedAt = new Date().toISOString();

    expect(getFirstVisibleWorksheet(envelope.teamgrid)).toBeNull();
  });
});
