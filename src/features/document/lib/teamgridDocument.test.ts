import { afterEach, describe, expect, it, vi } from "vitest";
import {
  createTeamGridDocument,
  DEFAULT_WORKSHEET_COLUMNS,
  DEFAULT_WORKSHEET_ROWS,
  getFirstVisibleWorksheet,
  isTeamGridEnvelope,
  migrateTeamGridDocument,
  normalizeTeamGridLocale,
  normalizeTags,
  readTags,
  readIsTemplate,
} from "@/features/document/lib/teamgridDocument";

describe("teamgrid document schema", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("creates one workbook with one visible worksheet tab", () => {
    const envelope = createTeamGridDocument("Planning");

    expect(envelope.subject).toBe("Planning");
    expect(envelope.tags).toEqual([]);
    expect(envelope.istemplate).toBe(false);
    expect(envelope.form).toBe("teamgrid");
    expect(isTeamGridEnvelope(envelope)).toBe(true);
    expect(envelope.teamgrid.workbook.worksheetOrder).toHaveLength(1);
    expect("title" in envelope.teamgrid.workbook).toBe(false);
    const worksheet = getFirstVisibleWorksheet(envelope.teamgrid);
    expect(worksheet?.title).toBe("Sheet 1");
    expect(worksheet?.rowOrder).toHaveLength(DEFAULT_WORKSHEET_ROWS);
    expect(worksheet?.columnOrder).toHaveLength(DEFAULT_WORKSHEET_COLUMNS);
    expect(worksheet?.chartOrder).toEqual([]);
    expect(worksheet?.chartsById).toEqual({});
  });

  it("initializes new documents from the browser locale when available", () => {
    vi.stubGlobal("navigator", {
      languages: ["de-DE", "en-US"],
      language: "de-DE",
    });

    expect(createTeamGridDocument().teamgrid.settings.locale).toBe("de-DE");
  });

  it("normalizes browser base language locales to supported regional choices", () => {
    vi.stubGlobal("navigator", {
      languages: ["de"],
      language: "de",
    });

    expect(createTeamGridDocument().teamgrid.settings.locale).toBe("de-DE");
    expect(normalizeTeamGridLocale("de_DE")).toBe("de-DE");
    expect(normalizeTeamGridLocale("fr")).toBe("fr-FR");
  });

  it("migrates unknown document data into a Teamgrid envelope", () => {
    const migrated = migrateTeamGridDocument({ subject: "Legacy title" });

    expect(migrated.subject).toBe("Legacy title");
    expect(migrated.tags).toEqual([]);
    expect(migrated.istemplate).toBe(false);
    expect(isTeamGridEnvelope(migrated)).toBe(true);
  });

  it("preserves the top-level template flag during migration", () => {
    const envelope = createTeamGridDocument("Planning", [], "en-US", true);

    expect(readIsTemplate(envelope as unknown as Record<string, unknown>)).toBe(true);
    expect(migrateTeamGridDocument(envelope as unknown as Record<string, unknown>).istemplate).toBe(true);
    expect(migrateTeamGridDocument({ subject: "Template", istemplate: true }).istemplate).toBe(true);
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
    expect(getFirstVisibleWorksheet(migrated.teamgrid)?.chartOrder).toEqual([]);
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
