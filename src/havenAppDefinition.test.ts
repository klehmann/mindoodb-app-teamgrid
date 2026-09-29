/**
 * `public/haven-app.json` is generated from the app store catalog
 * (`mindoodb-appstore/scripts/export-app-listing.mjs`). Haven reads it when the app is
 * installed from its URL, and the landing page reads its `listing` when the URL is
 * opened directly. One invalid value makes the whole definition unreadable — the
 * landing page then shows a generic text and the install fails — so it gets a test.
 */
import { existsSync, readFileSync } from "node:fs";
import { resolve } from "node:path";

import { validateMindooDBAppDefinition } from "mindoodb-app-sdk";
import { describe, expect, it } from "vitest";

// Resolved from the project root rather than `import.meta.url`: the jsdom environment
// does not give this module a `file:` URL.
const publicDir = resolve(process.cwd(), "public");
const raw: unknown = JSON.parse(readFileSync(resolve(publicDir, "haven-app.json"), "utf8"));

describe("haven-app.json", () => {
  it("is a valid app definition", () => {
    const { definition, errors } = validateMindooDBAppDefinition(raw);
    expect(errors).toEqual([]);
    expect(definition).not.toBeNull();
  });

  it("carries the store listing", () => {
    const { definition } = validateMindooDBAppDefinition(raw);
    expect(definition?.listing?.summary).toBeTruthy();
    expect(definition?.listing?.description).toBeTruthy();
  });

  it("ships every listing image it references", () => {
    const { definition } = validateMindooDBAppDefinition(raw);
    const files = [
      definition?.listing?.icon,
      ...(definition?.listing?.screenshots ?? []).map((screenshot) => screenshot.file),
    ].filter((file): file is string => Boolean(file));
    for (const file of files) {
      expect(existsSync(resolve(publicDir, file)), file).toBe(true);
    }
  });
});
