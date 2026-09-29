import type { MockMindooDBAppDatabaseDefinition } from "mindoodb-app-sdk/testing";

/**
 * Documents the test host starts with, per logical database id from haven-app.json.
 * Mock data only: it never reaches a real Haven and is gone on reload.
 */
export const seedDocuments: Record<string, MockMindooDBAppDatabaseDefinition["documents"]> = {
  teamgrid: [],
};
