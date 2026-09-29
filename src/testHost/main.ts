/**
 * The app's test URL, `/__haven-test/`: the real app in an iframe, launched by a mock
 * Haven with the databases, permissions and launch parameters from `haven-app.json`,
 * plus a control panel (theme, host focus, notifications, bridge requests). Use it for
 * Playwright, for AI agents driving the UI, and for clicking through the app without
 * installing it in a real Haven. `window.__havenTestHost` scripts the host.
 *
 * It is served by `pnpm dev` and only built with `HAVEN_TEST_HOST=1` (preview
 * deployments), never for the production deploy: the public URL shows the landing page.
 */
import type { MindooDBAppDefinition } from "mindoodb-app-sdk";
import { mockDatabasesFromDefinition, mountHavenTestHost } from "mindoodb-app-sdk/testing";

import { seedDocuments } from "@/testHost/seed";

async function start() {
  const response = await fetch(new URL("../haven-app.json", window.location.href));
  const definition = (await response.json()) as MindooDBAppDefinition;

  mountHavenTestHost({
    appUrl: "../",
    title: definition.label,
    launchContext: {
      appId: definition.appId,
      appVersion: definition.version,
      launchParameters: { ...definition.launchParameters },
      preferredDatabaseId: definition.defaultLaunchDatabaseId,
    },
    databases: mockDatabasesFromDefinition(definition, seedDocuments),
  });
}

void start();
