/**
 * `/__haven-test/`: the app in an iframe, launched by the SDK's mock Haven
 * with the databases and permissions from `haven-app.json`. Mock data only;
 * it is gone on reload. Served by `pnpm dev`, never part of a production
 * build. `window.__havenTestHost` scripts the host.
 */
import type { MindooDBAppDefinition } from 'mindoodb-app-sdk'
import { mockDatabasesFromDefinition, mountHavenTestHost } from 'mindoodb-app-sdk/testing'

async function start() {
  const response = await fetch(new URL('../haven-app.json', window.location.href))
  const definition = (await response.json()) as MindooDBAppDefinition
  mountHavenTestHost({
    appUrl: '../',
    title: definition.label,
    launchContext: {
      appId: definition.appId,
      launchParameters: { ...definition.launchParameters },
      ...(definition.version ? { appVersion: definition.version } : {}),
      ...(definition.defaultLaunchDatabaseId ? { preferredDatabaseId: definition.defaultLaunchDatabaseId } : {}),
    },
    // Real Automerge documents: saves merge at their baseHeads as in Haven, and
    // __havenTestHost.applyRemoteUpdate plays a second device.
    databases: mockDatabasesFromDefinition(definition, { teamgrid: [] }, { automerge: true }),
  })
}

void start()
