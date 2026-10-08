/**
 * `/__haven-test/`: the app in an iframe, launched by the SDK's mock Haven
 * with the databases and permissions from `haven-app.json`. Mock data only;
 * it is gone on reload. Served by `pnpm dev`, never part of a production
 * build. `window.__havenTestHost` scripts the host.
 */
import type { MindooDBAppDefinition } from 'mindoodb-app-sdk'
import { mockDatabasesFromDefinition, mountHavenTestHost } from 'mindoodb-app-sdk/testing'

import { CONTACTS, CONTACTS_VIEW, openContactsView } from './contacts-view'

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
      // A view configured as a data source, for view sheets.
      views: [CONTACTS_VIEW],
    },
    // Real Automerge documents: saves merge at their baseHeads as in Haven, and
    // __havenTestHost.applyRemoteUpdate plays a second device.
    databases: mockDatabasesFromDefinition(definition, { teamgrid: CONTACTS }, { automerge: true }).map((database) =>
      database.info.id === 'teamgrid'
        ? { ...database, methods: { ...database.methods, views: { open: (_id, options) => openContactsView(options) } } }
        : database,
    ),
  })
}

void start()
