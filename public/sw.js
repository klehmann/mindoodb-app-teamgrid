// TeamGrid 1.x registered a service worker (Workbox offline shell) at this
// URL. TeamGrid 2 has none; browsers that still run the old one fetch this
// file on their next update check. It clears the old caches, unregisters
// itself and reloads open tabs, so they load the current app from the network.
self.addEventListener('install', () => self.skipWaiting())
self.addEventListener('activate', (event) => {
  event.waitUntil(
    (async () => {
      for (const name of await caches.keys()) await caches.delete(name)
      await self.registration.unregister()
      for (const client of await self.clients.matchAll({ type: 'window' })) client.navigate(client.url)
    })(),
  )
})
