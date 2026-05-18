/// <reference lib="webworker" />

import { CacheableResponsePlugin } from "workbox-cacheable-response";
import { ExpirationPlugin } from "workbox-expiration";
import { cleanupOutdatedCaches, createHandlerBoundToURL, precacheAndRoute } from "workbox-precaching";
import { NavigationRoute, registerRoute } from "workbox-routing";
import { NetworkFirst, StaleWhileRevalidate } from "workbox-strategies";

declare let self: ServiceWorkerGlobalScope;

const TEAMGRID_RUNTIME_ASSET_CACHE = "mindoodb-teamgrid-runtime-assets-v1";
const TEAMGRID_NAVIGATION_SHELL_CACHE = "mindoodb-teamgrid-navigation-shell-v1";

precacheAndRoute(self.__WB_MANIFEST);
cleanupOutdatedCaches();

// Keep the newly installed worker waiting until the app asks users to reload.
self.addEventListener("install", () => {});

function isSkipWaitingMessage(data: unknown): data is { type: "SKIP_WAITING" } {
  return typeof data === "object"
    && data !== null
    && "type" in data
    && data.type === "SKIP_WAITING";
}

function readClientId(source: ExtendableMessageEvent["source"]) {
  if (!source || !("id" in source) || typeof source.id !== "string") {
    return null;
  }
  return source.id;
}

async function isTrustedClientSource(source: ExtendableMessageEvent["source"]) {
  const clientId = readClientId(source);
  if (!clientId) {
    return false;
  }
  const client = await self.clients.get(clientId);
  return Boolean(client && new URL(client.url).origin === self.location.origin);
}

self.onmessage = (event) => {
  if (!event.isTrusted || event.origin !== self.location.origin) {
    return;
  }
  if (!isSkipWaitingMessage(event.data)) {
    return;
  }
  event.waitUntil((async () => {
    if (await isTrustedClientSource(event.source)) {
      await self.skipWaiting();
    }
  })());
};

self.addEventListener("activate", (event) => {
  event.waitUntil(self.clients.claim());
});

const navigationNetworkFirst = new NetworkFirst({
  cacheName: TEAMGRID_NAVIGATION_SHELL_CACHE,
  networkTimeoutSeconds: 5,
  plugins: [
    new CacheableResponsePlugin({
      statuses: [200],
    }),
    new ExpirationPlugin({
      maxEntries: 16,
    }),
  ],
});
const navigationFallback = createHandlerBoundToURL("/index.html");

registerRoute(
  new NavigationRoute(async (options) => {
    try {
      return await navigationNetworkFirst.handle(options) ?? await navigationFallback(options);
    } catch {
      return navigationFallback(options);
    }
  }),
);

registerRoute(
  ({ request, url }) =>
    request.method === "GET"
    && url.origin === self.location.origin
    && (
      [
        "font",
        "image",
        "manifest",
      ].includes(request.destination)
      || url.pathname.endsWith(".wasm")
    ),
  new StaleWhileRevalidate({
    cacheName: TEAMGRID_RUNTIME_ASSET_CACHE,
    plugins: [
      new CacheableResponsePlugin({
        statuses: [200],
      }),
      new ExpirationPlugin({
        maxEntries: 160,
      }),
    ],
  }),
);
