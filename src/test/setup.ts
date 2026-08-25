/**
 * Vitest setup that polyfills DOM APIs jsdom does not implement out of the box
 * but PrimeVue and Vue Test Utils rely on. Keep this file tiny and only patch
 * APIs that are read by libraries we use; mocks for application code belong in
 * the individual test files.
 */
if (typeof window !== "undefined" && typeof window.matchMedia !== "function") {
  Object.defineProperty(window, "matchMedia", {
    writable: true,
    value: (query: string) => ({
      matches: false,
      media: query,
      onchange: null,
      addListener: () => {},
      removeListener: () => {},
      addEventListener: () => {},
      removeEventListener: () => {},
      dispatchEvent: () => false,
    }),
  });
}

import { config } from "@vue/test-utils";
import { i18n, setUiLanguage } from "@/i18n";

setUiLanguage("en");
config.global.plugins.push(i18n);
