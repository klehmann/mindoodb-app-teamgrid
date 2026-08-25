import { describe, expect, it } from "vitest";
import { detectBrowserUiLanguage, mapLocaleToUiLanguage, SUPPORTED_UI_LANGUAGES } from "./index";
import deMessages from "./locales/de.json";
import enMessages from "./locales/en.json";
import esMessages from "./locales/es.json";
import frMessages from "./locales/fr.json";
import itMessages from "./locales/it.json";
import nbMessages from "./locales/nb.json";
import nlMessages from "./locales/nl.json";
import plMessages from "./locales/pl.json";

const LOCALE_MESSAGES: Record<string, unknown> = {
  de: deMessages,
  en: enMessages,
  es: esMessages,
  fr: frMessages,
  it: itMessages,
  nb: nbMessages,
  nl: nlMessages,
  pl: plMessages,
};

/** Recursively collects dotted key paths from a nested translation object. */
function flattenKeys(value: unknown, prefix = ""): string[] {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return [prefix];
  }
  return Object.entries(value as Record<string, unknown>).flatMap(([key, child]) =>
    flattenKeys(child, prefix ? `${prefix}.${key}` : key),
  );
}

describe("mapLocaleToUiLanguage", () => {
  it("maps German-speaking country locales to de", () => {
    expect(mapLocaleToUiLanguage("de-DE")).toBe("de");
    expect(mapLocaleToUiLanguage("de-AT")).toBe("de");
    expect(mapLocaleToUiLanguage("de-CH")).toBe("de");
  });

  it("maps French and Spanish locales directly", () => {
    expect(mapLocaleToUiLanguage("fr-FR")).toBe("fr");
    expect(mapLocaleToUiLanguage("es-ES")).toBe("es");
  });

  it("maps the remaining supported single-country locales directly", () => {
    expect(mapLocaleToUiLanguage("nl-NL")).toBe("nl");
    expect(mapLocaleToUiLanguage("it-IT")).toBe("it");
    expect(mapLocaleToUiLanguage("pl-PL")).toBe("pl");
    expect(mapLocaleToUiLanguage("en-GB")).toBe("en");
  });

  it("maps Norwegian Bokmål under both the nb and legacy no language tags", () => {
    expect(mapLocaleToUiLanguage("nb-NO")).toBe("nb");
    expect(mapLocaleToUiLanguage("no-NO")).toBe("nb");
  });

  it("falls back to English for unsupported or missing locales", () => {
    expect(mapLocaleToUiLanguage("pt-BR")).toBe("en");
    expect(mapLocaleToUiLanguage(undefined)).toBe("en");
    expect(mapLocaleToUiLanguage("")).toBe("en");
  });
});

describe("detectBrowserUiLanguage", () => {
  it("returns the first supported browser language in preference order", () => {
    expect(detectBrowserUiLanguage(["de-DE", "en-US"])).toBe("de");
    expect(detectBrowserUiLanguage(["fr-CA", "fr-FR"])).toBe("fr");
    expect(detectBrowserUiLanguage(["en-GB"])).toBe("en");
  });

  it("skips unsupported languages instead of collapsing them to English early", () => {
    expect(detectBrowserUiLanguage(["pt-BR", "fr-FR"])).toBe("fr");
    expect(detectBrowserUiLanguage(["sv-SE", "it-IT", "de-DE"])).toBe("it");
    expect(detectBrowserUiLanguage(["es-ES", "fr-FR"])).toBe("es");
  });

  it("maps Norwegian tags (nb/nn/no) to nb", () => {
    expect(detectBrowserUiLanguage(["nb-NO"])).toBe("nb");
    expect(detectBrowserUiLanguage(["no"])).toBe("nb");
    expect(detectBrowserUiLanguage(["nn-NO"])).toBe("nb");
  });

  it("falls back to English when nothing is supported or the list is empty", () => {
    expect(detectBrowserUiLanguage(["sv-SE", "pt-BR"])).toBe("en");
    expect(detectBrowserUiLanguage([])).toBe("en");
  });
});

describe("locale message files", () => {
  it("ship a message file for every supported UI language", () => {
    for (const language of SUPPORTED_UI_LANGUAGES) {
      expect(LOCALE_MESSAGES[language]).toBeDefined();
    }
  });

  it("keep the exact same set of translation keys across all languages", () => {
    const [baseLanguage, ...otherLanguages] = SUPPORTED_UI_LANGUAGES;
    const baseKeys = new Set(flattenKeys(LOCALE_MESSAGES[baseLanguage]));
    expect(baseKeys.size).toBeGreaterThan(0);

    for (const language of otherLanguages) {
      const keys = new Set(flattenKeys(LOCALE_MESSAGES[language]));
      const missing = [...baseKeys].filter((key) => !keys.has(key));
      const extra = [...keys].filter((key) => !baseKeys.has(key));
      expect({ language, missing, extra }).toEqual({ language, missing: [], extra: [] });
    }
  });

  it("never ships an empty translation value", () => {
    for (const language of SUPPORTED_UI_LANGUAGES) {
      const messages = LOCALE_MESSAGES[language] as Record<string, unknown>;
      const emptyKeys = flattenKeys(messages).filter((key) => {
        const parts = key.split(".");
        let node: unknown = messages;
        for (const part of parts) {
          node = (node as Record<string, unknown>)?.[part];
        }
        if (Array.isArray(node)) {
          return node.some((entry) => typeof entry !== "string" || entry.trim() === "");
        }
        return typeof node === "string" && node.trim() === "";
      });
      expect({ language, emptyKeys }).toEqual({ language, emptyKeys: [] });
    }
  });
});
