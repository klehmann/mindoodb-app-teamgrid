import { createI18n } from "vue-i18n";
import de from "./locales/de.json";
import en from "./locales/en.json";
import es from "./locales/es.json";
import fr from "./locales/fr.json";
import it from "./locales/it.json";
import nb from "./locales/nb.json";
import nl from "./locales/nl.json";
import pl from "./locales/pl.json";

/**
 * UI languages shipped for TeamGrid — keep in sync with MindooDB Haven
 * (`de`, `en`, `es`, `fr`, `it`, `nb`, `nl`, `pl`). Phrases live in
 * `./locales/<code>.json`; English is the fallback for missing keys.
 */
export const SUPPORTED_UI_LANGUAGES = ["de", "en", "es", "fr", "it", "nb", "nl", "pl"] as const;
export type SupportedUiLanguage = (typeof SUPPORTED_UI_LANGUAGES)[number];
export const DEFAULT_UI_LANGUAGE: SupportedUiLanguage = "en";

/**
 * Maps a BCP-47 tag (e.g. `de-AT`, `nb-NO`, `no-NO`) to a shipped UI language.
 * Falls back to English for anything unrecognized so the app never silently
 * renders raw translation keys.
 */
export function mapLocaleToUiLanguage(bcp47: string | undefined): SupportedUiLanguage {
  const language = (bcp47?.split("-")[0] ?? "").toLowerCase();
  if (language === "no") return "nb";
  if ((SUPPORTED_UI_LANGUAGES as readonly string[]).includes(language)) {
    return language as SupportedUiLanguage;
  }
  return "en";
}

/** The browser's preferred languages, most-preferred first (empty outside a browser). */
export function browserLanguages(): string[] {
  if (typeof navigator === "undefined") return [];
  const list =
    Array.isArray(navigator.languages) && navigator.languages.length > 0
      ? navigator.languages
      : navigator.language
        ? [navigator.language]
        : [];
  return [...list];
}

/**
 * Picks the first browser-preferred language we ship translations for, scanning
 * the ordered `navigator.languages` list. Unlike `mapLocaleToUiLanguage` this
 * does not collapse an unsupported language to English prematurely — it keeps
 * scanning so `["pt-BR", "fr-FR"]` resolves to French, not English.
 */
export function detectBrowserUiLanguage(
  languages: readonly string[] = browserLanguages(),
): SupportedUiLanguage {
  for (const tag of languages) {
    const primary = (tag.split("-")[0] ?? "").toLowerCase();
    if (primary === "no" || primary === "nb" || primary === "nn") return "nb";
    if ((SUPPORTED_UI_LANGUAGES as readonly string[]).includes(primary)) {
      return primary as SupportedUiLanguage;
    }
  }
  return DEFAULT_UI_LANGUAGE;
}

/**
 * Picks the Polish plural form for `one | few | many` messages.
 *
 * Polish has three forms where the other shipped languages have two, so
 * vue-i18n's built-in rule (0 → first, 1 → second, rest → third) picks the
 * wrong one for every count below five. CLDR: `one` for exactly 1, `few` for
 * counts ending in 2–4 except the teens, `many` for everything else — zero
 * included.
 */
function polishPluralIndex(choice: number, choicesLength: number): number {
  if (choicesLength < 3) return choice === 1 ? 0 : 1;
  if (choice === 1) return 0;
  const lastDigit = Math.abs(choice) % 10;
  const lastTwoDigits = Math.abs(choice) % 100;
  const few = lastDigit >= 2 && lastDigit <= 4 && !(lastTwoDigits >= 12 && lastTwoDigits <= 14);
  return few ? 1 : 2;
}

export const i18n = createI18n({
  legacy: false,
  globalInjection: true,
  locale: detectBrowserUiLanguage(),
  fallbackLocale: "en",
  messages: { de, en, es, fr, it, nb, nl, pl },
  pluralRules: { pl: polishPluralIndex },
});

/** Sets the active UI language and keeps `<html lang>` in sync. */
export function setUiLanguage(bcp47: string | undefined): void {
  const language = mapLocaleToUiLanguage(bcp47);
  i18n.global.locale.value = language;
  if (typeof document !== "undefined") {
    document.documentElement.lang = language;
  }
}

/** Non-component translation access, e.g. from plain modules. */
export function t(key: string, params?: Record<string, unknown>): string {
  return params ? String(i18n.global.t(key, params)) : String(i18n.global.t(key));
}
