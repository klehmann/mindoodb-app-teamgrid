/**
 * PrimeVue/Aura theme glue for Teamgrid.
 *
 * Teamgrid uses the MindooDB primary palette and follows the Haven host's
 * light/dark mode. The host pushes theme changes through the bridge; this
 * module just translates `theme.mode` into the `data-theme` attribute the
 * Aura preset reads and into a PrimeVue theme object.
 */
import { Theme, definePreset } from "@primeuix/styled";
import Aura from "@primeuix/themes/aura";

import type { MindooDBAppHostTheme, MindooDBAppThemeMode } from "mindoodb-app-sdk";

/** Brand color ramp used in both light and dark mode. */
const MINDOO_PRIMARY_SCALE = {
  50: "#eef2ff",
  100: "#dce5ff",
  200: "#bfceff",
  300: "#98adff",
  400: "#6f88f6",
  500: "#4d67d9",
  600: "#354fb9",
  700: "#253c92",
  800: "#1f3a8a",
  900: "#182a63",
  950: "#0b1020",
} as const;

const MINDOO_THEME_PRESET = definePreset(Aura, {
  semantic: {
    primary: MINDOO_PRIMARY_SCALE,
    focusRing: {
      width: "2px",
      style: "solid",
      color: "rgba(212, 160, 23, 0.38)",
      offset: "2px",
      shadow: "none",
    },
  },
});

export const DEFAULT_THEME_PRESET = "mindoo";
export const DEFAULT_THEME_MODE: MindooDBAppThemeMode = "dark";

/** Build the PrimeVue config object used at `app.use(PrimeVue, ...)` time. */
export function buildPrimeVueTheme() {
  return {
    preset: MINDOO_THEME_PRESET,
    options: {
      darkModeSelector: "[data-theme='dark']",
    },
  };
}

/**
 * Apply a host-provided theme to the document and PrimeVue runtime.
 *
 * Called once on mount with the launch context theme, and again from the
 * bridge's `onThemeChange` listener whenever the host toggles light/dark mode.
 */
export function applyAppTheme(theme?: Partial<MindooDBAppHostTheme> | null) {
  const mode = theme?.mode === "light" ? "light" : DEFAULT_THEME_MODE;
  if (typeof document !== "undefined") {
    document.documentElement.dataset.theme = mode;
    document.documentElement.dataset.themePreset = DEFAULT_THEME_PRESET;
  }
  Theme.setTheme(buildPrimeVueTheme());
  return { mode, preset: DEFAULT_THEME_PRESET };
}
