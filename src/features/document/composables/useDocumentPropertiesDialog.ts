/**
 * State for the Spreadsheet Properties dialog (title, tags, and locale).
 *
 * Edits are batched into a single `setDocumentProperties` operation through
 * the document composable so they participate in the granular-save /
 * `baseHeads` machinery rather than going around it.
 */
import { computed, ref, watch } from "vue";
import { DEFAULT_TEAMGRID_LOCALE, normalizeTags, normalizeTeamGridLocale, readIsTemplate } from "@/features/document/lib/teamgridDocument";
import type { TeamGridAppApi } from "@/features/document/composables/useTeamGridDocument";

export interface UseDocumentPropertiesDialogOptions {
  app: TeamGridAppApi;
}

export function useDocumentPropertiesDialog(options: UseDocumentPropertiesDialogOptions) {
  const { app } = options;
  const propertiesDialogVisible = ref(false);
  const propertiesTitleDraft = ref("");
  const propertiesTagsDraft = ref("");
  const propertiesIsTemplateDraft = ref(false);
  const propertiesLocaleDraft = ref(DEFAULT_TEAMGRID_LOCALE);
  const baseLocaleOptions = [
    { value: "en-US", label: "English (United States)" },
    { value: "en-GB", label: "English (United Kingdom)" },
    { value: "de-DE", label: "German (Germany)" },
    { value: "de-AT", label: "German (Austria)" },
    { value: "de-CH", label: "German (Switzerland)" },
    { value: "fr-FR", label: "French (France)" },
    { value: "it-IT", label: "Italian (Italy)" },
    { value: "es-ES", label: "Spanish (Spain)" },
    { value: "nl-NL", label: "Dutch (Netherlands)" },
  ];
  const propertiesLocaleOptions = computed(() => {
    const currentLocale = normalizeTeamGridLocale(propertiesLocaleDraft.value);
    if (baseLocaleOptions.some((option) => option.value === currentLocale)) {
      return baseLocaleOptions;
    }
    return [
      ...baseLocaleOptions,
      { value: currentLocale, label: currentLocale },
    ];
  });

  function resetPropertiesDraft() {
    propertiesTitleDraft.value = app.activeSubject.value || "Untitled spreadsheet";
    propertiesTagsDraft.value = app.activeTags.value.join("\n");
    propertiesIsTemplateDraft.value = readIsTemplate(app.currentEnvelope.value);
    propertiesLocaleDraft.value = normalizeTeamGridLocale(app.activeGrid.value?.settings.locale);
  }

  function openPropertiesDialog() {
    if (!app.currentEnvelope.value || app.gridReadOnly.value) {
      return;
    }
    resetPropertiesDraft();
    propertiesDialogVisible.value = true;
  }

  function applyDocumentProperties() {
    if (!app.currentEnvelope.value) {
      return;
    }
    const subject = propertiesTitleDraft.value.trim() || "Untitled spreadsheet";
    const tags = normalizeTags(propertiesTagsDraft.value.split(/\r?\n/));
    const isTemplate = propertiesIsTemplateDraft.value;
    const locale = normalizeTeamGridLocale(propertiesLocaleDraft.value);
    app.updateGrid((_grid, envelope) => {
      envelope.subject = subject;
      envelope.tags = tags;
      envelope.istemplate = isTemplate;
      envelope.teamgrid.settings.locale = locale;
      return [{ type: "setDocumentProperties", subject, tags, isTemplate, locale }];
    });
    propertiesDialogVisible.value = false;
  }

  watch(
    () => app.currentEnvelope.value,
    () => {
      if (!propertiesDialogVisible.value) {
        resetPropertiesDraft();
      }
    },
    { immediate: true },
  );

  return {
    propertiesDialogVisible,
    propertiesTitleDraft,
    propertiesTagsDraft,
    propertiesIsTemplateDraft,
    propertiesLocaleDraft,
    propertiesLocaleOptions,
    openPropertiesDialog,
    applyDocumentProperties,
    resetPropertiesDraft,
  };
}
