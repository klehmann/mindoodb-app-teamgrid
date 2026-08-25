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
import {
  extraEncryptForUsernames,
  isSealedEncryptForDocument,
  recipientDiff,
} from "@/features/document/lib/sealedRecipients";
import { t } from "@/i18n";


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
  const propertiesIsSealed = ref(false);
  const propertiesRecipientsDraft = ref<string[]>([]);
  const propertiesRecipientsBaseline = ref<string[]>([]);
  const propertiesDirectoryUsers = ref<string[]>([]);
  const propertiesRecipientsError = ref("");
  const propertiesApplying = ref(false);

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
    propertiesTitleDraft.value = app.activeSubject.value || t("common.untitled");
    propertiesTagsDraft.value = app.activeTags.value.join("\n");
    propertiesIsTemplateDraft.value = readIsTemplate(app.currentEnvelope.value);
    propertiesLocaleDraft.value = normalizeTeamGridLocale(app.activeGrid.value?.settings.locale);
    const data = app.currentDocument.value?.data as Record<string, unknown> | undefined;
    propertiesIsSealed.value = isSealedEncryptForDocument(data);
    propertiesRecipientsDraft.value = extraEncryptForUsernames(data, app.currentUserCanonical.value);
    propertiesRecipientsBaseline.value = [...propertiesRecipientsDraft.value];
    propertiesRecipientsError.value = "";
  }

  async function openPropertiesDialog() {
    if (!app.currentEnvelope.value || app.gridReadOnly.value) {
      return;
    }
    resetPropertiesDraft();
    propertiesDirectoryUsers.value = [];
    propertiesDialogVisible.value = true;
    const database = app.currentDatabase.value;
    const info = app.currentDatabaseInfo.value;
    if (
      propertiesIsSealed.value
      && database
      && info?.capabilities.includes("directory")
      && typeof database.directory?.listUsers === "function"
    ) {
      try {
        propertiesDirectoryUsers.value = await database.directory.listUsers();
      } catch {
        propertiesDirectoryUsers.value = [];
      }
    }
  }

  async function applyDocumentProperties() {
    if (!app.currentEnvelope.value) {
      return;
    }
    propertiesApplying.value = true;
    propertiesRecipientsError.value = "";
    try {
      const database = app.currentDatabase.value;
      const document = app.currentDocument.value;
      if (propertiesIsSealed.value && database && document) {
        if (typeof database.documents.addRecipients !== "function" || typeof database.documents.removeRecipients !== "function") {
          propertiesRecipientsError.value = t("app.properties.recipientsUnavailable");
          return;
        }
        const { added, removed } = recipientDiff(propertiesRecipientsBaseline.value, propertiesRecipientsDraft.value);
        if (added.length) {
          await database.documents.addRecipients(document.id, added);
        }
        if (removed.length) {
          await database.documents.removeRecipients(document.id, removed);
        }
      }
      const subject = propertiesTitleDraft.value.trim() || t("common.untitled");
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
    } catch (error) {
      propertiesRecipientsError.value = error instanceof Error ? error.message : String(error);
    } finally {
      propertiesApplying.value = false;
    }
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
    propertiesIsSealed,
    propertiesRecipientsDraft,
    propertiesDirectoryUsers,
    propertiesRecipientsError,
    propertiesApplying,
    propertiesCurrentUserName: computed(() => app.currentUserName.value),
    propertiesCurrentUserCanonical: computed(() => app.currentUserCanonical.value),
    openPropertiesDialog,
    applyDocumentProperties,
    resetPropertiesDraft,
  };
}
