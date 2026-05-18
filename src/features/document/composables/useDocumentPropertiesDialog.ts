/**
 * State for the Spreadsheet Properties dialog (title + tags).
 *
 * Title and tag edits are batched into a single `setDocumentProperties`
 * operation through the document composable so they participate in the
 * granular-save / `baseHeads` machinery rather than going around it.
 */
import { ref, watch } from "vue";
import { normalizeTags } from "@/features/document/lib/teamgridDocument";
import type { TeamGridAppApi } from "@/features/document/composables/useTeamGridDocument";

export interface UseDocumentPropertiesDialogOptions {
  app: TeamGridAppApi;
}

export function useDocumentPropertiesDialog(options: UseDocumentPropertiesDialogOptions) {
  const { app } = options;
  const propertiesDialogVisible = ref(false);
  const propertiesTitleDraft = ref("");
  const propertiesTagsDraft = ref("");

  function resetPropertiesDraft() {
    propertiesTitleDraft.value = app.activeSubject.value || "Untitled spreadsheet";
    propertiesTagsDraft.value = app.activeTags.value.join("\n");
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
    app.updateGrid((_grid, envelope) => {
      envelope.subject = subject;
      envelope.tags = tags;
      return [{ type: "setDocumentProperties", subject, tags }];
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
    openPropertiesDialog,
    applyDocumentProperties,
    resetPropertiesDraft,
  };
}
