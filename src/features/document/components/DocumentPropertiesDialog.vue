<script setup lang="ts">
/**
 * Title, tags, and locale editor for the current spreadsheet document.
 *
 * Edits go through `useDocumentPropertiesDialog`, which batches them into a
 * single `setDocumentProperties` operation so they participate in the
 * granular-save machinery rather than going around it.
 */
import Button from "primevue/button";
import Dialog from "primevue/dialog";
import { useI18n } from "vue-i18n";
import DocumentRecipientsField from "@/features/document/components/DocumentRecipientsField.vue";
import type { useDocumentPropertiesDialog } from "@/features/document/composables/useDocumentPropertiesDialog";


const props = defineProps<{
  controller: ReturnType<typeof useDocumentPropertiesDialog>;
  readOnly: boolean;
}>();

const {
  propertiesDialogVisible,
  propertiesTitleDraft,
  propertiesTagsDraft,
  propertiesIsTemplateDraft,
  propertiesLocaleDraft,
  propertiesLocaleOptions,
  applyDocumentProperties,
  resetPropertiesDraft,
  propertiesIsSealed,
  propertiesRecipientsDraft,
  propertiesDirectoryUsers,
  propertiesRecipientsError,
  propertiesApplying,
  propertiesCurrentUserName,
  propertiesCurrentUserCanonical,
} = props.controller;

const { t } = useI18n();

function cancel() {
  propertiesDialogVisible.value = false;
  resetPropertiesDraft();
}
</script>

<template>
  <Dialog
    v-model:visible="propertiesDialogVisible"
    modal
    :header="t('app.properties.title')"
    :style="{ width: '34rem', maxWidth: '96vw' }"
  >
    <div class="properties-dialog">
      <label class="field">
        {{ t('app.properties.titleLabel') }}
        <input v-model="propertiesTitleDraft" class="native-input" type="text" autocomplete="off">
      </label>
      <label class="field">
        {{ t('app.properties.tagsLabel') }}
        <textarea v-model="propertiesTagsDraft" class="native-input native-input--textarea" rows="6" placeholder="Work\Planning&#10;Finance" />
      </label>
      <p class="properties-dialog__hint">{{ t("app.properties.tagsHint") }}</p>
      <label class="field">
        {{ t('app.properties.localeLabel') }}
        <select v-model="propertiesLocaleDraft" class="native-input">
          <option v-for="option in propertiesLocaleOptions" :key="option.value" :value="option.value">{{ option.label }}</option>
        </select>
      </label>
      <label class="properties-dialog__checkbox">
        <input v-model="propertiesIsTemplateDraft" type="checkbox">
        {{ t('app.properties.useAsTemplate') }}
      </label>
      <DocumentRecipientsField
        v-if="propertiesIsSealed"
        v-model="propertiesRecipientsDraft"
        :current-user-name="propertiesCurrentUserName"
        :current-user-canonical="propertiesCurrentUserCanonical"
        :directory-users="propertiesDirectoryUsers"
        :disabled="readOnly || propertiesApplying"
      />
      <p v-if="propertiesRecipientsError" class="properties-dialog__hint">{{ propertiesRecipientsError }}</p>
    </div>
    <template #footer>
      <Button :label="t('common.cancel')" text @click="cancel" />
      <Button :label="t('common.apply')" icon="pi pi-check" :disabled="readOnly || propertiesApplying" @click="applyDocumentProperties" />
    </template>
  </Dialog>
</template>

<style scoped>
.properties-dialog {
  display: grid;
  gap: 1rem;
}

.properties-dialog__hint {
  margin: 0;
  color: var(--muted);
  line-height: 1.5;
}

.properties-dialog__checkbox {
  display: inline-flex;
  align-items: center;
  gap: 0.5rem;
  color: var(--text);
}

.properties-dialog code {
  color: var(--formula-text);
}
</style>
