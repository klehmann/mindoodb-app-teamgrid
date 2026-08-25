<script setup lang="ts">
import { computed, ref } from "vue";
import Button from "primevue/button";
import { useI18n } from "vue-i18n";
import { abbreviateCanonicalName } from "mindoodb-app-sdk";
import { preferDirectoryUsername, recipientNamesEqual } from "@/features/document/lib/sealedRecipients";

const props = defineProps<{
  currentUserName: string;
  currentUserCanonical: string;
  directoryUsers: string[];
  disabled?: boolean;
}>();

const recipients = defineModel<string[]>({ default: () => [] });
const { t } = useI18n();
const recipientDraft = ref("");

const availableUsers = computed(() => {
  const excluded = [props.currentUserCanonical, ...recipients.value];
  return props.directoryUsers.filter(
    (name) => !excluded.some((existing) => recipientNamesEqual(existing, name)),
  );
});

function displayName(name: string) {
  const resolved = preferDirectoryUsername(name, props.directoryUsers);
  return abbreviateCanonicalName(resolved) || resolved;
}

function addRecipient(name: string) {
  const trimmed = name.trim();
  if (!trimmed || props.disabled) {
    return;
  }
  if (recipientNamesEqual(trimmed, props.currentUserCanonical)) {
    return;
  }
  if (recipients.value.some((existing) => recipientNamesEqual(existing, trimmed))) {
    return;
  }
  recipients.value = [...recipients.value, trimmed];
  recipientDraft.value = "";
}

function removeRecipient(name: string) {
  if (props.disabled) {
    return;
  }
  recipients.value = recipients.value.filter((existing) => existing !== name);
}
</script>

<template>
  <div class="recipients">
    <p class="field-hint field-hint--tight">
      {{ t("recipients.hint") }}
    </p>
    <div class="recipients__chips">
      <span class="recipients__chip recipients__chip--you">
        {{ currentUserName }}
        <em>{{ t("common.you") }}</em>
      </span>
      <button
        v-for="name in recipients"
        :key="name"
        type="button"
        class="recipients__chip"
        :disabled="disabled"
        :aria-label="t('recipients.removeAria', { name: displayName(name) })"
        @click="removeRecipient(name)"
      >
        <span>{{ displayName(name) }}</span>
        <span aria-hidden="true">×</span>
      </button>
    </div>
    <form class="recipients__add" @submit.prevent="addRecipient(recipientDraft)">
      <select v-model="recipientDraft" class="native-input" :disabled="disabled">
        <option value="">{{ t("recipients.selectUser") }}</option>
        <option
          v-for="name in availableUsers"
          :key="name"
          :value="name"
        >
          {{ displayName(name) }}
        </option>
      </select>
      <Button
        :label="t('common.add')"
        type="submit"
        severity="secondary"
        :disabled="disabled || !recipientDraft"
      />
    </form>
  </div>
</template>

<style scoped>
.recipients {
  display: grid;
  gap: 0.55rem;
}

.field-hint {
  margin: 0;
  color: var(--muted);
  font-size: 0.85rem;
}

.field-hint--tight {
  margin-top: -0.35rem;
}

.recipients__chips {
  display: flex;
  flex-wrap: wrap;
  gap: 0.4rem;
}

.recipients__chip {
  display: inline-flex;
  align-items: center;
  gap: 0.35rem;
  padding: 0.28rem 0.65rem;
  border: 0;
  border-radius: 999px;
  background: rgb(90 160 220 / 0.22);
  color: inherit;
  cursor: pointer;
}

.recipients__chip--you {
  cursor: default;
}

.recipients__chip em {
  font-style: normal;
  color: var(--muted);
}

.recipients__add {
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto;
  gap: 0.55rem;
}

.native-input {
  width: 100%;
  padding: 0.7rem 0.85rem;
  border: 1px solid var(--border);
  border-radius: 0.85rem;
  background: rgb(255 255 255 / 0.04);
  color: inherit;
}
</style>
