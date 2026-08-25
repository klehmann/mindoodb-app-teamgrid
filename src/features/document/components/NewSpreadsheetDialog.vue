<script setup lang="ts">
import { computed, ref, watch } from "vue";
import Button from "primevue/button";
import Dialog from "primevue/dialog";
import { useI18n } from "vue-i18n";
import {
  type MindooDBAppCreateKeyInfo,
  type MindooDBAppDatabaseInfo,
  type MindooDBAppSession,
} from "mindoodb-app-sdk";

import TagTreeList from "@/features/document/components/TagTreeList.vue";
import DocumentRecipientsField from "@/features/document/components/DocumentRecipientsField.vue";
import { normalizeTags } from "@/features/document/lib/teamgridDocument";
import {
  ALL_SPREADSHEETS_NODE_KEY,
  buildOpenCategoryTree,
  collectNavigatorEntries,
  createOpenViewDefinition,
  dedupeDocumentEntries,
  usableExistingTagNodes,
  type OpenCategoryNode,
} from "@/features/document/lib/viewOpen";

export type NewSpreadsheetDraft = {
  databaseId: string;
  title: string;
  tags: string[];
  encryption:
    | { mode: "shared"; decryptionKeyId?: string }
    | { mode: "people"; recipients: string[] };
};

const props = defineProps<{
  visible: boolean;
  databases: MindooDBAppDatabaseInfo[];
  session: MindooDBAppSession | null;
  currentUserName: string;
  currentUserCanonical: string;
  initialDatabaseId: string;
  creating: boolean;
}>();

const emit = defineEmits<{
  "update:visible": [value: boolean];
  create: [draft: NewSpreadsheetDraft];
}>();

const { t } = useI18n();
const selectedDatabaseId = ref("");
const titleDraft = ref("");
const tagsDraft = ref("");
const encryptionMode = ref<"shared" | "people">("shared");
const selectedKeyId = ref("");
const createKeys = ref<MindooDBAppCreateKeyInfo[]>([]);
const directoryUsers = ref<string[]>([]);
const recipients = ref<string[]>([]);
const canEncryptForPeople = ref(false);
const tagNodes = ref<OpenCategoryNode[]>([]);
const selectedTagKey = ref("");
const loadingOptions = ref(false);
const loadError = ref("");

const singleDatabase = computed(() => props.databases.length <= 1);
const canSubmit = computed(
  () => Boolean(selectedDatabaseId.value) && !props.creating && !loadingOptions.value,
);

function resetForm() {
  selectedDatabaseId.value =
    props.databases.some((database) => database.id === props.initialDatabaseId)
      ? props.initialDatabaseId
      : (props.databases[0]?.id ?? "");
  titleDraft.value = "";
  tagsDraft.value = "";
  encryptionMode.value = "shared";
  selectedKeyId.value = "";
  createKeys.value = [];
  directoryUsers.value = [];
  recipients.value = [];
  canEncryptForPeople.value = false;
  tagNodes.value = [];
  selectedTagKey.value = "";
  loadError.value = "";
}

function findTagNode(nodes: OpenCategoryNode[], key: string): OpenCategoryNode | null {
  for (const node of nodes) {
    if (node.key === key) {
      return node;
    }
    const nested = findTagNode(node.children, key);
    if (nested) {
      return nested;
    }
  }
  return null;
}

function addExistingTag(key: string) {
  selectedTagKey.value = key;
  const node = findTagNode(tagNodes.value, key);
  const tag = node?.tag?.trim();
  if (!tag) {
    return;
  }
  const current = normalizeTags(tagsDraft.value.split(/\r?\n/));
  if (current.includes(tag)) {
    return;
  }
  tagsDraft.value = [...current, tag].join("\n");
}

async function loadDatabaseOptions(databaseId: string) {
  const info = props.databases.find((database) => database.id === databaseId);
  if (!props.session || !info) {
    createKeys.value = [];
    directoryUsers.value = [];
    canEncryptForPeople.value = false;
    tagNodes.value = [];
    return;
  }

  loadingOptions.value = true;
  loadError.value = "";
  try {
    const database = await props.session.openDatabase(databaseId);

    try {
      if (typeof database.documents.listCreateKeys === "function") {
        createKeys.value = await database.documents.listCreateKeys();
      } else {
        const keyId = await database.documents.getDefaultCreateKeyId();
        createKeys.value = [{ keyId, isDefault: true }];
      }
    } catch {
      try {
        const keyId = await database.documents.getDefaultCreateKeyId();
        createKeys.value = [{ keyId, isDefault: true }];
      } catch {
        createKeys.value = [];
      }
    }
    selectedKeyId.value =
      createKeys.value.find((key) => key.isDefault)?.keyId ?? createKeys.value[0]?.keyId ?? "";

    canEncryptForPeople.value =
      info.capabilities.includes("directory") && typeof database.directory?.listUsers === "function";
    if (canEncryptForPeople.value) {
      try {
        directoryUsers.value = await database.directory.listUsers();
      } catch {
        canEncryptForPeople.value = false;
        directoryUsers.value = [];
      }
    } else {
      directoryUsers.value = [];
    }
    if (!canEncryptForPeople.value) {
      encryptionMode.value = "shared";
      recipients.value = [];
    }

    tagNodes.value = [];
    if (info.capabilities.includes("views")) {
      try {
        const navigator = await props.session.createViewNavigator({
          databaseIds: [databaseId],
          definition: createOpenViewDefinition("all"),
          categorizationStyle: "category_then_document",
          options: {
            includeCategories: true,
            includeDocuments: true,
            hideEmptyCategories: true,
          },
        });
        await navigator.expandAll();
        const entries = await collectNavigatorEntries(navigator);
        await navigator.dispose();
        const documents = dedupeDocumentEntries(entries);
        const tree = buildOpenCategoryTree(
          entries.filter((entry) => entry.kind === "category"),
          documents.length,
        );
        tagNodes.value = usableExistingTagNodes(
          (tree.roots[0]?.children ?? []).filter(
            (node) => node.key !== ALL_SPREADSHEETS_NODE_KEY,
          ),
        );
      } catch {
        tagNodes.value = [];
      }
    }
  } catch (error) {
    loadError.value = error instanceof Error ? error.message : String(error);
    createKeys.value = [];
    directoryUsers.value = [];
    canEncryptForPeople.value = false;
    tagNodes.value = [];
  } finally {
    loadingOptions.value = false;
  }
}

function submit() {
  if (!canSubmit.value) {
    return;
  }
  const encryption =
    encryptionMode.value === "people"
      ? { mode: "people" as const, recipients: [...recipients.value] }
      : {
          mode: "shared" as const,
          decryptionKeyId: selectedKeyId.value || undefined,
        };
  emit("create", {
    databaseId: selectedDatabaseId.value,
    title: titleDraft.value.trim(),
    tags: normalizeTags(tagsDraft.value.split(/\r?\n/)),
    encryption,
  });
}

function cancel() {
  emit("update:visible", false);
}

watch(
  () => props.visible,
  (visible) => {
    if (!visible) {
      return;
    }
    resetForm();
  },
  { immediate: true },
);

watch(
  () => [props.visible, selectedDatabaseId.value] as const,
  ([visible, databaseId]) => {
    if (!visible || !databaseId) {
      return;
    }
    void loadDatabaseOptions(databaseId);
  },
  { immediate: true },
);
</script>

<template>
  <Dialog
    :visible="visible"
    modal
    :header="t('newDocument.title')"
    :style="{ width: '38rem', maxWidth: '96vw' }"
    @update:visible="emit('update:visible', $event)"
  >
    <div class="dialog-content">
      <label class="field">
        <span class="field-label">{{ t("newDocument.database") }}</span>
        <select
          v-model="selectedDatabaseId"
          class="native-input"
          :disabled="singleDatabase || creating"
        >
          <option
            v-for="database in databases"
            :key="database.id"
            :value="database.id"
          >
            {{ database.title || database.id }}
          </option>
        </select>
      </label>

      <label class="field">
        <span class="field-label">{{ t("newDocument.titleLabel") }}</span>
        <input
          v-model="titleDraft"
          class="native-input"
          type="text"
          autocomplete="off"
          :placeholder="t('newDocument.titlePlaceholder')"
          :disabled="creating"
        />
      </label>

      <fieldset class="field field--choices">
        <legend class="field-label">{{ t("newDocument.encryption") }}</legend>
        <div class="choice-row">
          <label class="choice">
            <input v-model="encryptionMode" type="radio" value="shared" :disabled="creating" />
            <span>{{ t("newDocument.sharedKey") }}</span>
          </label>
          <label v-if="canEncryptForPeople" class="choice">
            <input v-model="encryptionMode" type="radio" value="people" :disabled="creating" />
            <span>{{ t("newDocument.specificPeople") }}</span>
          </label>
        </div>
      </fieldset>

      <label v-if="encryptionMode === 'shared'" class="field">
        <span class="field-label">{{ t("newDocument.sharedKey") }}</span>
        <select
          v-model="selectedKeyId"
          class="native-input"
          :disabled="creating || createKeys.length === 0"
        >
          <option v-if="createKeys.length === 0" value="">
            {{ t("newDocument.databaseDefault") }}
          </option>
          <option
            v-for="key in createKeys"
            :key="key.keyId"
            :value="key.keyId"
          >
            {{ key.isDefault ? t("newDocument.keyDefault", { key: key.keyId }) : key.keyId }}
          </option>
        </select>
      </label>

      <DocumentRecipientsField
        v-else
        v-model="recipients"
        :current-user-name="currentUserName"
        :current-user-canonical="currentUserCanonical"
        :directory-users="directoryUsers"
        :disabled="creating"
      />

      <label class="field">
        <span class="field-label">{{ t("newDocument.tagsLabel") }}</span>
        <textarea
          v-model="tagsDraft"
          class="native-input native-input--textarea"
          rows="4"
          :placeholder="t('newDocument.tagsPlaceholder')"
          :disabled="creating"
        />
      </label>
      <p class="field-hint">
        {{ t("newDocument.tagsHint") }}
      </p>
      <div v-if="tagNodes.length > 0" class="existing-tags">
        <span class="field-label">{{ t("newDocument.existingTags") }}</span>
        <div class="existing-tags__tree">
          <TagTreeList
            :nodes="tagNodes"
            :selected-key="selectedTagKey"
            @select="addExistingTag"
          />
        </div>
      </div>
      <p v-if="loadError" class="field-hint">{{ loadError }}</p>
    </div>
    <template #footer>
      <Button :label="t('common.cancel')" text :disabled="creating" @click="cancel" />
      <Button
        :label="t('common.create')"
        icon="pi pi-check"
        :disabled="!canSubmit"
        :loading="creating || loadingOptions"
        @click="submit"
      />
    </template>
  </Dialog>
</template>

<style scoped>
.dialog-content {
  display: grid;
  gap: 1rem;
}

.field {
  display: grid;
  gap: 0.35rem;
}

.field--choices {
  border: 0;
  margin: 0;
  padding: 0;
}

.choice-row {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 0.85rem 1.35rem;
}

.field-label {
  margin: 0;
  color: var(--muted);
  font-size: 0.78rem;
  text-transform: uppercase;
  letter-spacing: 0.08em;
}

.choice {
  display: inline-flex;
  align-items: center;
  gap: 0.45rem;
}

.native-input {
  width: 100%;
  padding: 0.7rem 0.85rem;
  border: 1px solid var(--border);
  border-radius: 0.85rem;
  background: rgb(255 255 255 / 0.04);
  color: inherit;
}

.native-input--textarea {
  min-height: 6rem;
  resize: vertical;
}

.field-hint {
  margin: 0;
  color: var(--muted);
  font-size: 0.85rem;
}

.field-hint--tight {
  margin-top: -0.35rem;
}

.field-hint code {
  color: var(--accent);
}

.existing-tags {
  display: grid;
  gap: 0.55rem;
}

.existing-tags__tree {
  max-height: 10rem;
  overflow: auto;
  padding: 0.35rem;
  border: 1px solid var(--border);
  border-radius: 0.85rem;
  background: rgb(255 255 255 / 0.025);
}
</style>
