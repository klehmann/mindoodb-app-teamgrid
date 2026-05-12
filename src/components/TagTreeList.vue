<script setup lang="ts">
/**
 * Recursive tree component used in the File/Open dialog left pane.
 *
 * Renders a list of {@link OpenCategoryNode}s and recurses into each node's
 * children. The component is presentational: it does not own the selection
 * state. The parent supplies the currently selected node's key via
 * `selectedKey`, and listens to `select` to change the selection. Counts
 * shown in the badge come from the view navigator's
 * `descendantDocumentCount`.
 *
 * Props:
 * - `nodes` (required): top-level category nodes to render at this level.
 * - `selectedKey` (required): key of the currently active node.
 *
 * Emits:
 * - `select(key)`: user clicked a node and wants the parent to make it active.
 */
import type { OpenCategoryNode } from "@/lib/viewOpen";

defineOptions({ name: "TagTreeList" });

defineProps<{
  nodes: OpenCategoryNode[];
  selectedKey: string;
}>();

defineEmits<{
  select: [key: string];
}>();
</script>

<template>
  <ul class="tag-tree-list">
    <li v-for="node in nodes" :key="node.key" class="tag-tree-list__item">
      <button
        type="button"
        class="tag-tree-list__button"
        :class="{ 'tag-tree-list__button--selected': node.key === selectedKey }"
        @click="$emit('select', node.key)"
      >
        <span>{{ node.label }}</span>
        <small>{{ node.count }}</small>
      </button>
      <TagTreeList
        v-if="node.children.length > 0"
        :nodes="node.children"
        :selected-key="selectedKey"
        @select="$emit('select', $event)"
      />
    </li>
  </ul>
</template>
