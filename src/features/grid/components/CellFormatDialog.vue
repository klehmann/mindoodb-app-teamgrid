<script setup lang="ts">
/**
 * Excel-like "Format cells" dialog: cell type, font, fill, and border tabs.
 *
 * All draft state and apply behaviour live in `useCellFormatDialog`. This
 * component renders the four tabs and the apply / cancel actions, and
 * forwards the read-only flag so the apply button can be disabled in
 * historical revisions and other guarded modes.
 */
import { computed, watch } from "vue";
import Button from "primevue/button";
import Dialog from "primevue/dialog";
import Select from "primevue/select";
import type { useCellFormatDialog } from "@/features/grid/composables/useCellFormatDialog";
import { DATE_FORMAT_OPTIONS, formatDatePreview } from "@/features/grid/lib/cellFormatting";

/**
 * Curated font suggestions for the editable Font family combobox.
 *
 * Each entry is also rendered using its own font name so users get a live
 * preview of how the cell will look. The combobox is editable, so any
 * string a user types — including XLSX-imported names like `Wingdings`,
 * `Aptos`, or `Segoe UI` that we don't bundle — is preserved verbatim.
 *
 * The Microsoft Office / Windows families (`Calibri`, `Cambria`,
 * `Courier New`, `Arial`, `Times New Roman`) are aliased to the bundled
 * metric-compatible Croscore fonts (`Carlito`, `Caladea`, `Cousine`,
 * `Arimo`, `Tinos`) via `@font-face` rules in `main.css`, so cells
 * imported from Office documents render with the right glyph widths
 * even on machines without Office installed.
 */
const FONT_FAMILY_OPTIONS: string[] = [
  "Inter",
  "Rubik",
  "JetBrains Mono",
  "Calibri",
  "Cambria",
  "Arial",
  "Helvetica",
  "Times New Roman",
  "Georgia",
  "Courier New",
  "Verdana",
  "Tahoma",
  "Trebuchet MS",
  "Consolas",
  "Segoe UI",
];

/**
 * Classic Excel font-size presets shown in the editable size combobox.
 *
 * Excel ships exactly this list (8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24,
 * 26, 28, 36, 48, 72). The combobox is editable so users can still type any
 * integer like 13 or 200; we clamp to a sane range when committing the
 * value into the cell style.
 */
const FONT_SIZE_OPTIONS: number[] = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
const MIN_FONT_SIZE = 4;
const MAX_FONT_SIZE = 200;

const props = defineProps<{
  controller: ReturnType<typeof useCellFormatDialog>;
  readOnly: boolean;
}>();

const {
  formatDialogVisible,
  formatDialogTab,
  formatDialogKind,
  formatDialogCurrency,
  formatDialogCustomNumFmt,
  formatDialogLocale,
  formatDialogFontFamily,
  formatDialogFontSize,
  formatDialogBold,
  formatDialogItalic,
  formatDialogUnderline,
  formatDialogTextColor,
  formatDialogFillEnabled,
  formatDialogFillColor,
  formatDialogBorderStyle,
  formatDialogBorderColor,
  formatDialogBorders,
  formatDialogHorizontalAlign,
  formatDialogVerticalAlign,
  formatDialogIndent,
  formatDialogWrapText,
  applySelectedCellFormat,
  updateCustomBordersFromLineSelection,
  toggleFormatDialogBorder,
  setFormatDialogBorderPreset,
  currentDialogBorderCss,
} = props.controller;

/** Indent applies only when horizontal alignment is Left or Right. */
const indentEnabled = computed(() => formatDialogHorizontalAlign.value === "left" || formatDialogHorizontalAlign.value === "right");
const dateFormatOptions = computed(() => {
  const options = DATE_FORMAT_OPTIONS.map((option) => ({
    label: formatDatePreview(option.excelNumFmt, formatDialogLocale.value),
    value: option.excelNumFmt,
    kind: option.kind,
  }));
  if (
    isDateFormatKind.value
    && formatDialogCustomNumFmt.value
    && !options.some((option) => option.kind === formatDialogKind.value && option.value === formatDialogCustomNumFmt.value)
  ) {
    options.push({
      label: formatDatePreview(formatDialogCustomNumFmt.value, formatDialogLocale.value),
      value: formatDialogCustomNumFmt.value,
      kind: formatDialogKind.value as "date" | "dateTime" | "time",
    });
  }
  return options;
});
const selectedDatePreview = computed(() => formatDialogCustomNumFmt.value
  ? formatDatePreview(formatDialogCustomNumFmt.value, formatDialogLocale.value)
  : "");
const isDateFormatKind = computed(() => formatDialogKind.value === "date" || formatDialogKind.value === "dateTime" || formatDialogKind.value === "time");
watch(formatDialogKind, (kind) => {
  if (kind !== "date" && kind !== "dateTime" && kind !== "time") {
    return;
  }
  if (DATE_FORMAT_OPTIONS.some((option) => option.kind === kind && option.excelNumFmt === formatDialogCustomNumFmt.value)) {
    return;
  }
  formatDialogCustomNumFmt.value = DATE_FORMAT_OPTIONS.find((option) => option.kind === kind)?.excelNumFmt ?? "";
});

/**
 * Two-way bridge between the editable size Select (which can hold a string
 * when the user types a custom value) and the numeric `formatDialogFontSize`
 * ref. We accept any non-empty string that parses to a positive integer,
 * clamp it to a sane range, and silently ignore invalid input so the cell
 * style never receives `NaN`.
 */
const fontSizeModel = computed<number | string>({
  get: () => formatDialogFontSize.value,
  set: (next) => {
    const parsed = typeof next === "number" ? next : Number.parseInt(String(next).trim(), 10);
    if (!Number.isFinite(parsed) || parsed <= 0) {
      return;
    }
    formatDialogFontSize.value = Math.min(MAX_FONT_SIZE, Math.max(MIN_FONT_SIZE, Math.round(parsed)));
  },
});
</script>

<template>
  <Dialog v-model:visible="formatDialogVisible" modal header="Format cells" :style="{ width: '48rem', maxWidth: '96vw' }">
    <div class="cell-format-dialog">
      <div class="cell-format-dialog__tabs" role="tablist" aria-label="Format cells sections">
        <button type="button" :class="{ 'cell-format-dialog__tab--active': formatDialogTab === 'cellType' }" @click="formatDialogTab = 'cellType'">Cell type</button>
        <button type="button" :class="{ 'cell-format-dialog__tab--active': formatDialogTab === 'alignment' }" @click="formatDialogTab = 'alignment'">Alignment</button>
        <button type="button" :class="{ 'cell-format-dialog__tab--active': formatDialogTab === 'font' }" @click="formatDialogTab = 'font'">Font</button>
        <button type="button" :class="{ 'cell-format-dialog__tab--active': formatDialogTab === 'fill' }" @click="formatDialogTab = 'fill'">Fill</button>
        <button type="button" :class="{ 'cell-format-dialog__tab--active': formatDialogTab === 'border' }" @click="formatDialogTab = 'border'">Border</button>
      </div>

      <section v-if="formatDialogTab === 'cellType'" class="cell-format-dialog__panel">
        <label class="field">
          Format
          <select v-model="formatDialogKind" class="native-input">
            <option value="text">Text</option>
            <option value="general">Standard</option>
            <option value="integer">Integer</option>
            <option value="decimal">Decimal</option>
            <option value="percent">Percent</option>
            <option value="currency">Currency</option>
            <option value="date">Date</option>
            <option value="dateTime">Date &amp; Time</option>
            <option value="time">Time</option>
            <option value="custom">Custom Excel format</option>
          </select>
        </label>
        <label v-if="formatDialogKind === 'currency'" class="field">
          Currency
          <select v-model="formatDialogCurrency" class="native-input">
            <option value="USD">USD</option>
            <option value="EUR">EUR</option>
          </select>
        </label>
        <label v-if="formatDialogKind === 'custom'" class="field">
          Excel number format
          <input
            v-model="formatDialogCustomNumFmt"
            class="native-input"
            type="text"
            autocomplete="off"
            placeholder="$0.00;[Red]-$0.00"
            @keyup.enter="applySelectedCellFormat"
          >
        </label>
        <label v-if="isDateFormatKind" class="field">
          Type
          <select v-model="formatDialogCustomNumFmt" class="native-input">
            <option
              v-for="option in dateFormatOptions.filter((candidate) => candidate.kind === formatDialogKind)"
              :key="option.value"
              :value="option.value"
            >
              {{ option.label }}
            </option>
          </select>
        </label>
        <div v-if="isDateFormatKind" class="field">
          <span>Example</span>
          <div class="cell-format-dialog__sample">{{ selectedDatePreview }}</div>
        </div>
        <p class="cell-format-dialog__hint">Numeric-looking text is converted for number, percent, and currency formats; date-looking text is converted for date formats. Incompatible values are left unchanged.</p>
      </section>

      <section v-else-if="formatDialogTab === 'alignment'" class="cell-format-dialog__panel cell-format-dialog__grid">
        <label class="field">
          Horizontal alignment
          <select v-model="formatDialogHorizontalAlign" class="native-input">
            <option value="general">General</option>
            <option value="left">Left</option>
            <option value="center">Center</option>
            <option value="right">Right</option>
          </select>
        </label>
        <label class="field">
          Vertical alignment
          <select v-model="formatDialogVerticalAlign" class="native-input">
            <option value="top">Top</option>
            <option value="middle">Middle</option>
            <option value="bottom">Bottom</option>
          </select>
        </label>
        <label class="field">
          Indent
          <input
            v-model.number="formatDialogIndent"
            class="native-input"
            type="number"
            min="0"
            max="15"
            step="1"
            :disabled="!indentEnabled"
          >
        </label>
        <div class="cell-format-dialog__checks">
          <label><input v-model="formatDialogWrapText" type="checkbox"> Wrap text</label>
        </div>
        <p class="cell-format-dialog__hint">"General" right-aligns numbers and dates and left-aligns text, the same way Excel's default does. Indent only applies to Left or Right alignment. Wrap text shows multi-line content (Alt+Enter) and long values on multiple lines within the cell instead of letting them spill into empty neighbours.</p>
      </section>

      <section v-else-if="formatDialogTab === 'font'" class="cell-format-dialog__panel cell-format-dialog__grid">
        <label class="field">
          Font family
          <Select
            v-model="formatDialogFontFamily"
            :options="FONT_FAMILY_OPTIONS"
            editable
            class="cell-format-dialog__font-family"
            placeholder="Pick or type a font"
          >
            <template #value="{ value, placeholder }">
              <span v-if="value" :style="{ fontFamily: `'${value}', ${value}, var(--font-body)` }">{{ value }}</span>
              <span v-else class="cell-format-dialog__font-family-placeholder">{{ placeholder }}</span>
            </template>
            <template #option="{ option }">
              <span :style="{ fontFamily: `'${option}', ${option}, var(--font-body)` }">{{ option }}</span>
            </template>
          </Select>
        </label>
        <label class="field">
          Font size
          <Select
            v-model="fontSizeModel"
            :options="FONT_SIZE_OPTIONS"
            editable
            class="cell-format-dialog__font-size"
            placeholder="Pick or type a size"
          />
        </label>
        <label class="field">
          Text color
          <input v-model="formatDialogTextColor" class="native-input native-input--color" type="color">
        </label>
        <div class="cell-format-dialog__checks">
          <label><input v-model="formatDialogBold" type="checkbox"> Bold</label>
          <label><input v-model="formatDialogItalic" type="checkbox"> Italic</label>
          <label><input v-model="formatDialogUnderline" type="checkbox"> Underline</label>
        </div>
      </section>

      <section v-else-if="formatDialogTab === 'fill'" class="cell-format-dialog__panel cell-format-dialog__grid">
        <div class="cell-format-dialog__checks">
          <label><input v-model="formatDialogFillEnabled" type="checkbox"> Use fill color</label>
        </div>
        <label v-if="formatDialogFillEnabled" class="field">
          Background color
          <input v-model="formatDialogFillColor" class="native-input native-input--color" type="color">
        </label>
      </section>

      <section v-else class="cell-format-dialog__panel cell-format-dialog__border">
        <div class="cell-format-dialog__border-tools">
          <label class="field">
            Line style
            <select v-model="formatDialogBorderStyle" class="native-input" @change="updateCustomBordersFromLineSelection">
              <option value="none">None</option>
              <option value="thin">Thin</option>
              <option value="medium">Medium</option>
              <option value="thick">Thick</option>
              <option value="dashed">Dashed</option>
              <option value="dotted">Dotted</option>
              <option value="double">Double</option>
            </select>
          </label>
          <label class="field">
            Line color
            <input v-model="formatDialogBorderColor" class="native-input native-input--color" type="color" :disabled="formatDialogBorderStyle === 'none'" @input="updateCustomBordersFromLineSelection">
          </label>
        </div>
        <div class="cell-format-dialog__border-main">
          <div class="cell-format-dialog__presets">
            <span>Presets</span>
            <button type="button" @click="setFormatDialogBorderPreset('none')">None</button>
            <button type="button" @click="setFormatDialogBorderPreset('outline')">Outline</button>
            <button type="button" @click="setFormatDialogBorderPreset('inside')">Inside</button>
            <button type="button" @click="setFormatDialogBorderPreset('all')">All</button>
          </div>
          <div class="cell-format-dialog__border-preview" aria-label="Border preview">
            <button type="button" class="cell-format-dialog__border-toggle cell-format-dialog__border-toggle--top" @click="toggleFormatDialogBorder('top')">Top</button>
            <button type="button" class="cell-format-dialog__border-toggle cell-format-dialog__border-toggle--right" @click="toggleFormatDialogBorder('right')">Right</button>
            <button type="button" class="cell-format-dialog__border-toggle cell-format-dialog__border-toggle--bottom" @click="toggleFormatDialogBorder('bottom')">Bottom</button>
            <button type="button" class="cell-format-dialog__border-toggle cell-format-dialog__border-toggle--left" @click="toggleFormatDialogBorder('left')">Left</button>
            <div class="cell-format-dialog__border-sample" :style="{ borderTop: currentDialogBorderCss(formatDialogBorders.top), borderRight: currentDialogBorderCss(formatDialogBorders.right), borderBottom: currentDialogBorderCss(formatDialogBorders.bottom), borderLeft: currentDialogBorderCss(formatDialogBorders.left) }">
              <span>Text</span>
              <span>Text</span>
            </div>
          </div>
        </div>
        <p class="cell-format-dialog__hint">Choose a line style and color, then use presets or click individual preview edges. Presets apply across the current selection.</p>
      </section>
    </div>
    <template #footer>
      <Button label="Cancel" text @click="formatDialogVisible = false" />
      <Button label="Apply" icon="pi pi-check" :disabled="readOnly || (formatDialogKind === 'custom' && !formatDialogCustomNumFmt.trim())" @click="applySelectedCellFormat" />
    </template>
  </Dialog>
</template>

<style scoped>
.cell-format-dialog {
  display: grid;
  gap: 1rem;
}

.cell-format-dialog__hint {
  margin: 0;
  color: var(--muted);
  line-height: 1.5;
}

.cell-format-dialog__sample {
  padding: 0.55rem 0.7rem;
  border: 1px solid var(--border);
  border-radius: 0.65rem;
  background: rgb(255 255 255 / 0.04);
}

/*
 * When a hint sits inside the auto-fit grid panel (Alignment, Font, Fill)
 * it would otherwise be slotted into a single grid column. Force it to
 * span the entire row so the explanatory copy reads naturally below the
 * controls instead of wrapping into a narrow column.
 */
.cell-format-dialog__grid > .cell-format-dialog__hint,
.cell-format-dialog__border > .cell-format-dialog__hint {
  grid-column: 1 / -1;
}

.cell-format-dialog__tabs {
  display: flex;
  flex-wrap: wrap;
  gap: 0.25rem;
  padding: 0.25rem;
  border: 1px solid var(--border);
  border-radius: 999px;
  background: rgb(255 255 255 / 0.04);
}

.cell-format-dialog__tabs button,
.cell-format-dialog__presets button,
.cell-format-dialog__border-toggle {
  border: 0;
  border-radius: 999px;
  background: transparent;
  color: var(--text);
  cursor: pointer;
}

.cell-format-dialog__tabs button {
  padding: 0.45rem 0.8rem;
}

.cell-format-dialog__tab--active {
  background: rgb(212 160 23 / 0.22) !important;
}

.cell-format-dialog__panel {
  display: grid;
  gap: 1rem;
}

.cell-format-dialog__grid {
  grid-template-columns: repeat(auto-fit, minmax(12rem, 1fr));
}

.cell-format-dialog__checks {
  grid-column: 1 / -1;
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 0.75rem;
}

.cell-format-dialog__checks label {
  display: inline-flex;
  gap: 0.35rem;
  align-items: center;
}

.cell-format-dialog__font-family,
.cell-format-dialog__font-size {
  width: 100%;
}

.cell-format-dialog__font-family-placeholder {
  color: var(--muted);
}

.cell-format-dialog__border {
  grid-template-columns: minmax(12rem, 0.8fr) minmax(18rem, 1.2fr);
}

.cell-format-dialog__border-tools,
.cell-format-dialog__presets {
  display: grid;
  gap: 0.75rem;
  align-content: start;
}

.cell-format-dialog__border-main {
  display: grid;
  gap: 1rem;
}

.cell-format-dialog__presets {
  grid-template-columns: repeat(4, minmax(0, 1fr));
}

.cell-format-dialog__presets span {
  grid-column: 1 / -1;
  color: var(--muted);
}

.cell-format-dialog__presets button,
.cell-format-dialog__border-toggle {
  padding: 0.45rem 0.6rem;
  border: 1px solid var(--border);
  background: rgb(255 255 255 / 0.04);
}

.cell-format-dialog__border-preview {
  position: relative;
  display: grid;
  place-items: center;
  min-height: 14rem;
  padding: 2.4rem;
  border: 1px solid var(--border);
  border-radius: 0.85rem;
}

.cell-format-dialog__border-sample {
  display: grid;
  grid-template-columns: 1fr 1fr;
  place-items: center;
  width: 14rem;
  height: 8rem;
  color: var(--text);
  outline: 1px dashed var(--border);
}

.cell-format-dialog__border-toggle {
  position: absolute;
}

.cell-format-dialog__border-toggle--top {
  top: 0.6rem;
}

.cell-format-dialog__border-toggle--right {
  right: 0.6rem;
}

.cell-format-dialog__border-toggle--bottom {
  bottom: 0.6rem;
}

.cell-format-dialog__border-toggle--left {
  left: 0.6rem;
}

@media (max-width: 720px) {
  .cell-format-dialog__border {
    grid-template-columns: 1fr;
  }
}
</style>
