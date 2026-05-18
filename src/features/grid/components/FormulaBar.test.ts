import { flushPromises, mount } from "@vue/test-utils";
import { describe, expect, it } from "vitest";
import PrimeVue from "primevue/config";

import FormulaBar from "@/features/grid/components/FormulaBar.vue";

function mountFormulaBar(props: Partial<{ modelValue: string; activeAddress: string; readonly: boolean; errorMessage: string | null }> = {}) {
  return mount(FormulaBar, {
    props: {
      modelValue: "",
      activeAddress: "A1",
      readonly: false,
      errorMessage: null,
      ...props,
    },
    global: { plugins: [PrimeVue] },
    attachTo: document.body,
  });
}

describe("FormulaBar", () => {
  it("renders the active address and current draft", () => {
    const wrapper = mountFormulaBar({ modelValue: "=SUM(A1:B1)", activeAddress: "C2" });

    expect(wrapper.find(".formula-bar__address").text()).toBe("C2");
    expect(wrapper.find<HTMLInputElement>(".formula-bar__input").element.value).toBe("=SUM(A1:B1)");
    wrapper.unmount();
  });

  it("emits begin-edit and update:modelValue on input and commit on Enter", async () => {
    const wrapper = mountFormulaBar({ modelValue: "" });
    const input = wrapper.find<HTMLInputElement>(".formula-bar__input");

    await input.setValue("=1+2");

    expect(wrapper.emitted("begin-edit")).toHaveLength(1);
    expect(wrapper.emitted<[string]>("update:modelValue")?.at(-1)?.[0]).toBe("=1+2");

    await input.trigger("keydown.enter");
    expect(wrapper.emitted<[string]>("commit")?.at(-1)?.[0]).toBe("=1+2");

    await input.trigger("keydown.escape");
    expect(wrapper.emitted("cancel")).toHaveLength(1);
    wrapper.unmount();
  });

  it("surfaces an error message under the input when provided", () => {
    const wrapper = mountFormulaBar({ errorMessage: "Invalid formula" });
    expect(wrapper.find(".formula-bar__error").text()).toBe("Invalid formula");
    wrapper.unmount();
  });

  it("disables actions and the input in readonly mode", async () => {
    const wrapper = mountFormulaBar({ readonly: true });
    expect(wrapper.find<HTMLInputElement>(".formula-bar__input").element.readOnly).toBe(true);

    const buttons = document.body.querySelectorAll<HTMLButtonElement>(".formula-bar__actions button");
    buttons.forEach((button) => expect(button.hasAttribute("disabled")).toBe(true));
    await flushPromises();
    wrapper.unmount();
  });
});
