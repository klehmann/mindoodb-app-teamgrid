/**
 * Toggle the "Something went wrong" dialog whenever the document composable
 * surfaces a new error, and clear the underlying error message when the
 * user dismisses it.
 */
import { ref, watch, type Ref } from "vue";

export interface UseErrorDialogOptions {
  lastErrorMessage: Ref<string | null>;
  clearLastError: () => void;
}

export function useErrorDialog(options: UseErrorDialogOptions) {
  const errorDialogVisible = ref(false);

  watch(options.lastErrorMessage, (message) => {
    if (message) {
      errorDialogVisible.value = true;
    }
  });

  function dismissErrorDialog() {
    errorDialogVisible.value = false;
    options.clearLastError();
  }

  return {
    errorDialogVisible,
    dismissErrorDialog,
  };
}
