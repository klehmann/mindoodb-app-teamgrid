const TEAMGRID_SERVICE_WORKER_URL = `${import.meta.env.BASE_URL}sw.js`;
const TEAMGRID_SERVICE_WORKER_SCOPE = import.meta.env.BASE_URL || "/";
const TEAMGRID_BOOT_RECOVERY_ATTEMPTS_KEY = "mindoodb-teamgrid-boot-recovery-attempts";

export const TEAMGRID_BOOT_COMPLETED_EVENT = "mindoodb-teamgrid:boot-complete";

type RecoveryStatus = "checking" | "failed";

type BootRecoveryController = {
  reportBootFailure(reason: string, error?: unknown): Promise<void>;
};

function readRecoveryAttempts() {
  try {
    return Number.parseInt(window.sessionStorage.getItem(TEAMGRID_BOOT_RECOVERY_ATTEMPTS_KEY) || "0", 10) || 0;
  } catch {
    return 0;
  }
}

function writeRecoveryAttempts(value: number) {
  try {
    window.sessionStorage.setItem(TEAMGRID_BOOT_RECOVERY_ATTEMPTS_KEY, String(value));
  } catch {
    // Ignore storage failures so recovery still works in hardened iframe contexts.
  }
}

function clearRecoveryAttempts() {
  try {
    window.sessionStorage.removeItem(TEAMGRID_BOOT_RECOVERY_ATTEMPTS_KEY);
  } catch {
    // Ignore storage failures so recovery still works in hardened iframe contexts.
  }
}

function createRecoveryUi() {
  const root = document.createElement("div");
  root.setAttribute("data-teamgrid-boot-recovery", "true");
  root.innerHTML = `
    <style>
      [data-teamgrid-boot-recovery="true"] {
        position: fixed;
        inset: 0;
        z-index: 2147483647;
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 24px;
        background: linear-gradient(180deg, #081325 0%, #0f1d34 100%);
        color: #f6f8ff;
        font-family: Inter, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      }

      [data-teamgrid-boot-recovery="true"] .teamgrid-boot-recovery__panel {
        width: min(100%, 420px);
        padding: 24px;
        border: 1px solid rgba(255, 255, 255, 0.12);
        border-radius: 20px;
        background: rgba(8, 19, 37, 0.88);
        box-shadow: 0 24px 60px rgba(0, 0, 0, 0.35);
      }

      [data-teamgrid-boot-recovery="true"] .teamgrid-boot-recovery__eyebrow {
        margin: 0 0 8px;
        font-size: 12px;
        font-weight: 700;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        color: #8fc7ff;
      }

      [data-teamgrid-boot-recovery="true"] h1 {
        margin: 0 0 12px;
        font-size: 24px;
        line-height: 1.2;
      }

      [data-teamgrid-boot-recovery="true"] p {
        margin: 0;
        font-size: 15px;
        line-height: 1.6;
        color: rgba(246, 248, 255, 0.8);
      }

      [data-teamgrid-boot-recovery="true"] .teamgrid-boot-recovery__actions {
        display: flex;
        gap: 12px;
        margin-top: 20px;
      }

      [data-teamgrid-boot-recovery="true"] button {
        appearance: none;
        border: 0;
        border-radius: 999px;
        padding: 10px 16px;
        font: inherit;
        font-weight: 700;
        cursor: pointer;
      }

      [data-teamgrid-boot-recovery="true"] .teamgrid-boot-recovery__primary {
        background: #7ec8ff;
        color: #081325;
      }

      [data-teamgrid-boot-recovery="true"] .teamgrid-boot-recovery__secondary {
        background: rgba(255, 255, 255, 0.08);
        color: #f6f8ff;
      }
    </style>
    <div class="teamgrid-boot-recovery__panel">
      <p class="teamgrid-boot-recovery__eyebrow">MindooDB TeamGrid</p>
      <h1>Recovering startup</h1>
      <p data-teamgrid-boot-recovery-message>Checking for an updated offline shell so TeamGrid can reload cleanly.</p>
      <div class="teamgrid-boot-recovery__actions">
        <button type="button" class="teamgrid-boot-recovery__primary" data-teamgrid-boot-recovery-retry hidden>
          Clear cache & reload
        </button>
        <button type="button" class="teamgrid-boot-recovery__secondary" data-teamgrid-boot-recovery-dismiss hidden>
          Dismiss
        </button>
      </div>
    </div>
  `;
  document.body.append(root);

  const message = root.querySelector<HTMLElement>("[data-teamgrid-boot-recovery-message]");
  const retryButton = root.querySelector<HTMLButtonElement>("[data-teamgrid-boot-recovery-retry]");
  const dismissButton = root.querySelector<HTMLButtonElement>("[data-teamgrid-boot-recovery-dismiss]");

  return {
    setStatus(status: RecoveryStatus) {
      if (!message || !retryButton || !dismissButton) {
        return;
      }
      if (status === "checking") {
        message.textContent = "Checking for an updated offline shell so TeamGrid can reload cleanly.";
        retryButton.hidden = true;
        dismissButton.hidden = true;
        return;
      }
      message.textContent =
        "TeamGrid hit a startup error. Clear the offline cache for this app and reload to fetch a fresh copy.";
      retryButton.hidden = false;
      dismissButton.hidden = false;
    },
    onRetry(listener: () => void) {
      retryButton?.addEventListener("click", listener, { once: true });
    },
    onDismiss(listener: () => void) {
      dismissButton?.addEventListener("click", listener, { once: true });
    },
    remove() {
      root.remove();
    },
  };
}

async function waitForInstallingWorker(worker: ServiceWorker | null) {
  if (
    !worker
    || worker.state === "installed"
    || worker.state === "activated"
    || worker.state === "redundant"
  ) {
    return;
  }

  await new Promise<void>((resolve) => {
    const handleStateChange = () => {
      if (worker.state === "installed" || worker.state === "activated" || worker.state === "redundant") {
        worker.removeEventListener("statechange", handleStateChange);
        resolve();
      }
    };
    worker.addEventListener("statechange", handleStateChange);
  });
}

async function waitForControllerChange() {
  return new Promise<boolean>((resolve) => {
    let settled = false;
    const finish = (value: boolean) => {
      if (settled) {
        return;
      }
      settled = true;
      navigator.serviceWorker.removeEventListener("controllerchange", handleControllerChange);
      window.clearTimeout(timeoutId);
      resolve(value);
    };
    const handleControllerChange = () => finish(true);
    const timeoutId = window.setTimeout(() => finish(false), 4000);
    navigator.serviceWorker.addEventListener("controllerchange", handleControllerChange, { once: true });
  });
}

async function updateAndActivateNewestWorker(registration: ServiceWorkerRegistration | null) {
  if (!registration) {
    return false;
  }

  await registration.update().catch((error) => {
    console.warn("Could not check for a newer MindooDB TeamGrid version during startup recovery.", error);
  });
  await waitForInstallingWorker(registration.installing);

  if (!registration.waiting) {
    return false;
  }

  const controllerChanged = waitForControllerChange();
  registration.waiting.postMessage({ type: "SKIP_WAITING" });
  return controllerChanged;
}

async function ensureRecoveryServiceWorker() {
  if (typeof window === "undefined" || !("serviceWorker" in navigator)) {
    return null;
  }

  try {
    return await navigator.serviceWorker.register(TEAMGRID_SERVICE_WORKER_URL, {
      scope: TEAMGRID_SERVICE_WORKER_SCOPE,
      updateViaCache: "none",
    });
  } catch (error) {
    console.warn("Could not register the MindooDB TeamGrid service worker during startup recovery.", error);
    return null;
  }
}

async function clearOfflineCaches() {
  if (!("caches" in window)) {
    return;
  }
  const cacheNames = await caches.keys();
  await Promise.all(cacheNames.map((cacheName) => caches.delete(cacheName)));
}

async function unregisterAppServiceWorkers() {
  if (!("serviceWorker" in navigator)) {
    return;
  }
  const registrations = await navigator.serviceWorker.getRegistrations();
  await Promise.all(registrations.map((registration) => registration.unregister()));
}

async function clearOfflineShellAndReload() {
  clearRecoveryAttempts();
  try {
    await clearOfflineCaches();
    await unregisterAppServiceWorkers();
  } catch (error) {
    console.warn("Could not fully clear the MindooDB TeamGrid offline shell during startup recovery.", error);
  }
  window.location.reload();
}

export async function installBootRecovery(): Promise<BootRecoveryController> {
  if (import.meta.env.DEV || typeof window === "undefined") {
    return {
      async reportBootFailure(reason: string, error?: unknown) {
        console.error("MindooDB TeamGrid startup failed during development.", reason, error);
      },
    };
  }

  const registrationPromise = ensureRecoveryServiceWorker();
  let bootCompleted = false;
  let recovering = false;
  let recoveryUi: ReturnType<typeof createRecoveryUi> | null = null;
  let hardResetWired = false;

  const finishBoot = () => {
    if (bootCompleted) {
      return;
    }
    bootCompleted = true;
    clearRecoveryAttempts();
    recoveryUi?.remove();
    recoveryUi = null;
    window.removeEventListener("error", handleWindowError);
    window.removeEventListener("unhandledrejection", handleUnhandledRejection);
    window.removeEventListener(TEAMGRID_BOOT_COMPLETED_EVENT, handleBootCompleted);
  };

  const attemptRecovery = async (reason: string, error?: unknown) => {
    if (bootCompleted || recovering) {
      return;
    }

    recovering = true;
    recoveryUi ??= createRecoveryUi();
    recoveryUi.setStatus("checking");

    console.error("MindooDB TeamGrid startup recovery triggered.", { reason, error });

    const currentAttempts = readRecoveryAttempts();
    const shouldAutoReload = currentAttempts < 1;
    writeRecoveryAttempts(currentAttempts + 1);

    const registration = await registrationPromise;
    const activatedNewWorker = await updateAndActivateNewestWorker(registration);
    if (activatedNewWorker) {
      window.location.reload();
      return;
    }

    recoveryUi.setStatus("failed");
    if (!hardResetWired) {
      hardResetWired = true;
      recoveryUi.onRetry(() => {
        void clearOfflineShellAndReload();
      });
      recoveryUi.onDismiss(() => {
        recoveryUi?.remove();
        recoveryUi = null;
      });
    }

    recovering = false;
    if (shouldAutoReload) {
      window.location.reload();
    }
  };

  const handleWindowError = (event: ErrorEvent) => {
    void attemptRecovery("window-error", event.error ?? event.message);
  };
  const handleUnhandledRejection = (event: PromiseRejectionEvent) => {
    void attemptRecovery("unhandledrejection", event.reason);
  };
  const handleBootCompleted = () => finishBoot();

  window.addEventListener("error", handleWindowError);
  window.addEventListener("unhandledrejection", handleUnhandledRejection);
  window.addEventListener(TEAMGRID_BOOT_COMPLETED_EVENT, handleBootCompleted, { once: true });

  return {
    reportBootFailure(reason: string, error?: unknown) {
      return attemptRecovery(reason, error);
    },
  };
}
