import { readonly, ref } from "vue";

const TEAMGRID_SERVICE_WORKER_URL = `${import.meta.env.BASE_URL}sw.js`;
const TEAMGRID_SERVICE_WORKER_SCOPE = import.meta.env.BASE_URL || "/";

const updateAvailable = ref(false);
const updateDismissedForSession = ref(false);
const updateReloading = ref(false);
const updateRegistration = ref<ServiceWorkerRegistration | null>(null);

let registrationStarted = false;
let controllerChangeWired = false;

function notifyUpdateAvailable(registration: ServiceWorkerRegistration) {
  updateRegistration.value = registration;
  updateReloading.value = false;
  updateAvailable.value = !updateDismissedForSession.value;
}

function observeInstallingWorker(registration: ServiceWorkerRegistration, worker: ServiceWorker | null) {
  if (!worker) {
    return;
  }
  worker.addEventListener("statechange", () => {
    if (worker.state === "installed" && navigator.serviceWorker.controller) {
      notifyUpdateAvailable(registration);
    }
  });
}

function wireControllerReload() {
  if (controllerChangeWired) {
    return;
  }
  controllerChangeWired = true;
  let reloadHandled = false;
  navigator.serviceWorker.addEventListener("controllerchange", () => {
    if (reloadHandled || !updateReloading.value) {
      return;
    }
    reloadHandled = true;
    window.location.reload();
  });
}

function wireUpdateChecks(registration: ServiceWorkerRegistration) {
  if (registration.waiting) {
    notifyUpdateAvailable(registration);
  }
  observeInstallingWorker(registration, registration.installing);
  registration.addEventListener("updatefound", () => {
    observeInstallingWorker(registration, registration.installing);
  });

  const triggerUpdateCheck = () => {
    void registration.update().catch((error) => {
      console.warn("Could not check for a newer MindooDB TeamGrid version.", error);
    });
  };

  triggerUpdateCheck();
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") {
      triggerUpdateCheck();
    }
  });
  window.setInterval(triggerUpdateCheck, 5 * 60 * 1000);
}

export async function registerTeamGridServiceWorker() {
  if (registrationStarted || import.meta.env.DEV || typeof window === "undefined" || !("serviceWorker" in navigator)) {
    return null;
  }

  registrationStarted = true;
  try {
    const registration = await navigator.serviceWorker.register(TEAMGRID_SERVICE_WORKER_URL, {
      scope: TEAMGRID_SERVICE_WORKER_SCOPE,
      updateViaCache: "none",
    });
    wireControllerReload();
    wireUpdateChecks(registration);
    return registration;
  } catch (error) {
    console.warn("Could not register the MindooDB TeamGrid service worker.", error);
    return null;
  }
}

export function reloadForTeamGridUpdate() {
  updateReloading.value = true;
  if (updateRegistration.value?.waiting) {
    updateRegistration.value.waiting.postMessage({ type: "SKIP_WAITING" });
    window.setTimeout(() => {
      window.location.reload();
    }, 10000);
    return;
  }
  window.location.reload();
}

export function dismissTeamGridUpdate() {
  updateAvailable.value = false;
  updateDismissedForSession.value = true;
  updateReloading.value = false;
}

export function useTeamGridAppUpdate() {
  return {
    updateAvailable: readonly(updateAvailable),
    updateReloading: readonly(updateReloading),
    reloadForUpdate: reloadForTeamGridUpdate,
    dismissUpdate: dismissTeamGridUpdate,
  };
}
