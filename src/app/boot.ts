import { installBootRecovery } from "@/app/pwa/bootRecovery";

void installBootRecovery()
  .then((bootRecovery) =>
    import("./main").catch((error) => {
      void bootRecovery.reportBootFailure("main-import-failed", error);
    }),
  )
  .catch((error) => {
    console.error("MindooDB TeamGrid boot recovery could not start.", error);
    return import("./main");
  });
