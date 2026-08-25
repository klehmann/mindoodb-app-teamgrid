import { abbreviateCanonicalName } from "mindoodb-app-sdk";

type EncryptForEntry = {
  kind?: string;
  removedAt?: number;
  label?: string;
};

export function recipientNamesEqual(left: string, right: string): boolean {
  const a = left.trim().toLowerCase();
  const b = right.trim().toLowerCase();
  if (!a || !b) {
    return false;
  }
  if (a === b) {
    return true;
  }
  return (
    abbreviateCanonicalName(left).toLowerCase() ===
    abbreviateCanonicalName(right).toLowerCase()
  );
}

export function isSealedEncryptForDocument(
  data: Record<string, unknown> | null | undefined,
): boolean {
  const raw = data?._encryptFor;
  return Boolean(raw && typeof raw === "object" && !Array.isArray(raw));
}

export function activeEncryptForUsernames(
  data: Record<string, unknown> | null | undefined,
): string[] {
  const raw = data?._encryptFor;
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) {
    return [];
  }
  return Object.entries(raw as Record<string, EncryptForEntry>)
    .filter(
      ([, entry]) =>
        entry &&
        typeof entry === "object" &&
        !entry.removedAt &&
        entry.kind !== "device",
    )
    .map(([id, entry]) => entry.label || id.split("#")[0]);
}

/** Prefer the Directory spelling when it matches a stored `_encryptFor` name. */
export function preferDirectoryUsername(
  name: string,
  directoryUsers: readonly string[],
): string {
  const match = directoryUsers.find((user) => recipientNamesEqual(user, name));
  return match ?? name;
}

/** Active user readers excluding the author (who is always included). */
export function extraEncryptForUsernames(
  data: Record<string, unknown> | null | undefined,
  author: string,
): string[] {
  const authorName = author.trim();
  return activeEncryptForUsernames(data).filter(
    (name) => !authorName || !recipientNamesEqual(name, authorName),
  );
}

export function recipientDiff(
  current: string[],
  next: string[],
): { added: string[]; removed: string[] } {
  const added = next.filter(
    (name) => !current.some((existing) => recipientNamesEqual(existing, name)),
  );
  const removed = current.filter(
    (name) => !next.some((existing) => recipientNamesEqual(existing, name)),
  );
  return { added, removed };
}
