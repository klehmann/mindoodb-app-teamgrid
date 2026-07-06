// Canonical MindooDB pnpmfile — kept byte-for-byte identical across every
// deployed app so its checksum (recorded in pnpm-lock.yaml) is the same
// everywhere and a single edit can be propagated with one copy.
//
// DEFAULT (no flag): this is a pure no-op. Every first-party package resolves
// from the npm registry, so the committed lockfile and the Cloudflare Pages
// build never see a local `file:` path.
//
// LOCAL_MINDOODB=1 pnpm install: redirect the first-party packages below to
// the newest `pnpm pack` tarball found in the sibling checkouts
// (../mindoodb, ../mindoodb-app-sdk, ../mindoodb-view-language), so you can
// test not-yet-published builds without touching package.json. This rewrites
// the lockfile to `file:` paths — NEVER commit a lockfile produced this way.
const fs = require("node:fs");
const path = require("node:path");

const useLocal = process.env.LOCAL_MINDOODB === "1";

// First-party packages, mapped to the sibling directory that holds their
// `pnpm pack` tarballs (relative to this file).
const localPackages = {
  mindoodb: "../mindoodb",
  "mindoodb-app-sdk": "../mindoodb-app-sdk",
  "mindoodb-view-language": "../mindoodb-view-language",
};

// Compare two semver-ish version strings (e.g. "0.0.29" vs "0.0.10").
// Returns > 0 when `a` is newer than `b`.
function compareVersions(a, b) {
  const pa = a.split(/[.-]/);
  const pb = b.split(/[.-]/);
  const len = Math.max(pa.length, pb.length);
  for (let i = 0; i < len; i++) {
    const na = Number.parseInt(pa[i], 10);
    const nb = Number.parseInt(pb[i], 10);
    const aNum = Number.isNaN(na) ? -1 : na;
    const bNum = Number.isNaN(nb) ? -1 : nb;
    if (aNum !== bNum) return aNum - bNum;
  }
  return 0;
}

// Find the newest `<name>-<version>.tgz` in `dir` and return its `file:` spec,
// or null when no matching tarball exists (e.g. on CI or fresh clones).
function findLatestTarball(name, dir) {
  const absDir = path.resolve(__dirname, dir);
  let entries;
  try {
    entries = fs.readdirSync(absDir);
  } catch {
    return null;
  }
  const pattern = new RegExp(`^${name}-(\\d+(?:\\.\\d+)*(?:-[\\w.]+)?)\\.tgz$`);
  let best = null;
  for (const entry of entries) {
    const match = entry.match(pattern);
    if (!match) continue;
    if (best === null || compareVersions(match[1], best.version) > 0) {
      best = { version: match[1], file: entry };
    }
  }
  return best ? `file:${path.resolve(absDir, best.file)}` : null;
}

const localTarballs = {};
if (useLocal) {
  for (const [name, dir] of Object.entries(localPackages)) {
    const spec = findLatestTarball(name, dir);
    if (spec) localTarballs[name] = spec;
  }
}

function redirect(deps) {
  if (!deps) return;
  for (const [name, spec] of Object.entries(localTarballs)) {
    if (deps[name]) deps[name] = spec;
  }
}

function readPackage(pkg) {
  if (!useLocal) return pkg;
  redirect(pkg.dependencies);
  redirect(pkg.devDependencies);
  return pkg;
}

module.exports = { hooks: { readPackage } };
