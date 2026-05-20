import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

export type OoxmlZip = Record<string, Uint8Array>;

export function readOoxmlZip(buffer: ArrayBuffer | Uint8Array): OoxmlZip {
  return unzipSync(buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer));
}

export function writeOoxmlZip(zip: OoxmlZip): ArrayBuffer {
  const output = zipSync(zip);
  return output.buffer.slice(output.byteOffset, output.byteOffset + output.byteLength) as ArrayBuffer;
}

export function readZipText(zip: OoxmlZip, path: string) {
  const entry = zip[normalizeZipPath(path)];
  return entry ? strFromU8(entry) : null;
}

export function writeZipText(zip: OoxmlZip, path: string, text: string) {
  zip[normalizeZipPath(path)] = strToU8(text);
}

export function normalizeZipPath(path: string) {
  return path.replace(/^\/+/, "").replace(/\\/g, "/");
}

export function resolveZipTarget(basePath: string, target: string) {
  if (target.startsWith("/")) {
    return normalizeZipPath(target);
  }
  const baseParts = normalizeZipPath(basePath).split("/");
  baseParts.pop();
  for (const part of target.split("/")) {
    if (!part || part === ".") {
      continue;
    }
    if (part === "..") {
      baseParts.pop();
      continue;
    }
    baseParts.push(part);
  }
  return normalizeZipPath(baseParts.join("/"));
}

export function relsPathForPart(partPath: string) {
  const normalized = normalizeZipPath(partPath);
  const parts = normalized.split("/");
  const fileName = parts.pop();
  return `${parts.join("/")}/_rels/${fileName}.rels`;
}
