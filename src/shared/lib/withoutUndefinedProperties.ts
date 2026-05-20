/**
 * Recursively drop object properties whose value is `undefined`.
 *
 * Automerge rejects explicit `undefined` assignments, so imported or
 * in-memory structures must omit optional fields instead of setting them
 * to `undefined` before they are written into a document.
 */
export function withoutUndefinedProperties<T>(value: T): T {
  if (Array.isArray(value)) {
    return value.map(withoutUndefinedProperties) as T;
  }
  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value).flatMap(([key, nested]) => (
        nested === undefined ? [] : [[key, withoutUndefinedProperties(nested)]]
      )),
    ) as T;
  }
  return value;
}
