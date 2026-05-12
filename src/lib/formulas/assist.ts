/**
 * Caret-aware helpers for formula content assist.
 *
 * These functions intentionally operate on plain strings instead of parser
 * nodes. Content assist often runs while the formula is incomplete, so it
 * needs to tolerate fragments such as `=SU` or `=SUM(A1,`.
 */

export interface FormulaInsertResult {
  next: string;
  nextCaret: number;
}

/**
 * Return the identifier fragment directly before the caret.
 *
 * The fragment is only considered function-like when it appears at the start
 * of a formula expression or after an operator, comma, or opening paren. This
 * prevents cell references such as `B1` from being treated as function names.
 */
export function getActiveFunctionFragment(draft: string, caretPos = draft.length) {
  const beforeCaret = draft.slice(0, clampCaret(draft, caretPos));
  const match = /(?:^|[=(,+\-*/])\s*([A-Za-z_][A-Za-z0-9_]*)$/.exec(beforeCaret);
  const fragment = match?.[1] ?? "";
  return /^[A-Za-z]+[1-9][0-9]*$/.test(fragment) ? "" : fragment;
}

/**
 * Return the name of the innermost function call enclosing the caret.
 *
 * The scan is deliberately lightweight but quote-aware. It keeps a small stack
 * of `NAME(` openings and pops on `)`, which is enough for signature help in
 * incomplete user drafts.
 */
export function getEnclosingFunctionName(draft: string, caretPos = draft.length) {
  const source = draft.slice(0, clampCaret(draft, caretPos));
  const stack: string[] = [];
  let index = 0;
  let inString = false;

  while (index < source.length) {
    const character = source[index];
    if (character === '"') {
      inString = !inString;
      index += 1;
      continue;
    }
    if (inString) {
      index += 1;
      continue;
    }
    if (/[A-Za-z_]/.test(character)) {
      const match = /^[A-Za-z_][A-Za-z0-9_]*/.exec(source.slice(index));
      const identifier = match?.[0] ?? "";
      const nextIndex = index + identifier.length;
      const nextNonSpace = source.slice(nextIndex).match(/^\s*/)?.[0].length ?? 0;
      if (source[nextIndex + nextNonSpace] === "(") {
        stack.push(identifier.toUpperCase());
        index = nextIndex + nextNonSpace + 1;
        continue;
      }
      index = nextIndex;
      continue;
    }
    if (character === ")") {
      stack.pop();
    }
    index += 1;
  }

  return stack.at(-1) ?? null;
}

/**
 * Insert a function call at the current caret position.
 *
 * If the caret sits after a partial identifier, that fragment is replaced with
 * the selected function name. The returned caret lands just after the opening
 * paren so the user can immediately type arguments.
 */
export function insertFunctionAtCaret(draft: string, caretPos: number, functionName: string): FormulaInsertResult {
  const caret = clampCaret(draft, caretPos);
  const beforeCaret = draft.slice(0, caret);
  const fragmentMatch = /([A-Za-z_][A-Za-z0-9_]*)$/.exec(beforeCaret);
  const fragmentStart = fragmentMatch ? caret - fragmentMatch[1].length : caret;
  const next = `${draft.slice(0, fragmentStart)}${functionName.toUpperCase()}(${draft.slice(caret)}`;
  return {
    next,
    nextCaret: fragmentStart + functionName.length + 1,
  };
}

function clampCaret(draft: string, caretPos: number) {
  return Math.max(0, Math.min(draft.length, caretPos));
}
