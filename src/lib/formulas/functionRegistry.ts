export type FunctionCategory = "Math" | "Statistics" | "Text" | "Date";

export interface FunctionDefinition {
  name: string;
  category: FunctionCategory;
  signature: string;
  description: string;
  minArgs: number;
  maxArgs?: number;
  evaluate(args: FormulaRuntimeValue[]): FormulaRuntimeValue;
}

export type FormulaRuntimeValue =
  | { kind: "number"; value: number }
  | { kind: "string"; value: string }
  | { kind: "date"; isoDate: string }
  | { kind: "empty" }
  | { kind: "error"; code: "#REF!" | "#CYCLE!" | "#VALUE!" | "#NAME?" | "#DIV/0!" };

export const FUNCTION_REGISTRY: Record<string, FunctionDefinition> = {
  SUM: {
    name: "SUM",
    category: "Math",
    signature: "SUM(number1, [number2], ...)",
    description: "Adds all numeric arguments and cells in ranges.",
    minArgs: 1,
    evaluate(args) {
      return { kind: "number", value: flattenNumbers(args).reduce((sum, value) => sum + value, 0) };
    },
  },
  AVERAGE: {
    name: "AVERAGE",
    category: "Statistics",
    signature: "AVERAGE(number1, [number2], ...)",
    description: "Returns the average of numeric arguments and cells in ranges.",
    minArgs: 1,
    evaluate(args) {
      const numbers = flattenNumbers(args);
      return numbers.length === 0 ? { kind: "error", code: "#VALUE!" } : { kind: "number", value: numbers.reduce((sum, value) => sum + value, 0) / numbers.length };
    },
  },
  MIN: {
    name: "MIN",
    category: "Statistics",
    signature: "MIN(number1, [number2], ...)",
    description: "Returns the smallest numeric argument.",
    minArgs: 1,
    evaluate(args) {
      const numbers = flattenNumbers(args);
      return numbers.length === 0 ? { kind: "error", code: "#VALUE!" } : { kind: "number", value: Math.min(...numbers) };
    },
  },
  MAX: {
    name: "MAX",
    category: "Statistics",
    signature: "MAX(number1, [number2], ...)",
    description: "Returns the largest numeric argument.",
    minArgs: 1,
    evaluate(args) {
      const numbers = flattenNumbers(args);
      return numbers.length === 0 ? { kind: "error", code: "#VALUE!" } : { kind: "number", value: Math.max(...numbers) };
    },
  },
  COUNT: {
    name: "COUNT",
    category: "Statistics",
    signature: "COUNT(value1, [value2], ...)",
    description: "Counts numeric arguments and numeric cells in ranges.",
    minArgs: 1,
    evaluate(args) {
      return { kind: "number", value: flattenNumbers(args).length };
    },
  },
  CONCAT: {
    name: "CONCAT",
    category: "Text",
    signature: "CONCAT(text1, [text2], ...)",
    description: "Joins values as text.",
    minArgs: 1,
    evaluate(args) {
      return { kind: "string", value: args.map((arg) => stringifyRuntimeValue(arg)).join("") };
    },
  },
  TODAY: {
    name: "TODAY",
    category: "Date",
    signature: "TODAY()",
    description: "Returns today's date.",
    minArgs: 0,
    maxArgs: 0,
    evaluate() {
      const date = new Date();
      date.setHours(0, 0, 0, 0);
      return { kind: "date", isoDate: date.toISOString() };
    },
  },
};

export function suggestFunctions(query: string) {
  const normalizedQuery = query.trim().toUpperCase();
  return Object.values(FUNCTION_REGISTRY)
    .filter((definition) => normalizedQuery === "" || definition.name.startsWith(normalizedQuery))
    .sort((left, right) => left.name.localeCompare(right.name));
}

function flattenNumbers(values: FormulaRuntimeValue[]) {
  return values.flatMap((value) => {
    if (value.kind === "number") {
      return [value.value];
    }
    if (value.kind === "string") {
      const number = Number(value.value);
      return Number.isFinite(number) ? [number] : [];
    }
    return [];
  });
}

function stringifyRuntimeValue(value: FormulaRuntimeValue) {
  switch (value.kind) {
    case "number":
      return String(value.value);
    case "string":
      return value.value;
    case "date":
      return value.isoDate;
    case "error":
      return value.code;
    default:
      return "";
  }
}
