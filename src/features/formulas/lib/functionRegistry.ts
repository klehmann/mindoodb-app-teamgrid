/**
 * Registry of built-in formula functions used by both the evaluator and the
 * formula bar's content assist.
 *
 * Each {@link FunctionDefinition} carries enough metadata for editor UX
 * (`signature`, `description`, `category`). Evaluation is delegated to
 * `fast-formula-parser`; curated entries below get richer docs while every
 * engine-supported function still appears in content assist with generic help.
 */
import { SUPPORTED_FORMULA_NAMES } from "@/features/formulas/lib/fastFormulaParserAdapter";

/** Grouping shown in the content-assist help overlay. */
export type FunctionCategory = "Math" | "Statistics" | "Text" | "Date" | "Logical" | "Other";

/** One callable formula function. */
export interface FunctionDefinition {
  name: string;
  category: FunctionCategory;
  signature: string;
  description: string;
  minArgs: number;
  maxArgs?: number;
  evaluate?: (args: FormulaRuntimeValue[]) => FormulaRuntimeValue;
}

/**
 * Discriminated runtime value used inside the evaluator. It's a superset of
 * the persisted {@link CellValue} shapes plus an `error` variant, so the
 * first error short-circuits its enclosing operation.
 */
export type FormulaRuntimeValue =
  | { kind: "number"; value: number }
  | { kind: "string"; value: string }
  | { kind: "date"; isoDate: string }
  | { kind: "empty" }
  | { kind: "error"; code: "#REF!" | "#CYCLE!" | "#VALUE!" | "#NAME?" | "#DIV/0!" };

/** Curated formula metadata, keyed by uppercase name. */
const CURATED_FUNCTION_REGISTRY: Record<string, FunctionDefinition> = {
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
  IF: {
    name: "IF",
    category: "Logical",
    signature: "IF(logical_test, value_if_true, [value_if_false])",
    description: "Returns one value when a condition is true and another when it is false.",
    minArgs: 2,
    maxArgs: 3,
  },
  ROUND: {
    name: "ROUND",
    category: "Math",
    signature: "ROUND(number, num_digits)",
    description: "Rounds a number to a specified number of digits.",
    minArgs: 2,
    maxArgs: 2,
  },
  COUNTIF: {
    name: "COUNTIF",
    category: "Statistics",
    signature: "COUNTIF(range, criteria)",
    description: "Counts cells in a range that match a criterion.",
    minArgs: 2,
    maxArgs: 2,
  },
  LEN: {
    name: "LEN",
    category: "Text",
    signature: "LEN(text)",
    description: "Returns the number of characters in text.",
    minArgs: 1,
    maxArgs: 1,
  },
  LOWER: {
    name: "LOWER",
    category: "Text",
    signature: "LOWER(text)",
    description: "Converts text to lowercase.",
    minArgs: 1,
    maxArgs: 1,
  },
  UPPER: {
    name: "UPPER",
    category: "Text",
    signature: "UPPER(text)",
    description: "Converts text to uppercase.",
    minArgs: 1,
    maxArgs: 1,
  },
  LEFT: {
    name: "LEFT",
    category: "Text",
    signature: "LEFT(text, [num_chars])",
    description: "Returns characters from the start of text.",
    minArgs: 1,
    maxArgs: 2,
  },
  RIGHT: {
    name: "RIGHT",
    category: "Text",
    signature: "RIGHT(text, [num_chars])",
    description: "Returns characters from the end of text.",
    minArgs: 1,
    maxArgs: 2,
  },
  DATE: {
    name: "DATE",
    category: "Date",
    signature: "DATE(year, month, day)",
    description: "Returns the Excel serial number for a date.",
    minArgs: 3,
    maxArgs: 3,
  },
  YEAR: {
    name: "YEAR",
    category: "Date",
    signature: "YEAR(serial_number)",
    description: "Returns the year from an Excel date serial number.",
    minArgs: 1,
    maxArgs: 1,
  },
  MONTH: {
    name: "MONTH",
    category: "Date",
    signature: "MONTH(serial_number)",
    description: "Returns the month from an Excel date serial number.",
    minArgs: 1,
    maxArgs: 1,
  },
  DAY: {
    name: "DAY",
    category: "Date",
    signature: "DAY(serial_number)",
    description: "Returns the day of the month from an Excel date serial number.",
    minArgs: 1,
    maxArgs: 1,
  },
};

/**
 * Filter the registry by name prefix for the content-assist popup.
 *
 * Returns all functions when `query` is empty, otherwise functions whose
 * name starts with the (case-insensitive) query.
 */
export function suggestFunctions(query: string) {
  const normalizedQuery = query.trim().toUpperCase();
  return Object.values(FUNCTION_REGISTRY)
    .filter((definition) => normalizedQuery === "" || definition.name.startsWith(normalizedQuery))
    .sort((left, right) => left.name.localeCompare(right.name));
}

function createFallbackFunctionDefinition(name: string): FunctionDefinition {
  const category = inferFunctionCategory(name);
  return {
    name,
    category,
    signature: `${name}(...)`,
    description: createFallbackDescription(name, category),
    minArgs: 0,
  };
}

function createFallbackDescription(name: string, category: FunctionCategory) {
  const specificDescription = FALLBACK_FUNCTION_DESCRIPTIONS[name];
  if (specificDescription) {
    return specificDescription;
  }

  if (name.endsWith(".DIST")) {
    return `${name} returns a probability distribution value for the given inputs.`;
  }
  if (name.endsWith(".INV")) {
    return `${name} returns the inverse of a probability distribution for the given probability.`;
  }
  if (name.endsWith(".TEST") || name.endsWith("TEST")) {
    return `${name} returns a statistical test result for the supplied data.`;
  }

  switch (category) {
    case "Logical":
      return `${name} evaluates logical conditions and returns a calculated result.`;
    case "Date":
      return `${name} works with dates, times, or Excel date serial values.`;
    case "Text":
      return `${name} transforms, extracts, formats, or compares text values.`;
    case "Statistics":
      return `${name} calculates a statistical result from numbers or ranges.`;
    case "Math":
      return `${name} performs a numeric calculation using values or ranges.`;
    default:
      return `${name} is supported by the formula engine for Excel-compatible worksheets.`;
  }
}

const FALLBACK_FUNCTION_DESCRIPTIONS: Record<string, string> = {
  ABS: "Returns the absolute value of a number.",
  ACOS: "Returns the arccosine of a number, in radians.",
  ACOSH: "Returns the inverse hyperbolic cosine of a number.",
  ACOT: "Returns the arccotangent of a number, in radians.",
  ACOTH: "Returns the inverse hyperbolic cotangent of a number.",
  AND: "Returns TRUE when all supplied conditions are true.",
  ARABIC: "Converts a Roman numeral to an Arabic number.",
  ASIN: "Returns the arcsine of a number, in radians.",
  ASINH: "Returns the inverse hyperbolic sine of a number.",
  ATAN: "Returns the arctangent of a number, in radians.",
  ATAN2: "Returns the arctangent from x and y coordinates.",
  ATANH: "Returns the inverse hyperbolic tangent of a number.",
  AVEDEV: "Returns the average of the absolute deviations from the mean.",
  BASE: "Converts a number into text using the specified base.",
  CEILING: "Rounds a number up to the nearest multiple of significance.",
  CHAR: "Returns the character specified by a character code.",
  CLEAN: "Removes non-printable characters from text.",
  CODE: "Returns the numeric code for the first character in text.",
  COMBIN: "Returns the number of combinations for a given number of items.",
  CONCATENATE: "Joins several text values into one text value.",
  CORREL: "Returns the correlation coefficient between two data sets.",
  COS: "Returns the cosine of an angle given in radians.",
  COSH: "Returns the hyperbolic cosine of a number.",
  COT: "Returns the cotangent of an angle given in radians.",
  COTH: "Returns the hyperbolic cotangent of a number.",
  COUNTBLANK: "Counts empty cells in a range.",
  COUNTIFS: "Counts cells that match multiple criteria.",
  COVARIANCE: "Returns covariance, the average of paired deviations.",
  DATEVALUE: "Converts a date stored as text into an Excel date serial number.",
  DAYS: "Returns the number of days between two dates.",
  DAYS360: "Returns the number of days between dates using a 360-day year.",
  DECIMAL: "Converts text in a given base into a decimal number.",
  DEGREES: "Converts radians to degrees.",
  DELTA: "Returns 1 when two numbers are equal, otherwise 0.",
  DEVSQ: "Returns the sum of squared deviations from the mean.",
  EDATE: "Returns the date a specified number of months before or after a date.",
  EOMONTH: "Returns the last day of the month before or after a date.",
  EVEN: "Rounds a number up to the nearest even integer.",
  EXACT: "Returns TRUE when two text values are exactly the same.",
  EXP: "Returns e raised to the power of a number.",
  FACT: "Returns the factorial of a number.",
  FALSE: "Returns the logical value FALSE.",
  FIND: "Returns the position of text within another text value, case-sensitive.",
  FIXED: "Formats a number as text with a fixed number of decimals.",
  FLOOR: "Rounds a number down to the nearest multiple of significance.",
  FORECAST: "Returns a value predicted from existing x and y values.",
  FREQUENCY: "Returns how often values occur within specified bins.",
  GCD: "Returns the greatest common divisor of two or more integers.",
  GEOMEAN: "Returns the geometric mean of positive numbers.",
  HARMEAN: "Returns the harmonic mean of positive numbers.",
  HOUR: "Returns the hour from a time value.",
  IFERROR: "Returns a fallback value when a formula evaluates to an error.",
  IFNA: "Returns a fallback value when a formula evaluates to #N/A.",
  IFS: "Checks multiple conditions and returns the first matching result.",
  INT: "Rounds a number down to the nearest integer.",
  INTERCEPT: "Returns the intercept of a linear regression line.",
  ISOWEEKNUM: "Returns the ISO week number for a date.",
  KURT: "Returns the kurtosis of a data set.",
  LCM: "Returns the least common multiple of two or more integers.",
  LN: "Returns the natural logarithm of a number.",
  LOG: "Returns the logarithm of a number using the specified base.",
  LOG10: "Returns the base-10 logarithm of a number.",
  MAXIFS: "Returns the largest value among cells that match criteria.",
  MEDIAN: "Returns the median of the supplied numbers.",
  MID: "Returns characters from the middle of a text value.",
  MINIFS: "Returns the smallest value among cells that match criteria.",
  MINUTE: "Returns the minute from a time value.",
  MOD: "Returns the remainder after division.",
  MROUND: "Rounds a number to the nearest specified multiple.",
  MULTINOMIAL: "Returns the multinomial of a set of numbers.",
  NETWORKDAYS: "Returns the number of whole working days between two dates.",
  NOT: "Reverses TRUE to FALSE, or FALSE to TRUE.",
  NOW: "Returns the current date and time.",
  ODD: "Rounds a number up to the nearest odd integer.",
  OR: "Returns TRUE when at least one supplied condition is true.",
  PEARSON: "Returns the Pearson correlation coefficient between two data sets.",
  PI: "Returns the value of pi.",
  POWER: "Raises a number to a specified power.",
  PRODUCT: "Multiplies numbers and returns the product.",
  PROPER: "Capitalizes the first letter of each word in text.",
  QUARTILE: "Returns the quartile of a data set.",
  QUOTIENT: "Returns the integer portion of a division.",
  RADIANS: "Converts degrees to radians.",
  RAND: "Returns a random number between 0 and 1.",
  RANDBETWEEN: "Returns a random integer between two numbers.",
  RANK: "Returns the rank of a number within a list of numbers.",
  REPLACE: "Replaces part of a text value with different text.",
  REPT: "Repeats text a specified number of times.",
  ROMAN: "Converts an Arabic number to a Roman numeral.",
  SEARCH: "Returns the position of text within another text value, case-insensitive.",
  SECOND: "Returns the second from a time value.",
  SIGN: "Returns the sign of a number as -1, 0, or 1.",
  SIN: "Returns the sine of an angle given in radians.",
  SINH: "Returns the hyperbolic sine of a number.",
  SKEW: "Returns the skewness of a data set.",
  SLOPE: "Returns the slope of a linear regression line.",
  SQRT: "Returns the positive square root of a number.",
  STANDARDIZE: "Returns a normalized value from a mean and standard deviation.",
  SUBTOTAL: "Returns a subtotal using the specified aggregate function.",
  SUMIF: "Adds cells that match a criterion.",
  SUMIFS: "Adds cells that match multiple criteria.",
  SUMPRODUCT: "Multiplies corresponding array values and returns the sum.",
  T: "Returns text when the value is text, otherwise an empty string.",
  TAN: "Returns the tangent of an angle given in radians.",
  TANH: "Returns the hyperbolic tangent of a number.",
  TEXT: "Formats a value as text using a number format.",
  TIME: "Returns the Excel serial number for a time.",
  TIMEVALUE: "Converts a time stored as text into an Excel time serial number.",
  TRIM: "Removes extra spaces from text.",
  TRUE: "Returns the logical value TRUE.",
  TRUNC: "Truncates a number to an integer or fixed decimal precision.",
  WEEKDAY: "Returns the day of the week for a date.",
  WEEKNUM: "Returns the week number for a date.",
  WORKDAY: "Returns a working day before or after a start date.",
  XOR: "Returns TRUE when an odd number of supplied conditions are true.",
  YEARFRAC: "Returns the fraction of a year between two dates.",
};

/** Built-in functions, keyed by uppercase name. */
export const FUNCTION_REGISTRY: Record<string, FunctionDefinition> = Object.fromEntries(
  SUPPORTED_FORMULA_NAMES.map((name) => [name, CURATED_FUNCTION_REGISTRY[name] ?? createFallbackFunctionDefinition(name)]),
);

function inferFunctionCategory(name: string): FunctionCategory {
  if (/^(AND|FALSE|IF|IFERROR|IFNA|IFS|NOT|OR|TRUE|XOR)$/.test(name)) {
    return "Logical";
  }
  if (/^(DATE|DATEDIF|DATEVALUE|DAY|DAYS|DAYS360|EDATE|EOMONTH|HOUR|ISOWEEKNUM|MINUTE|MONTH|NETWORKDAYS|NOW|SECOND|TIME|TIMEVALUE|TODAY|WEEKDAY|WEEKNUM|WORKDAY|YEAR|YEARFRAC)/.test(name)) {
    return "Date";
  }
  if (/^(ASC|BAHTTEXT|CHAR|CLEAN|CODE|CONCAT|CONCATENATE|DBCS|DOLLAR|EXACT|FIND|FIXED|LEFT|LEN|LOWER|MID|NUMBERVALUE|PROPER|REPLACE|REPT|RIGHT|SEARCH|T|TEXT|TRIM|UNICHAR|UNICODE|UPPER)/.test(name)) {
    return "Text";
  }
  if (/^(AVEDEV|AVERAGE|CORREL|COUNT|COUNTIF|COVARIANCE|DEVSQ|EXPON|FISHER|FORECAST|FREQUENCY|GAMMA|GAUSS|GEOMEAN|GROWTH|HARMEAN|HYPGEOM|INTERCEPT|KURT|LOGNORM|MAX|MEDIAN|MIN|NEGBINOM|NORM|PEARSON|POISSON|QUARTILE|RANK|SKEW|SLOPE|STANDARDIZE|STDEV|T\.|VAR|WEIBULL)/.test(name)) {
    return "Statistics";
  }
  if (/^(ABS|ACOS|ACOSH|ACOT|ARABIC|ASIN|ATAN|BASE|CEILING|COMBIN|COS|DECIMAL|DEGREES|EVEN|EXP|FACT|FLOOR|GCD|INT|LCM|LN|LOG|MOD|MROUND|MULTINOMIAL|ODD|PI|POWER|PRODUCT|QUOTIENT|RADIANS|RAND|ROMAN|ROUND|SEC|SERIESSUM|SIGN|SIN|SQRT|SUBTOTAL|SUM|TAN|TRUNC)/.test(name)) {
    return "Math";
  }
  return "Other";
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
