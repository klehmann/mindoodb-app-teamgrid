import { parseCellAddress, type GridProjection } from "@/lib/gridProjection";
import type { FormulaErrorCode, FormulaReference, WorksheetId } from "@/lib/teamgridDocument";
import { FUNCTION_REGISTRY } from "@/lib/formulas/functionRegistry";

export type FormulaNode =
  | { kind: "number"; value: number }
  | { kind: "string"; value: string }
  | { kind: "cell"; address: string; reference: Extract<FormulaReference, { kind: "cell" }> }
  | { kind: "range"; startAddress: string; endAddress: string; reference: Extract<FormulaReference, { kind: "range" }> }
  | { kind: "function"; name: string; args: FormulaNode[] }
  | { kind: "binary"; operator: "+" | "-" | "*" | "/"; left: FormulaNode; right: FormulaNode };

export interface ParsedFormula {
  source: string;
  ast: FormulaNode;
  references: FormulaReference[];
}

export interface FormulaParseError {
  code: FormulaErrorCode;
  message: string;
}

type Token =
  | { kind: "number"; value: string }
  | { kind: "string"; value: string }
  | { kind: "identifier"; value: string }
  | { kind: "operator"; value: "+" | "-" | "*" | "/" | ":" }
  | { kind: "paren"; value: "(" | ")" }
  | { kind: "comma"; value: "," }
  | { kind: "eof"; value: "" };

export function parseFormula(source: string, worksheetId: WorksheetId, projection: GridProjection): ParsedFormula | FormulaParseError {
  try {
    const parser = new FormulaParser(tokenize(source.startsWith("=") ? source.slice(1) : source), worksheetId, projection);
    const ast = parser.parseExpression();
    parser.expect("eof");
    return {
      source: source.startsWith("=") ? source : `=${source}`,
      ast,
      references: collectReferences(ast),
    };
  } catch (error) {
    return {
      code: "#VALUE!",
      message: error instanceof Error ? error.message : "Formula could not be parsed.",
    };
  }
}

class FormulaParser {
  private position = 0;

  constructor(
    private readonly tokens: Token[],
    private readonly worksheetId: WorksheetId,
    private readonly projection: GridProjection,
  ) {}

  parseExpression(): FormulaNode {
    return this.parseAdditive();
  }

  expect(kind: Token["kind"], value?: Token["value"]) {
    const token = this.peek();
    if (token.kind !== kind || (value != null && token.value !== value)) {
      throw new Error(`Expected ${value ?? kind}.`);
    }
    this.position += 1;
    return token;
  }

  private parseAdditive(): FormulaNode {
    let node = this.parseMultiplicative();
    while (this.matchOperator("+") || this.matchOperator("-")) {
      const operator = this.previous().value as "+" | "-";
      node = { kind: "binary", operator, left: node, right: this.parseMultiplicative() };
    }
    return node;
  }

  private parseMultiplicative(): FormulaNode {
    let node = this.parsePrimary();
    while (this.matchOperator("*") || this.matchOperator("/")) {
      const operator = this.previous().value as "*" | "/";
      node = { kind: "binary", operator, left: node, right: this.parsePrimary() };
    }
    return node;
  }

  private parsePrimary(): FormulaNode {
    const token = this.peek();
    if (token.kind === "number") {
      this.position += 1;
      return { kind: "number", value: Number(token.value) };
    }
    if (token.kind === "string") {
      this.position += 1;
      return { kind: "string", value: token.value };
    }
    if (token.kind === "identifier") {
      this.position += 1;
      if (this.matchParen("(")) {
        return this.parseFunctionCall(token.value.toUpperCase());
      }
      return this.parseCellOrRange(token.value.toUpperCase());
    }
    if (this.matchParen("(")) {
      const expression = this.parseExpression();
      this.expect("paren", ")");
      return expression;
    }
    throw new Error("Unexpected formula token.");
  }

  private parseFunctionCall(name: string): FormulaNode {
    if (!FUNCTION_REGISTRY[name]) {
      throw new Error(`Unknown function ${name}.`);
    }
    const args: FormulaNode[] = [];
    if (!this.matchParen(")")) {
      do {
        args.push(this.parseExpression());
      } while (this.matchComma());
      this.expect("paren", ")");
    }
    return { kind: "function", name, args };
  }

  private parseCellOrRange(address: string): FormulaNode {
    const start = parseCellAddress(address, this.projection);
    if (!start) {
      throw new Error(`Unknown cell ${address}.`);
    }
    if (this.matchOperator(":")) {
      const endToken = this.expect("identifier");
      const endAddress = endToken.value.toUpperCase();
      const end = parseCellAddress(endAddress, this.projection);
      if (!end) {
        throw new Error(`Unknown cell ${endAddress}.`);
      }
      return {
        kind: "range",
        startAddress: address,
        endAddress,
        reference: {
          kind: "range",
          worksheetId: this.worksheetId,
          startRowId: start.rowId,
          endRowId: end.rowId,
          startColumnId: start.columnId,
          endColumnId: end.columnId,
        },
      };
    }
    return {
      kind: "cell",
      address,
      reference: {
        kind: "cell",
        worksheetId: this.worksheetId,
        rowId: start.rowId,
        columnId: start.columnId,
      },
    };
  }

  private matchOperator(value: "+" | "-" | "*" | "/" | ":") {
    if (this.peek().kind === "operator" && this.peek().value === value) {
      this.position += 1;
      return true;
    }
    return false;
  }

  private matchParen(value: "(" | ")") {
    if (this.peek().kind === "paren" && this.peek().value === value) {
      this.position += 1;
      return true;
    }
    return false;
  }

  private matchComma() {
    if (this.peek().kind === "comma") {
      this.position += 1;
      return true;
    }
    return false;
  }

  private previous() {
    return this.tokens[this.position - 1];
  }

  private peek() {
    return this.tokens[this.position] ?? { kind: "eof", value: "" as const };
  }
}

function tokenize(source: string): Token[] {
  const tokens: Token[] = [];
  let index = 0;
  while (index < source.length) {
    const character = source[index];
    if (/\s/.test(character)) {
      index += 1;
      continue;
    }
    if (/[0-9.]/.test(character)) {
      const match = /^[0-9]+(?:\.[0-9]+)?/.exec(source.slice(index));
      if (!match) {
        throw new Error("Invalid number.");
      }
      tokens.push({ kind: "number", value: match[0] });
      index += match[0].length;
      continue;
    }
    if (character === "\"") {
      const endIndex = source.indexOf("\"", index + 1);
      if (endIndex === -1) {
        throw new Error("Unterminated string.");
      }
      tokens.push({ kind: "string", value: source.slice(index + 1, endIndex) });
      index = endIndex + 1;
      continue;
    }
    if (/[A-Za-z_]/.test(character)) {
      const match = /^[A-Za-z_][A-Za-z0-9_]*/.exec(source.slice(index));
      if (!match) {
        throw new Error("Invalid identifier.");
      }
      tokens.push({ kind: "identifier", value: match[0] });
      index += match[0].length;
      continue;
    }
    if (character === "," ) {
      tokens.push({ kind: "comma", value: "," });
      index += 1;
      continue;
    }
    if (character === "(" || character === ")") {
      tokens.push({ kind: "paren", value: character });
      index += 1;
      continue;
    }
    if (character === "+" || character === "-" || character === "*" || character === "/" || character === ":") {
      tokens.push({ kind: "operator", value: character });
      index += 1;
      continue;
    }
    throw new Error(`Unexpected character ${character}.`);
  }
  tokens.push({ kind: "eof", value: "" });
  return tokens;
}

function collectReferences(node: FormulaNode): FormulaReference[] {
  switch (node.kind) {
    case "cell":
    case "range":
      return [node.reference];
    case "binary":
      return [...collectReferences(node.left), ...collectReferences(node.right)];
    case "function":
      return node.args.flatMap((arg) => collectReferences(arg));
    default:
      return [];
  }
}
