declare module "fast-formula-parser" {
  export interface FormulaPosition {
    row: number;
    col: number;
    sheet?: string;
  }

  export interface FormulaCellRef {
    row: number;
    col: number;
    sheet?: string;
  }

  export interface FormulaRangeRef {
    sheet?: string;
    from: {
      row?: number;
      col?: number;
    };
    to: {
      row?: number;
      col?: number;
    };
  }

  export type FormulaReference = FormulaCellRef | FormulaRangeRef;

  export class FormulaError extends Error {
    static DIV0: FormulaError;
    static ERROR(message: string, details?: unknown): FormulaError;
    static NAME: FormulaError;
    static NA: FormulaError;
    static NULL: FormulaError;
    static NUM: FormulaError;
    static REF: FormulaError;
    static VALUE: FormulaError;
    static NOT_IMPLEMENTED(name: string): FormulaError;
    details?: unknown;
    equals?(other: FormulaError): boolean;
  }

  export class DepParser {
    constructor(config?: { onVariable?: (name: string, sheetName?: string) => FormulaReference | null });
    parse(inputText: string, position: FormulaPosition, ignoreError?: boolean): FormulaReference[];
  }

  export interface FormulaParserConfig {
    functions?: Record<string, (...args: unknown[]) => unknown>;
    functionsNeedContext?: Record<string, (...args: unknown[]) => unknown>;
    onVariable?: (name: string, sheetName?: string, position?: FormulaPosition) => FormulaReference | null;
    onCell?: (ref: FormulaCellRef) => unknown;
    onRange?: (ref: FormulaRangeRef) => unknown[][];
  }

  export default class FormulaParser {
    static MAX_ROW: number;
    static MAX_COLUMN: number;
    static DepParser: typeof DepParser;
    static FormulaError: typeof FormulaError;
    constructor(config?: FormulaParserConfig, isTest?: boolean);
    parse(inputText: string, position?: FormulaPosition, allowReturnArray?: boolean): unknown;
    parseAsync(inputText: string, position?: FormulaPosition, allowReturnArray?: boolean): Promise<unknown>;
    supportedFunctions(): string[];
  }
}
