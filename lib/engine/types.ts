export type ErrorCode =
  | "#DIV/0!"
  | "#VALUE!"
  | "#REF!"
  | "#NAME?"
  | "#N/A"
  | "#NUM!"
  | "#CYCLE!"
  | "#ERROR!";

export class CellError {
  readonly code: ErrorCode;
  readonly message?: string;

  constructor(code: ErrorCode, message?: string) {
    this.code = code;
    this.message = message;
  }

  toString() {
    return this.code;
  }
}

/** A single computed cell value. `null` means the cell is empty. */
export type Scalar = number | string | boolean | null | CellError;

/** A rectangular block of values produced by a range reference, rows first. */
export type Matrix = Scalar[][];

export type Value = Scalar | Matrix;

export interface CellAddress {
  row: number;
  col: number;
}

export const isMatrix = (v: Value): v is Matrix => Array.isArray(v);
export const isError = (v: unknown): v is CellError => v instanceof CellError;
