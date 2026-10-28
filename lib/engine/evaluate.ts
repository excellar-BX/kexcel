import { checkNumber, compareScalars, toBoolean, toNumber, toScalar, toText } from "./coerce";
import { FUNCTIONS } from "./functions";
import type { Node } from "./parser";
import { CellError, isError, type Scalar, type Value } from "./types";

export type CellReader = (row: number, col: number) => Scalar;

function evaluateNode(node: Node, read: CellReader): Value {
  switch (node.type) {
    case "number":
    case "string":
    case "boolean":
      return node.value;
    case "ref":
      return read(node.addr.row, node.addr.col);
    case "range": {
      const rows: Scalar[][] = [];
      for (let r = node.start.row; r <= node.end.row; r++) {
        const row: Scalar[] = [];
        for (let c = node.start.col; c <= node.end.col; c++) row.push(read(r, c));
        rows.push(row);
      }
      return rows;
    }
    case "refError":
      throw new CellError("#REF!", "Reference points outside the sheet");
    case "name":
      throw new CellError("#NAME?", `Unknown name "${node.name}"`);
    case "unary": {
      const n = toNumber(evaluateNode(node.arg, read));
      return node.op === "-" ? -n : n;
    }
    case "percent":
      return toNumber(evaluateNode(node.arg, read)) / 100;
    case "binary":
      return evaluateBinary(node, read);
    case "call":
      return evaluateCall(node, read);
  }
}

function evaluateBinary(node: Extract<Node, { type: "binary" }>, read: CellReader): Value {
  const left = evaluateNode(node.left, read);
  const right = evaluateNode(node.right, read);
  switch (node.op) {
    case "+": return checkNumber(toNumber(left) + toNumber(right));
    case "-": return checkNumber(toNumber(left) - toNumber(right));
    case "*": return checkNumber(toNumber(left) * toNumber(right));
    case "/": {
      const divisor = toNumber(right);
      const dividend = toNumber(left);
      if (divisor === 0) throw new CellError("#DIV/0!", "Division by zero");
      return checkNumber(dividend / divisor);
    }
    case "^": return checkNumber(toNumber(left) ** toNumber(right));
    case "&": return toText(left) + toText(right);
  }
  const c = compareScalars(toScalar(left), toScalar(right));
  switch (node.op) {
    case "=": return c === 0;
    case "<>": return c !== 0;
    case "<": return c < 0;
    case ">": return c > 0;
    case "<=": return c <= 0;
    case ">=": return c >= 0;
  }
}

/** Evaluates an argument, turning a thrown error into an error value the function can inspect. */
function evaluateArg(node: Node, read: CellReader): Value {
  try {
    return evaluateNode(node, read);
  } catch (e) {
    if (e instanceof CellError) return e;
    throw e;
  }
}

function evaluateCall(node: Extract<Node, { type: "call" }>, read: CellReader): Value {
  // IF and IFERROR only evaluate the branch they need.
  if (node.name === "IF") {
    if (node.args.length < 2 || node.args.length > 3) {
      throw new CellError("#ERROR!", "IF expects 2-3 arguments");
    }
    const cond = toBoolean(evaluateNode(node.args[0], read));
    const branch = cond ? node.args[1] : node.args[2];
    return branch ? evaluateNode(branch, read) : false;
  }
  if (node.name === "IFERROR") {
    if (node.args.length !== 2) throw new CellError("#ERROR!", "IFERROR expects 2 arguments");
    const v = evaluateArg(node.args[0], read);
    return isError(v) ? evaluateNode(node.args[1], read) : v;
  }

  const fn = FUNCTIONS[node.name];
  if (!fn) throw new CellError("#NAME?", `Unknown function ${node.name}`);
  return fn(node.args.map((a) => evaluateArg(a, read)));
}

/** Evaluates a parsed formula to a single cell value; errors become error values. */
export function evaluateFormula(node: Node, read: CellReader): Scalar {
  try {
    return toScalar(evaluateNode(node, read));
  } catch (e) {
    if (e instanceof CellError) return e;
    return new CellError("#ERROR!", e instanceof Error ? e.message : String(e));
  }
}
