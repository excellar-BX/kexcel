// A small XML reader for the parts of .xlsx files Kexcel understands. It runs
// anywhere (no DOMParser), keeps element order, and drops namespace prefixes
// from element names so "x:row" and "row" read the same.

export interface XmlNode {
  name: string;
  attrs: Record<string, string>;
  children: XmlNode[];
  /** Text directly inside this element, entities decoded. */
  text: string;
}

const ENTITIES: Record<string, string> = { lt: "<", gt: ">", amp: "&", quot: '"', apos: "'" };

export function decodeEntities(s: string): string {
  return s.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (m, e: string) => {
    if (e[0] === "#") {
      const code = e[1] === "x" || e[1] === "X" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return Number.isFinite(code) ? String.fromCodePoint(code) : m;
    }
    return ENTITIES[e.toLowerCase()] ?? m;
  });
}

export function escapeXml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    // Characters XML 1.0 can't hold at all.
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
}

const localName = (name: string) => name.slice(name.indexOf(":") + 1);
const ATTR_RE = /([^\s=]+)\s*=\s*("([^"]*)"|'([^']*)')/g;

export function parseXml(src: string): XmlNode {
  const root: XmlNode = { name: "#document", attrs: {}, children: [], text: "" };
  const stack: XmlNode[] = [root];
  let i = 0;
  while (i < src.length) {
    const lt = src.indexOf("<", i);
    const top = stack[stack.length - 1];
    if (lt < 0) {
      top.text += decodeEntities(src.slice(i));
      break;
    }
    if (lt > i) top.text += decodeEntities(src.slice(i, lt));

    if (src.startsWith("<!--", lt)) {
      const end = src.indexOf("-->", lt);
      i = end < 0 ? src.length : end + 3;
      continue;
    }
    if (src.startsWith("<![CDATA[", lt)) {
      const end = src.indexOf("]]>", lt);
      top.text += src.slice(lt + 9, end < 0 ? src.length : end);
      i = end < 0 ? src.length : end + 3;
      continue;
    }
    if (src[lt + 1] === "?" || src[lt + 1] === "!") {
      const end = src.indexOf(">", lt);
      i = end < 0 ? src.length : end + 1;
      continue;
    }

    const gt = src.indexOf(">", lt);
    if (gt < 0) break;
    const tag = src.slice(lt + 1, gt);
    i = gt + 1;
    if (tag[0] === "/") {
      // Closing tag: pop back to the matching element, tolerating sloppy nesting.
      const name = localName(tag.slice(1).trim());
      for (let k = stack.length - 1; k > 0; k--) {
        if (stack[k].name === name) {
          stack.length = k;
          break;
        }
      }
      continue;
    }
    const selfClosing = tag.endsWith("/");
    const body = selfClosing ? tag.slice(0, -1) : tag;
    const space = body.search(/\s/);
    const name = localName(space < 0 ? body : body.slice(0, space));
    const attrs: Record<string, string> = {};
    if (space >= 0) {
      for (const m of body.slice(space).matchAll(ATTR_RE)) attrs[m[1]] = decodeEntities(m[3] ?? m[4] ?? "");
    }
    const node: XmlNode = { name, attrs, children: [], text: "" };
    top.children.push(node);
    if (!selfClosing) stack.push(node);
  }
  return root;
}

export const child = (node: XmlNode | undefined, name: string) => node?.children.find((c) => c.name === name);
export const children = (node: XmlNode | undefined, name: string) =>
  node?.children.filter((c) => c.name === name) ?? [];

/** All text inside an element and its descendants, in order. */
export function deepText(node: XmlNode | undefined): string {
  if (!node) return "";
  return node.text + node.children.map(deepText).join("");
}
