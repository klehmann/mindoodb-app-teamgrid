export interface OoxmlRelationship {
  id: string;
  type: string;
  target: string;
}

export function parseXml(source: string, mimeType: DOMParserSupportedType = "application/xml") {
  return new DOMParser().parseFromString(source, mimeType);
}

export function serializeXml(document: Document) {
  return new XMLSerializer().serializeToString(document);
}

export function child(element: Element | Document, localName: string) {
  return children(element, localName)[0] ?? null;
}

export function children(element: Element | Document, localName: string) {
  return Array.from(element.childNodes)
    .filter((node): node is Element => node.nodeType === Node.ELEMENT_NODE && (node as Element).localName === localName);
}

export function descendants(element: Element | Document, localName: string) {
  return Array.from(element.getElementsByTagName("*")).filter((item) => item.localName === localName);
}

export function textOf(element: Element | Document | null, localName: string) {
  const value = element ? child(element, localName)?.textContent?.trim() : "";
  return value || null;
}

export function attr(element: Element | null | undefined, name: string) {
  if (!element) {
    return null;
  }
  return element.getAttribute(name)
    ?? element.getAttribute(`r:${name}`)
    ?? Array.from(element.attributes).find((attribute) => attribute.localName === name)?.value
    ?? null;
}

export function readRelationships(xml: string): OoxmlRelationship[] {
  const document = parseXml(xml);
  return descendants(document, "Relationship").flatMap((relationship) => {
    const id = relationship.getAttribute("Id");
    const type = relationship.getAttribute("Type");
    const target = relationship.getAttribute("Target");
    return id && type && target ? [{ id, type, target }] : [];
  });
}

export function upsertChildText(document: Document, parent: Element, namespace: string, qualifiedName: string, text: string) {
  let node = child(parent, qualifiedName.split(":").pop() ?? qualifiedName);
  if (!node) {
    node = document.createElementNS(namespace, qualifiedName);
    parent.appendChild(node);
  }
  node.textContent = text;
}
