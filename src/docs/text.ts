import { findFile } from "../drive";

export function replaceTextInFile(
  folderPath: string,
  fileName: string,
  searchPattern: string,
  replacementText: string,
): number {
  if (!folderPath || typeof folderPath !== "string") {
    throw new Error("Folder path must be a non-empty string");
  }
  if (!fileName || typeof fileName !== "string") {
    throw new Error("File name must be a non-empty string");
  }
  if (searchPattern === null || searchPattern === undefined || typeof searchPattern !== "string") {
    throw new Error("Search pattern must be a string");
  }
  if (replacementText === null || replacementText === undefined || typeof replacementText !== "string") {
    throw new Error("Replacement text must be a string");
  }
  const file = findFile(folderPath, fileName);
  const doc = DocumentApp.openById(file.getId());
  const body = doc.getBody();
  let baseRegex: RegExp;
  try {
    baseRegex = new RegExp(searchPattern, "g");
  } catch (error) {
    throw new Error(
      `Invalid search pattern "${searchPattern}": ${(error as Error).message}`,
    );
  }
  const patternSource = baseRegex.source;
  const patternFlags = baseRegex.flags;
  let replacements = 0;
  const childCount = body.getNumChildren();
  for (let i = 0; i < childCount; i += 1) {
    const child = body.getChild(i);
    const type = child.getType();
    if (
      type === DocumentApp.ElementType.PARAGRAPH ||
      type === DocumentApp.ElementType.LIST_ITEM
    ) {
      const paragraph = child.asParagraph();
      const text = paragraph.getText();
      const matchRegex = new RegExp(patternSource, patternFlags);
      const matches = text.match(matchRegex);
      if (matches && matches.length > 0) {
        const replaceRegex = new RegExp(patternSource, patternFlags);
        paragraph.setText(text.replace(replaceRegex, replacementText));
        replacements += matches.length;
      }
    }
  }
  doc.saveAndClose();
  Logger.log(
    `Replaced ${replacements} occurrence(s) of "${searchPattern}" in document "${fileName}"`,
  );
  return replacements;
}

