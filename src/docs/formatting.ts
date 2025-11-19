export function formatParagraph(
  paragraph:
    | GoogleAppsScript.Document.Paragraph
    | GoogleAppsScript.Document.ListItem,
  fontFamily: string = "Roboto"
): void {
  paragraph.setAttributes({
    [DocumentApp.Attribute.FONT_FAMILY]: fontFamily,
  });
  paragraph.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);

  // Get the document from the paragraph's parent and save it
  // Paragraph -> Body -> Document (Body's parent is the Document)
  try {
    const parent = paragraph.getParent();
    if (parent && parent.getType() === DocumentApp.ElementType.BODY_SECTION) {
      const body = parent.asBody();
      // Body's parent is the Document - check type and save
      const bodyParent = body.getParent();
      if (
        bodyParent &&
        bodyParent.getType() === DocumentApp.ElementType.DOCUMENT
      ) {
        // TypeScript requires casting through unknown first
        const doc = bodyParent as unknown as GoogleAppsScript.Document.Document;
        doc.saveAndClose();
      }
    }
  } catch (error) {
    // If we can't save, that's okay - the document might already be closed
    // or managed elsewhere
    Logger.log("Note: Could not auto-save document in formatParagraph");
  }
}
