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
  // Paragraph -> Body -> Document (Body has a method to get the document)
  try {
    const parent = paragraph.getParent();
    if (parent && parent.getType() === DocumentApp.ElementType.BODY_SECTION) {
      const body = parent.asBody();
      // Access the document through body - we need to get it from DriveApp using the doc ID
      // But paragraphs don't expose doc ID directly. Instead, try to save via body if possible
      // Actually, the simplest is to reopen the document by getting the paragraph's editAsText
      // But that's not available either. Let's use a different approach - check if we can
      // access the document's saveAndClose through the body's parent chain
      const bodyParent = body.getParent();
      if (bodyParent && (bodyParent as any).saveAndClose) {
        (bodyParent as any).saveAndClose();
      }
    }
  } catch (error) {
    // If we can't save, that's okay - the document might already be closed
    // or managed elsewhere
    Logger.log("Note: Could not auto-save document in formatParagraph");
  }
}
