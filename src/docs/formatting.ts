export function formatParagraph(
  paragraph: GoogleAppsScript.Document.Paragraph | GoogleAppsScript.Document.ListItem,
  fontFamily: string = "Roboto",
): void {
  paragraph.setAttributes({
    [DocumentApp.Attribute.FONT_FAMILY]: fontFamily,
  });
  paragraph.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
}
