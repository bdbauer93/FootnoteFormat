/**
 * @OnlyCurrentDoc
 */
function onOpen() {
  DocumentApp.getUi()
    .createMenu('Document Tools')
    .addItem('Format Footnotes', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Footnote Formatter')
    .setWidth(300);
  DocumentApp.getUi().showSidebar(html);
}

function applyFootnoteFormatting(config) {
  const doc = DocumentApp.getActiveDocument();
  const footnotes = doc.getFootnotes();
  
  if (footnotes.length === 0) {
    throw new Error('No footnotes found.');
  }

  const size = parseInt(config.size);
  
  footnotes.forEach(footnote => {
    const contents = footnote.getFootnoteContents();
    contents.setFontFamily(config.font).setFontSize(size);
  });

  DocumentApp.getUi().alert('Success', 'Footnotes have been successfully formatted!', DocumentApp.getUi().ButtonSet.OK);
}
