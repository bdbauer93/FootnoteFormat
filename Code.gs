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

  let tabsToProcess = [];

  if (config.allTabs) {
    // Collect every tab in the document (including nested ones)
    tabsToProcess = getAllTabs(doc);
  } else {
    // Just the currently active tab
    tabsToProcess = [doc.getActiveTab()];
  }

  const size = parseInt(config.size);
  let footnotesProcessed = 0;

  tabsToProcess.forEach(tab => {
    // Tabs must be cast to DocumentTab to access body/footnotes
    if (tab.getType() === DocumentApp.TabType.DOCUMENT_TAB) {
      const docTab = tab.asDocumentTab();
      const footnotes = docTab.getFootnotes();
      
      footnotes.forEach(footnote => {
        const contents = footnote.getFootnoteContents();
        contents.setFontFamily(config.font).setFontSize(size);
        footnotesProcessed++;
      });
    }
  });

  const target = config.allTabs ? "all tabs" : "the current tab";
  const notes = footnotesProcessed > 1 ? "footnotes" : "footnote";
  DocumentApp.getUi().alert('Success', `Formatted ${footnotesProcessed} ${notes} across ${target}.`, DocumentApp.getUi().ButtonSet.OK);
}

/**
 * Returns a flat list of all tabs in the document, in the order
 * they would appear in the UI (i.e. top-down ordering). Includes
 * all child tabs.
 */
function getAllTabs(doc) {
  const allTabs = [];
  // Iterate over all tabs and recursively add any child tabs to
  // generate a flat list of Tabs.
  for (const tab of doc.getTabs()) {
    addCurrentAndChildTabs(tab, allTabs);
  }
  return allTabs;
}

/**
 * Adds the provided tab to the list of all tabs, and recurses
 * through and adds all child tabs.
 */
function addCurrentAndChildTabs(tab, allTabs) {
  allTabs.push(tab);
  for (const childTab of tab.getChildTabs()) {
    addCurrentAndChildTabs(childTab, allTabs);
  }
}
