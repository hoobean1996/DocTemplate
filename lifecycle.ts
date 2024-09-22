function showSidebar() {
  const ui =
    HtmlService.createHtmlOutputFromFile("sidebar").setTitle("DocTemplate");
  DocumentApp.getUi().showSidebar(ui);
}

function showCreatePropertiesDialog() {
  const template = HtmlService.createTemplateFromFile("createPropertiesDialog");
  const html = template.evaluate().setWidth(800).setHeight(600);
  DocumentApp.getUi().showModalDialog(html, "Create Property");
}

function openGenDocumentDialog() {
  const template = HtmlService.createTemplateFromFile("genDocumentDialog");
  const listNamedRangesResponse = listNamedRanges(false, false);
  template.placeholders = listNamedRangesResponse.placeholders;
  const html = template.evaluate().setWidth(800).setHeight(600);
  DocumentApp.getUi().showModalDialog(html, "Generate Document");
}

function showCreatePlaceholderDialog() {
  const template = HtmlService.createTemplateFromFile("createPlaceholderDialog");
  const selectedText = getSelectedText();
  if (selectedText.length === 0) {
    template.selected = false;
    template.selectedText = 'You did not select anything, please close dialog and select something first.'
  } else {
    template.selected = true;
    template.selectedText = selectedText;
  }
  const html = template.evaluate()
    .setWidth(800)
    .setHeight(600);
  DocumentApp.getUi().showModalDialog(html, "Create Placeholder");
}

function showPlaceholderDialog(id: string) {
  const doc = DocumentApp.getActiveDocument();
  const namedRange = doc.getNamedRangeById(id);
  const template = HtmlService.createTemplateFromFile("editPlaceholderDialog");
  template.id = id;
  template.name = namedRange.getName();
  template.inspect = inspectNamedRange(namedRange);
  const html = template.evaluate()
    .setWidth(800)
    .setHeight(600);
  DocumentApp.getUi().showModalDialog(html, "Edit Placeholder");
}

function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  DocumentApp.getUi()
    .createAddonMenu()
    .addItem("Start", "showSidebar")
    .addToUi();
}

