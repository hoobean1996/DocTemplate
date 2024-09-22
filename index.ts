type CreateNamedRangeOptions = {
  name: string;
  kind: 'filler' | 'condition',
  scope: "one" | "all";
};

type CreateNamedRangeResponse = {
  status: Number;
  message: "OK" | "Error";
  data: ListNamedRangesResponse | null,
};

function createNamedRange(
  options: CreateNamedRangeOptions
): CreateNamedRangeResponse {
  const ui = DocumentApp.getUi();
  const document = DocumentApp.getActiveDocument();
  const selection = document.getSelection();
  const rangeBuilder = document.newRange();
  if (!selection) {
    ui.alert("Please select something to create Placeholder");
    return {
      status: -1,
      message: "Error",
      data: listNamedRanges(true),
    };
  }
  try {
    var elements = selection.getRangeElements();
    for (let i = 0; i < elements.length; i++) {
      const element = elements[i].getElement();
      // @ts-ignore
      // 判断当前元素是否能够转换为Text
      if (element.editAsText) {
        // @ts-ignore
        var text = element.editAsText();
        var startOffset = elements[i].isPartial()
          ? elements[i].getStartOffset()
          : 0;
        var endOffset = elements[i].isPartial()
          ? elements[i].getEndOffsetInclusive()
          : text.getText().length - 1;
        text.setBackgroundColor(startOffset, endOffset, "#00FF00"); // 设置为绿色
        rangeBuilder.addElement(text, startOffset, endOffset);
      }
    }
    if (options.kind === 'filler') {
      document.addNamedRange('[F]' + options.name, rangeBuilder.build());
    } else {
      document.addNamedRange('[C]' + options.name, rangeBuilder.build());
    }
    return {
      status: 0,
      message: "OK",
      data: listNamedRanges(true),
    };
  } catch (err) {
    return {
      status: 0,
      // @ts-ignore
      message: 'Error: ' + err.message,
      data: null,
    };
  }
}

/**
 * Represents the response structure for listing named ranges.
 * @typedef {Object} ListNamedRangesResponse
 * @property {Array<{name: string, start: number, end: number}>} placeholders - Array of placeholder objects.
 */
type ListNamedRangesResponse = {
  placeholders: {
    id: string;
    name: string;
    color: string | null;
  }[];
};

/**
 * Lists all named ranges in the current document.
 * @returns {ListNamedRangesResponse} An object containing an array of placeholders.
 */
function listNamedRanges(withColor: boolean, updateColor: boolean = true): ListNamedRangesResponse {
  const doc = DocumentApp.getActiveDocument();
  const namedRanges = doc.getNamedRanges();
  var colors = ['#FFD700', '#FF69B4', '#00CED1', '#32CD32', '#FFA500', '#9370DB', '#20B2AA'];
  const placeholders = namedRanges.map((range, i) => {
    var color = colors[i % colors.length]; // 循环使用颜色
    var rangeElements = range.getRange().getRangeElements();
    for (var j = 0; j < rangeElements.length; j++) {
      var element = rangeElements[j].getElement();
      // @ts-ignore
      if (element.editAsText) {
        // @ts-ignore
        var text = element.editAsText();
        var startOffset = rangeElements[j].getStartOffset();
        var endOffset = rangeElements[j].getEndOffsetInclusive();

        if (startOffset != null && endOffset != null) {
          if (updateColor) {
            text.setBackgroundColor(startOffset, endOffset,
              withColor ? color : null
            );
          }
        } else {
          if (updateColor) {
            text.setBackgroundColor(
              withColor ? color : null
            );
          }
        }
      }
    }

    return {
      id: range.getId(),
      name: range.getName(),
      color: withColor ? color : null,
    };
  });

  return {
    placeholders: placeholders,
  };
}

function removeNamedRange(id: string): ListNamedRangesResponse {
  const doc = DocumentApp.getActiveDocument();
  const namedRange = doc.getNamedRangeById(id);
  if (namedRange != null) {
    var rangeElements = namedRange.getRange().getRangeElements();
    for (var j = 0; j < rangeElements.length; j++) {
      var element = rangeElements[j].getElement();
      // @ts-ignore
      if (element.editAsText) {
        // @ts-ignore
        var text = element.editAsText();
        var startOffset = rangeElements[j].getStartOffset();
        var endOffset = rangeElements[j].getEndOffsetInclusive();

        if (startOffset != null && endOffset != null) {
          text.setBackgroundColor(startOffset, endOffset,
            null
          );
        } else {
          text.setBackgroundColor(
            null
          );
        }
      }
    }
    namedRange.remove();
  }
  return listNamedRanges(true);
}

function removeAllNamedRanges(): ListNamedRangesResponse {
  var doc = DocumentApp.getActiveDocument();
  var namedRanges = doc.getNamedRanges();
  namedRanges.forEach((namedRange) => {
    removeNamedRange(namedRange.getId());
  })
  return listNamedRanges(true);
}

function locateNamedRange(id: string): void {
  const doc = DocumentApp.getActiveDocument();
  const namedRange = doc.getNamedRangeById(id);
  if (!namedRange) {
    return;
  }
  const rangeElement = namedRange.getRange().getRangeElements()[0];
  const element = rangeElement.getElement();
  const startOffset = rangeElement.getStartOffset();
  const position = doc.newPosition(element, startOffset);
  doc.setCursor(position);
}


function getSelectedTextByRangeElements(elements: GoogleAppsScript.Document.RangeElement[]): Array<string> {
  const doc = DocumentApp.getActiveDocument();
  const text: Array<string> = [];
  for (let i = 0; i < elements.length; ++i) {
    if (elements[i].isPartial()) {
      const element = elements[i].getElement().asText();
      const startIndex = elements[i].getStartOffset();
      const endIndex = elements[i].getEndOffsetInclusive();

      // @ts-ignore
      text.push(element.getText().substring(startIndex, endIndex + 1));
    } else {
      const element = elements[i].getElement();
      // @ts-ignore
      if (element.editAsText) {
        const elementText = element.asText().getText();
        // This check is necessary to exclude images, which return a blank
        // text element.
        if (elementText) {
          // @ts-ignore
          text.push(elementText);
        }
      }
    }
  }
  return text;
}

function getSelectedText(): Array<string> {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  const text: Array<string> = [];
  if (selection) {
    const elements = selection.getRangeElements();
    return getSelectedTextByRangeElements(elements);
  }
  return text;
}

function getNamedRangeText(id: string): Array<string> {
  const doc = DocumentApp.getActiveDocument();
  const text: Array<string> = [];
  const nameRange = doc.getNamedRangeById(id);
  return getSelectedTextByRangeElements(nameRange.getRange().getRangeElements());
}

type Property = {
  name: string,
  value: string
}

type ListPropertyResponse = {
  properties: Property[],
}

function listProperties(): ListPropertyResponse {
  const ps = PropertiesService.getDocumentProperties();
  const properties = ps.getProperties();
  return {
    properties: Object.entries(properties).map(([name, value]) => ({
      name,
      value
    }))
  };
}

type CreatePropertyResponse = {
  status: Number;
  message: "OK" | "Error";
  data: ListPropertyResponse,
}

function createProperty(name: string, value: string): CreatePropertyResponse {
  const ps = PropertiesService.getDocumentProperties();
  ps.setProperty(name, value);
  return {
    status: 0,
    message: 'OK',
    data: listProperties(),
  };
}

function removeProperty(name: string): ListPropertyResponse {
  const ps = PropertiesService.getDocumentProperties();
  ps.deleteProperty(name);
  return listProperties();
}

type GenerateDocumentRequest = {
  mockedPlaceholders: {
    id: string,
    value: string | boolean,
  }[]
}



function getTopLevelContainer(element) {
  while (
    element.getParent() &&
    element.getParent().getType() !== DocumentApp.ElementType.BODY_SECTION
  ) {
    element = element.getParent();
  }
  return element;
}

/// Remove given document's namedRange's color
function removeNamedRangesColorByDocumentID(id: string): void {
  const doc = DocumentApp.openById(id);
  const namedRanges = doc.getNamedRanges();
  namedRanges.forEach((range, i) => {
    var rangeElements = range.getRange().getRangeElements();
    for (var j = 0; j < rangeElements.length; j++) {
      var element = rangeElements[j].getElement();
      // @ts-ignore
      if (element.editAsText) {
        // @ts-ignore
        var text = element.editAsText();
        var startOffset = rangeElements[j].getStartOffset();
        var endOffset = rangeElements[j].getEndOffsetInclusive();

        if (startOffset != null && endOffset != null) {
          text.setBackgroundColor(startOffset, endOffset, null);
        } else {
          text.setBackgroundColor(null);
        }
      }
    }
    range.remove();
  });
}

function generateDocument(request: GenerateDocumentRequest): string {
  listNamedRanges(false);
  const mockedPlaceholders = request.mockedPlaceholders;
  const sourceDoc = DocumentApp.getActiveDocument();
  const sourceID = sourceDoc.getId();
  const destDocFile = DriveApp.getFileById(sourceID).makeCopy("Test Case of " + sourceDoc.getName());
  const destDocID = destDocFile.getId();
  const destDoc = DocumentApp.openById(destDocID);
  const operations: { rangeElementType: string, parentElementType: string, parentElementText: string }[] = [];
  try {
    const requests = mockedPlaceholders.map((placeholder) => {
      switch (typeof placeholder.value) {
        case 'string': {
          const ps = PropertiesService.getDocumentProperties();
          const namedRange = sourceDoc.getNamedRangeById(placeholder.id);
          const name = namedRange.getName();
          let value = (placeholder.value.trim().length > 0) ? placeholder.value : ps.getProperty(name.substring(3,));
          if (value == null) {
            value = 'Unset';
          }
          return {
            replaceNamedRangeContent: {
              namedRangeId: 'kix.' + placeholder.id,
              text: value,
            }
          } as GoogleAppsScript.Docs.Schema.Request;
        }
        case 'boolean': {
          if (placeholder.value === false) {
            return null;
          }
          const namedRange = destDoc.getNamedRangeById(placeholder.id);
          const range = namedRange.getRange();
          const rangeElements = range.getRangeElements();
          for (let i = rangeElements.length - 1; i >= 0; i--) {
            const rangeElement = rangeElements[i];
            const element = rangeElement.getElement();
            const parent = element.getParent();
            const topLevelParent = getTopLevelContainer(element);
            parent.removeFromParent();
            if (topLevelParent.getType() === DocumentApp.ElementType.TABLE) {
              try {
                topLevelParent.removeFromParent();
              } catch (err) { }
            }
          }
          return null;
        }
      }
    })
    if (requests.filter(r => r != null).length > 0) {
      Docs.Documents?.batchUpdate({ 'requests': requests.filter(r => r != null) }, destDocID);
    }
    removeEmptyParagraphs(destDocID);
    return JSON.stringify({ 'requests': requests, 'operations': operations });
  } catch (e) {
    return e.message;
  }
}


function inspectNamedRange(namedRange: GoogleAppsScript.Document.NamedRange): string {
  var result = {
    rangeElementsCount: 0,
  };

  const rangeElements = namedRange.getRange().getRangeElements();
  result.rangeElementsCount = rangeElements.length;

  const rangeElementInfos = [];
  try {
    for (let i = 0; i < rangeElements.length; i++) {
      var rangeElementInfo = {
        startOffset: -1,
        endOffset: -1,
        elementType: DocumentApp.ElementType.UNSUPPORTED,
        topLevelElementType: DocumentApp.ElementType.UNSUPPORTED,
      };
      const rangeElement = rangeElements[i];
      rangeElementInfo.startOffset = rangeElement.getStartOffset();
      rangeElementInfo.endOffset = rangeElement.getEndOffsetInclusive();
      const originalElement = rangeElement.getElement();
      rangeElementInfo.elementType = originalElement.getType();
      const parent = originalElement.getParent();
      const topLevelParent = getTopLevelContainer(originalElement);
      rangeElementInfo.parent = parent.getType();
      rangeElementInfo.topLevelElementType = topLevelParent.getType();

      if (topLevelParent.getType() === DocumentApp.ElementType.TABLE) {
        const table = topLevelParent.asTable() as GoogleAppsScript.Document.Table;
        rangeElementInfo.isTable = true;
        rangeElementInfo.text = table.getText();
      }

      if (topLevelParent.getType() === DocumentApp.ElementType.LIST_ITEM) {
        const listItem = topLevelParent.asListItem() as GoogleAppsScript.Document.ListItem;
        rangeElementInfo.isList = true;
        rangeElementInfo.listId = listItem.getListId();
      }

      // @ts-ignore
      rangeElementInfos.push(rangeElementInfo);
    }
  } catch (err) { }
  result.rangeElementInfos = rangeElementInfos;
  return JSON.stringify(result);
}

function removeEmptyParagraphs(id: string) {
  var doc = DocumentApp.openById(id);
  var body = doc.getBody();
  var numChildren = body.getNumChildren();

  // 从后向前遍历，以避免删除元素时影响索引
  for (var i = numChildren - 1; i >= 0; i--) {
    var child = body.getChild(i);

    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      var paragraph = child.asParagraph();

      try {
        if (isEmptyParagraph(paragraph)) {
          body.removeChild(paragraph);
        }
      } catch (e) { }
    }
  }
}

function isEmptyParagraph(paragraph) {
  // 检查段落的文本内容
  var text = paragraph.getText().trim();
  console.log('text: ' + text);
  // 检查段落是否包含任何内联图片
  var numChildren = paragraph.getNumChildren();
  var hasInlineImages = false;

  for (var i = 0; i < numChildren; i++) {
    if (paragraph.getChild(i).getType() === DocumentApp.ElementType.INLINE_IMAGE) {
      hasInlineImages = true;
      break;
    }
  }

  // 如果文本为空且没有内联图片，则认为段落为空
  return text === "" && !hasInlineImages;
}