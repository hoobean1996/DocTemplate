/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

type CreateNamedRangeOptions = {
  name: string;
  scope: "one" | "all";
};

type CreateNamedRangeResponse = {
  status: Number;
  message: "OK" | "Error";
  data: ListNamedRangesResponse,
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
  document.addNamedRange(options.name, rangeBuilder.build());
  return {
    status: 0,
    message: "OK",
    data: listNamedRanges(true),
  };
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
function listNamedRanges(withColor: boolean): ListNamedRangesResponse {
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
          text.setBackgroundColor(startOffset, endOffset,
            withColor ? color : null
          );
        } else {
          text.setBackgroundColor(
            withColor ? color : null
          );
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

function createProperty(name:string, value:string): CreatePropertyResponse {
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