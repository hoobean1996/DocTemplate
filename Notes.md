```
// 删除Condition
const namedRange = destDoc.getNamedRangeById(placeholder.id);
          const range = namedRange.getRange();
          const elements = range.getRangeElements();

          // 从后向前遍历元素，以避免删除影响索引
          for (let i = elements.length - 1; i >= 0; i--) {
            const element = elements[i];
            const isPartial = (element.isPartial());
            const rangeElement = element.getElement();
            if (rangeElement.editAsText) {
              // 文本元素
              const text = rangeElement.editAsText();
              if (isPartial) {
                // 部分文本
                const startIndex = element.getStartOffset();
                const endIndex = element.getEndOffsetInclusive();
                text.deleteText(startIndex, endIndex);
              } else {
                // 整个文本元素
                rangeElement.removeFromParent();
              }
            } else {
              // 非文本元素（如图片、表格等）
              rangeElement.removeFromParent();
            }
          }
        ```