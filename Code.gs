/**
 * @OnlyCurrentDoc
 */

function onOpen() {
  DocumentApp.getUi()
    .createMenu('破音字標注 Poyinzi')
    .addItem('開啟破音字校正', 'showNativeSidebar')
    .addItem('一鍵套用注音字型 (反白文字)', 'applyMagicFontToSelection')
    .addToUi();
}

function showNativeSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('破音字校正')
      .setWidth(300);
  DocumentApp.getUi().showSidebar(html);
}

// 自動幫反白的文字套用字型 (不強制改變字體大小)
function applyMagicFontToSelection() {
  var doc = DocumentApp.getActiveDocument();
  var selection = doc.getSelection();
  
  if (!selection) throw new Error('請先「反白」你想變成注音體的文字！');

  var elements = selection.getRangeElements();
  elements.forEach(function(el) {
    var textElement = el.getElement();
    
    if (textElement.editAsText) {
      var text = textElement.asText();
      var start = el.getStartOffset();
      var end = el.getEndOffsetInclusive();
      
      if (start === -1) start = 0;
      if (end === -1) end = text.getText().length - 1;

      if (end >= start) {
        // 1. 套用字嗨注音標楷
        text.setFontFamily(start, end, "Bpmf Zihi Kai Std");
      }
    }
  });
  return "✨ 字型套用與排版完成！";
}

function getSelectedChar() {
  var doc = DocumentApp.getActiveDocument();
  var selection = doc.getSelection();
  if (!selection) return { error: '請先在文件中「反白」一個字！' };

  var elements = selection.getRangeElements();
  var textElement = elements[0].getElement();
  if (textElement.getType() !== DocumentApp.ElementType.TEXT) return { error: '選取的不是純文字喔！' };
  
  var startOffset = elements[0].getStartOffset();
  var endOffset = elements[0].getEndOffsetInclusive();
  if (startOffset === -1 || endOffset === -1) return { error: '請精準反白文字。' };

  var selectedText = textElement.asText().getText().substring(startOffset, endOffset + 1);
  var match = selectedText.match(/[\u4e00-\u9fa5]/);
  if (!match) return { error: '沒有找到中文字喔！' };
  
  return { success: true, char: match[0] };
}

function changePhonetic(payload) {
  var doc = DocumentApp.getActiveDocument();
  var selection = doc.getSelection();

  if (!selection) {
    throw new Error("請先反白選取要修改的字！");
  }

  var baseChar = payload.baseChar;
  var level = payload.variantLevel;
  
  // 組合出新的字元 (包含變體選擇符)
  var newText = baseChar;
  if (level > 0) {
    newText += String.fromCodePoint(0xE01E0 + level);
  }

  var elements = selection.getRangeElements();
  var modified = false;

  // 使用反向迴圈 (從後往前改)，這樣插入/刪除文字時才不會導致後面的索引位置 (Offset) 跑掉
  for (var i = elements.length - 1; i >= 0; i--) {
    var rangeElement = elements[i];
    var element = rangeElement.getElement();

    // 確保選取到的是文字元素
    if (element.getType() === DocumentApp.ElementType.TEXT) {
      var textElement = element.asText();
      
      // 取得選取範圍的起點與終點
      var startOffset = rangeElement.isPartial() ? rangeElement.getStartOffset() : 0;
      var endOffset = rangeElement.isPartial() ? rangeElement.getEndOffsetInclusive() : textElement.getText().length - 1;

      // 1. 取得原本第一個字的樣式屬性 (包含大小、粗體、文字顏色、背景色等)
      var originalAttributes = textElement.getAttributes(startOffset);
      
      // 2. 覆寫字體屬性：確保修改後一定是「字嗨注音標楷體」，否則注音出不來
      originalAttributes[DocumentApp.Attribute.FONT_FAMILY] = 'Bpmf Zihi Kai Std';

      // 3. 刪除原本的字
      textElement.deleteText(startOffset, endOffset);
      
      // 4. 在原位置插入新的字 (包含變體碼)
      textElement.insertText(startOffset, newText);
      
      // 5. 將原本的樣式 (加上注音字型) 完整套用回新插入的字上
      textElement.setAttributes(startOffset, startOffset + newText.length - 1, originalAttributes);
      
      modified = true;
    }
  }

  if (modified) {
    return "已套用並保留原本格式！";
  } else {
    throw new Error("找不到可修改的文字。");
  }
}