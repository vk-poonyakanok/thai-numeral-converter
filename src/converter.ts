/* global Word */

export const ARABIC_TO_THAI_MAP: { [key: string]: string } = {
  '0': '๐',
  '1': '๑',
  '2': '๒',
  '3': '๓',
  '4': '๔',
  '5': '๕',
  '6': '๖',
  '7': '๗',
  '8': '๘',
  '9': '๙',
};

/**
 * Converts Arabic numerals to Thai numerals in a string.
 */
export function convertText(text: string, useSmartIgnore: boolean): string {
  if (useSmartIgnore) {
    const smartRegex = /(?<![a-zA-Z0-9])[0-9]+(?![a-zA-Z0-9])/g;
    return text.replace(smartRegex, (match: string) => {
      return match.split('').map((char: string) => ARABIC_TO_THAI_MAP[char] || char).join('');
    });
  } else {
    return text.replace(/[0-9]/g, (match: string) => ARABIC_TO_THAI_MAP[match] || match);
  }
}

/**
 * Core logic to process a range and replace numbers while preserving formatting.
 */
async function processRange(range: Word.Range, useSmartIgnore: boolean, context: Word.RequestContext) {
  if (useSmartIgnore) {
    const results = range.search("[a-zA-Z0-9]{1,}", { matchWildcards: true });
    results.load("items");
    await context.sync();

    for (let i = 0; i < results.items.length; i++) {
      const blockRange = results.items[i];
      blockRange.load("text");
      await context.sync();

      const text = blockRange.text;
      if (/^[0-9]+$/.test(text)) {
        const thaiText = text.split('').map((char: string) => ARABIC_TO_THAI_MAP[char] || char).join('');
        blockRange.insertText(thaiText, "Replace");
      }
    }
  } else {
    const results = range.search("[0-9]{1,}", { matchWildcards: true });
    results.load("items");
    await context.sync();

    for (let i = 0; i < results.items.length; i++) {
      const numRange = results.items[i];
      numRange.load("text");
      await context.sync();

      const originalText = numRange.text;
      const thaiText = originalText.split('').map((char: string) => ARABIC_TO_THAI_MAP[char] || char).join('');
      
      if (originalText !== thaiText) {
        numRange.insertText(thaiText, "Replace");
      }
    }
  }
}

/**
 * Processes all relevant parts of a document: Body, Shapes, Headers, Footers.
 */
async function processAllDocumentParts(context: Word.RequestContext, useSmartIgnore: boolean) {
  // 1. Process Body
  const body = context.document.body;
  await processRange(body.getRange(), useSmartIgnore, context);

  // 2. Process Shapes (Textboxes)
  const shapes = body.shapes;
  // Load textFrame and the nested textRange
  shapes.load("items/textFrame/textRange");
  await context.sync();

  for (let i = 0; i < shapes.items.length; i++) {
    const shape = shapes.items[i];
    if (shape.textFrame && shape.textFrame.hasText) {
      const shapeRange = shape.textFrame.textRange;
      await processRange(shapeRange, useSmartIgnore, context);
    }
  }

  // 3. Process Headers and Footers
  const sections = context.document.sections;
  sections.load("items");
  await context.sync();

  for (let i = 0; i < sections.items.length; i++) {
    const section = sections.items[i];
    
    const parts = [
      section.getHeader("Primary"),
      section.getHeader("FirstPage"),
      section.getHeader("EvenPages"),
      section.getFooter("Primary"),
      section.getFooter("FirstPage"),
      section.getFooter("EvenPages")
    ];

    for (const part of parts) {
      await processRange(part.getRange(), useSmartIgnore, context);
    }
  }
}

/**
 * Converts numerals in the current selection.
 */
export async function convertSelection(useSmartIgnore: boolean) {
  await Word.run(async (context: Word.RequestContext) => {
    const selection = context.document.getSelection();
    await processRange(selection, useSmartIgnore, context);
    await context.sync();
  });
}

/**
 * Converts numerals in the entire document.
 */
export async function convertDocument(useSmartIgnore: boolean) {
  await Word.run(async (context: Word.RequestContext) => {
    await processAllDocumentParts(context, useSmartIgnore);
    await context.sync();
  });
}
