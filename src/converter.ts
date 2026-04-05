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
 * Used for testing and internal logic.
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
    // Search for all contiguous alphanumeric blocks
    const results = range.search("[a-zA-Z0-9]{1,}", { matchWildcards: true });
    results.load("items");
    await context.sync();

    for (let i = 0; i < results.items.length; i++) {
      const blockRange = results.items[i];
      blockRange.load("text");
      await context.sync();

      const text = blockRange.text;
      // If the block is purely numeric, convert it
      if (/^[0-9]+$/.test(text)) {
        const thaiText = text.split('').map((char: string) => ARABIC_TO_THAI_MAP[char] || char).join('');
        blockRange.insertText(thaiText, "Replace");
      }
      // If it's a mix (e.g., spin9) or purely letters, we ignore it.
    }
  } else {
    // Standard replacement: Find all digit sequences regardless of surrounding text
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
 * Converts numerals in the entire document body.
 */
export async function convertDocument(useSmartIgnore: boolean) {
  await Word.run(async (context: Word.RequestContext) => {
    const body = context.document.body;
    await processRange(body.getRange(), useSmartIgnore, context);
    await context.sync();
  });
}
