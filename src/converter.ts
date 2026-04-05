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
    return text.replace(smartRegex, (match) => {
      return match.split('').map(char => ARABIC_TO_THAI_MAP[char] || char).join('');
    });
  } else {
    return text.replace(/[0-9]/g, (match) => ARABIC_TO_THAI_MAP[match] || match);
  }
}

/**
 * Helper to check if a character is a Latin letter.
 */
function isLetter(char: string | undefined): boolean {
  if (!char) return false;
  return /[a-zA-Z]/.test(char);
}

/**
 * Core logic to process a range and replace numbers while preserving formatting.
 */
async function processRange(range: Word.Range, useSmartIgnore: boolean, context: Word.RequestContext) {
  // Search for all digit sequences
  const results = range.search("[0-9]{1,}", { matchWildcards: true });
  results.load("items");
  await context.sync();

  for (let i = 0; i < results.items.length; i++) {
    const numRange = results.items[i];
    
    if (useSmartIgnore) {
      // To implement Smart Ignore with formatting preservation, we check the characters 
      // immediately before and after the matched range.
      
      const rangeBefore = numRange.getRange("Before").expandTo(numRange.getRange("Before").move("Character", -1));
      const rangeAfter = numRange.getRange("After").expandTo(numRange.getRange("After").move("Character", 1));
      
      rangeBefore.load("text");
      rangeAfter.load("text");
      await context.sync();

      const charBefore = rangeBefore.text;
      const charAfter = rangeAfter.text;

      // If either side is a letter, skip this conversion
      if (isLetter(charBefore) || isLetter(charAfter)) {
        continue;
      }
    }

    // Load the current text of the number range
    numRange.load("text");
    await context.sync();

    const originalText = numRange.text;
    const thaiText = originalText.split('').map(char => ARABIC_TO_THAI_MAP[char] || char).join('');
    
    if (originalText !== thaiText) {
      numRange.insertText(thaiText, "Replace");
    }
  }
}

/**
 * Converts numerals in the current selection.
 */
export async function convertSelection(useSmartIgnore: boolean) {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    await processRange(selection, useSmartIgnore, context);
    await context.sync();
  });
}

/**
 * Converts numerals in the entire document body.
 */
export async function convertDocument(useSmartIgnore: boolean) {
  await Word.run(async (context) => {
    const body = context.document.body;
    await processRange(body, useSmartIgnore, context);
    await context.sync();
  });
}
