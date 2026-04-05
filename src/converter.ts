/* global Word */

export const ARABIC_TO_THAI_MAP: { [key: string]: string } = {
  '0': '๐', '1': '๑', '2': '๒', '3': '๓', '4': '๔',
  '5': '๕', '6': '๖', '7': '๗', '8': '๘', '9': '๙',
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
  try {
    const searchPattern = useSmartIgnore ? "[a-zA-Z0-9]{1,}" : "[0-9]{1,}";
    const results = range.search(searchPattern, { matchWildcards: true });
    results.load("items");
    await context.sync();

    for (let i = 0; i < results.items.length; i++) {
      const blockRange = results.items[i];
      blockRange.load("text");
      await context.sync();

      const text = blockRange.text;
      if (useSmartIgnore) {
        // If purely numeric, convert. Otherwise ignore (it's mixed like spin9).
        if (/^[0-9]+$/.test(text)) {
          const thaiText = text.split('').map((char: string) => ARABIC_TO_THAI_MAP[char] || char).join('');
          blockRange.insertText(thaiText, "Replace");
        }
      } else {
        // Not using smart ignore, convert all digits found.
        const thaiText = text.split('').map((char: string) => ARABIC_TO_THAI_MAP[char] || char).join('');
        blockRange.insertText(thaiText, "Replace");
      }
    }
  } catch (e) {
    // Ignore inaccessible ranges
  }
}

/**
 * Thoroughly processes a Body object, including its text, nested shapes, and fields.
 */
async function processBodyExhaustive(body: Word.Body, useSmartIgnore: boolean, context: Word.RequestContext) {
  if (!body) return;

  // 1. Process the main text content of this body
  await processRange(body.getRange(), useSmartIgnore, context);

  // 2. Process all shapes (like Textboxes) within this body
  try {
    const shapes = body.shapes;
    shapes.load("items/body");
    await context.sync();

    for (let i = 0; i < shapes.items.length; i++) {
      const shape = shapes.items[i];
      // Note: shape.body is only accessible for text-supporting shapes
      if (shape.body) {
        // Recursively process the shape's body content
        await processBodyExhaustive(shape.body, useSmartIgnore, context);
      }
    }
  } catch (e) {
    // Some shapes don't support bodies
  }

  // 3. Process all fields (like dynamic Page Numbers or Captions) in this body
  try {
    const fields = body.fields;
    fields.load("items/result");
    await context.sync();

    for (let i = 0; i < fields.items.length; i++) {
      const field = fields.items[i];
      if (field.result) {
        // Process the dynamic result of the field
        await processRange(field.result, useSmartIgnore, context);
      }
    }
  } catch (e) {
    // Ignore field access errors
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
 * Converts numerals in the entire document (Body, Headers, Footers, Shapes, Fields).
 */
export async function convertDocument(useSmartIgnore: boolean) {
  await Word.run(async (context: Word.RequestContext) => {
    // 1. Process the main document body
    await processBodyExhaustive(context.document.body, useSmartIgnore, context);

    // 2. Process all headers and footers across all sections
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    for (let i = 0; i < sections.items.length; i++) {
      const section = sections.items[i];
      const hfTypes: Word.HeaderFooterType[] = ["Primary", "FirstPage", "EvenPages"];
      
      for (const type of hfTypes) {
        try {
          const header = section.getHeader(type);
          await processBodyExhaustive(header, useSmartIgnore, context);
          
          const footer = section.getFooter(type);
          await processBodyExhaustive(footer, useSmartIgnore, context);
        } catch (e) {
          // Some header/footer types might not exist in the document
        }
      }
    }
    await context.sync();
  });
}
