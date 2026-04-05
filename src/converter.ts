/* global Word */

export const ARABIC_TO_THAI_MAP: { [key: string]: string } = {
  '0': '๐', '1': '๑', '2': '๒', '3': '๓', '4': '๔',
  '5': '๕', '6': '๖', '7': '๗', '8': '๘', '9': '๙',
};

const MONTHS_MAP: { [key: string]: string } = {
  'January': 'มกราคม', 'February': 'กุมภาพันธ์', 'March': 'มีนาคม',
  'April': 'เมษายน', 'May': 'พฤษภาคม', 'June': 'มิถุนายน',
  'July': 'กรกฎาคม', 'August': 'สิงหาคม', 'September': 'กันยายน',
  'October': 'ตุลาคม', 'November': 'พฤศจิกายน', 'December': 'ธันวาคม',
  'Jan': 'ม.ค.', 'Feb': 'ก.พ.', 'Mar': 'มี.ค.', 'Apr': 'เม.ย.', 'Jun': 'มิ.ย.',
  'Jul': 'ก.ค.', 'Aug': 'ส.ค.', 'Sep': 'ก.ย.', 'Oct': 'ต.ค.', 'Nov': 'พ.ย.', 'Dec': 'ธ.ค.'
};

/**
 * Converts Arabic numerals to Thai numerals in a string.
 */
export function convertText(text: string, useSmartIgnore: boolean): string {
  // Bonus: Auto-convert AD Year to BE Year if it looks like a year (2000-2100)
  // and is not preceded by "ค.ศ."
  let processedText = text;
  
  if (useSmartIgnore) {
    // 1. Process Dates (e.g., "5 May 2024" -> "5 พฤษภาคม 2567")
    // Regex for Month Day, Year or Day Month Year
    const dateRegex = /\b(\d{1,2})?\s?(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s?(\d{1,2})?,?\s(20\d{2})\b/gi;
    
    processedText = processedText.replace(dateRegex, (match, d1, month, d2, year) => {
      // Check if "ค.ศ." is before the match (simple check)
      const index = processedText.indexOf(match);
      const before = processedText.substring(Math.max(0, index - 10), index);
      if (before.includes("ค.ศ.")) return match;

      const thaiMonth = MONTHS_MAP[month] || month;
      const beYear = parseInt(year) + 543;
      const day = d1 || d2 || "";
      
      return `${day} ${thaiMonth} ${beYear}`.trim();
    });

    // 2. Smart Numeral Conversion
    const smartRegex = /(?<![a-zA-Z0-9])[0-9]+(?![a-zA-Z0-9])/g;
    return processedText.replace(smartRegex, (match: string) => {
      return match.split('').map((char: string) => ARABIC_TO_THAI_MAP[char] || char).join('');
    });
  } else {
    return processedText.replace(/[0-9]/g, (match: string) => ARABIC_TO_THAI_MAP[match] || match);
  }
}

/**
 * Core logic to process a range and replace numbers while preserving formatting.
 */
async function processRange(range: Word.Range, useSmartIgnore: boolean, context: Word.RequestContext) {
  try {
    // First, do a pass for the Date Formatter logic on the whole range text
    // This is a bit destructive to formatting if we replace the whole range, 
    // so we only do it if we find a date match.
    range.load("text");
    await context.sync();
    
    const originalFullText = range.text;
    const dateConvertedText = convertText(originalFullText, useSmartIgnore);
    
    if (originalFullText !== dateConvertedText) {
        // If dates were found and converted, we have to replace the text.
        // To preserve formatting, we'd need a more complex range-by-range search.
        // For now, we update the text.
        range.insertText(dateConvertedText, "Replace");
        return;
    }

    // Standard numeral-by-numeral replacement (Best for formatting)
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
        if (/^[0-9]+$/.test(text)) {
          const thaiText = text.split('').map((char: string) => ARABIC_TO_THAI_MAP[char] || char).join('');
          blockRange.insertText(thaiText, "Replace");
        }
      } else {
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
async function processBodyExhaustive(body: Word.Body, useSmartIgnore: boolean, context: Word.RequestContext, flattenLists: boolean) {
  if (!body) return;

  // 1. Process List Flattening
  if (flattenLists) {
    const paragraphs = body.paragraphs;
    paragraphs.load("items/isListItem,items/listItem/retrieveLabel");
    await context.sync();

    for (let i = 0; i < paragraphs.items.length; i++) {
      const para = paragraphs.items[i];
      if (para.isListItem) {
        const label = para.listItem.retrieveLabel();
        await context.sync();
        const thaiLabel = convertText(label.value, false);
        // Delete auto-numbering and insert Thai text
        para.listItem.deleteNumbering();
        para.insertText(thaiLabel + " ", "Start");
      }
    }
  }

  // 2. Process the main text content
  await processRange(body.getRange(), useSmartIgnore, context);

  // 3. Process all shapes (like Textboxes)
  try {
    const shapes = body.shapes;
    shapes.load("items/body");
    await context.sync();

    for (let i = 0; i < shapes.items.length; i++) {
      const shape = shapes.items[i];
      if (shape.body) {
        await processBodyExhaustive(shape.body, useSmartIgnore, context, flattenLists);
      }
    }
  } catch (e) {}

  // 4. Process all fields (Page Numbers, etc.)
  try {
    const fields = body.fields;
    fields.load("items/result");
    await context.sync();

    for (let i = 0; i < fields.items.length; i++) {
      const field = fields.items[i];
      if (field.result) {
        await processRange(field.result, useSmartIgnore, context);
      }
    }
  } catch (e) {}
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
export async function convertDocument(useSmartIgnore: boolean, includeHF: boolean, flattenLists: boolean) {
  await Word.run(async (context: Word.RequestContext) => {
    // 1. Process Body
    await processBodyExhaustive(context.document.body, useSmartIgnore, context, flattenLists);

    // 2. Process Headers/Footers
    if (includeHF) {
      const sections = context.document.sections;
      sections.load("items");
      await context.sync();

      for (let i = 0; i < sections.items.length; i++) {
        const section = sections.items[i];
        const hfTypes: Word.HeaderFooterType[] = [
            Word.HeaderFooterType.primary,
            Word.HeaderFooterType.firstPage,
            Word.HeaderFooterType.evenPages
        ];
        
        for (const type of hfTypes) {
          try {
            await processBodyExhaustive(section.getHeader(type), useSmartIgnore, context, flattenLists);
            await processBodyExhaustive(section.getFooter(type), useSmartIgnore, context, flattenLists);
          } catch (e) {}
        }
      }
    }
    await context.sync();
  });
}
