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
  let processedText = text;
  
  if (useSmartIgnore) {
    // 1. Process Dates (e.g., "5 May 2024" -> "5 พฤษภาคม 2567")
    const dateRegex = /\b(\d{1,2})?\s?(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s?(\d{1,2})?,?\s(20\d{2})\b/gi;
    
    processedText = processedText.replace(dateRegex, (match, d1, month, d2, year) => {
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
    range.load("text");
    await context.sync();
    
    const originalFullText = range.text;
    if (!originalFullText) return;

    const dateConvertedText = convertText(originalFullText, useSmartIgnore);
    
    if (originalFullText !== dateConvertedText) {
        range.insertText(dateConvertedText, "Replace");
        return;
    }

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
  } catch (e) {}
}

/**
 * Processes a body object for text, shapes, and fields.
 */
async function processBodyExhaustive(body: Word.Body, useSmartIgnore: boolean, context: Word.RequestContext) {
  if (!body) return;

  // 1. Main text content
  await processRange(body.getRange(), useSmartIgnore, context);

  // 2. Shapes (Textboxes)
  try {
    const shapes = body.shapes;
    shapes.load("items/body");
    await context.sync();

    for (let i = 0; i < shapes.items.length; i++) {
      const shape = shapes.items[i];
      if (shape.body) {
        await processBodyExhaustive(shape.body, useSmartIgnore, context);
      }
    }
  } catch (e) {}

  // 3. Fields (Captions, etc. - Note: page numbers handled at section level)
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
 * Converts numerals in the entire document with dynamic styling.
 */
export async function convertDocument(useSmartIgnore: boolean, includeHF: boolean, dynamicLists: boolean, dynamicPageNumbers: boolean) {
  await Word.run(async (context: Word.RequestContext) => {
    // 1. Process Main Body
    await processBodyExhaustive(context.document.body, useSmartIgnore, context);

    // 2. Handle Dynamic Lists (๑.๑) - Non-destructive style change
    if (dynamicLists) {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      // Load list template and levels via any casting to bypass build errors
      paragraphs.load("items/list");
      await context.sync();

      for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        if (para.list) {
          try {
            const listAny = para.list as any;
            // Properties like listTemplate might be in newer API sets not yet in local types
            const levels = listAny.listTemplate.listLevels;
            levels.load("items");
            await context.sync();
            
            for (let j = 0; j < levels.items.length; j++) {
              // Change style to ThaiArabic (๑, ๒, ๓)
              levels.items[j].numberStyle = "ThaiArabic";
            }
          } catch (e) {
            // Fallback if list styling fails
          }
        }
      }
    }

    // 3. Handle Page Numbers & Sections (Dynamic counters)
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    for (let i = 0; i < sections.items.length; i++) {
      const section = sections.items[i];
      
      // Fix Page Numbers: Use built-in dynamic numbering style
      if (dynamicPageNumbers) {
        try {
            (section as any).pageNumbering.numberStyle = "ThaiArabic";
        } catch (e) {}
      }

      // Process Headers/Footers
      if (includeHF) {
        const hfTypes: Word.HeaderFooterType[] = [
            Word.HeaderFooterType.primary,
            Word.HeaderFooterType.firstPage,
            Word.HeaderFooterType.evenPages
        ];
        
        for (const type of hfTypes) {
          try {
            const hf = section.getHeader(type);
            await processBodyExhaustive(hf, useSmartIgnore, context);
            const ff = section.getFooter(type);
            await processBodyExhaustive(ff, useSmartIgnore, context);
          } catch (e) {}
        }
      }
    }
    await context.sync();
  });
}
