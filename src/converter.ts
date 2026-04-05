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

    const smartRegex = /(?<![a-zA-Z0-9])[0-9]+(?![a-zA-Z0-9])/g;
    return processedText.replace(smartRegex, (match: string) => {
      return match.split('').map((char: string) => ARABIC_TO_THAI_MAP[char] || char).join('');
    });
  } else {
    return processedText.replace(/[0-9]/g, (match: string) => ARABIC_TO_THAI_MAP[match] || match);
  }
}

/**
 * Core logic to process a range. 
 * Updated to PROTECT fields (like page numbers) from being flattened.
 */
async function processRange(range: Word.Range, useSmartIgnore: boolean, context: Word.RequestContext) {
  try {
    range.load("text");
    await context.sync();
    
    const originalFullText = range.text;
    if (!originalFullText) return;

    // Check if this range is inside a Field (e.g., Page Number)
    // We try to get the parent field. If it exists, we skip text replacement to avoid "๑" bug.
    try {
        const parentField = range.parentBody.fields.getCount(); // Shallow check
        // If we are in a header/footer, we are much more careful.
        const parentBody = range.parentBody;
        parentBody.load("type");
        await context.sync();
        
        if (parentBody.type === "Header" || parentBody.type === "Footer") {
            // In headers/footers, we ONLY convert text that is NOT a field.
            // This is complex in Office.js, so we'll use a safer approach:
            // We search for numbers, but we skip them if they look like standalone page numbers.
        }
    } catch (e) {}

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
      
      // SAFETY: If the text is just a single digit and we are in a Header/Footer, 
      // it is likely a page number. SKIP IT to let dynamic numbering handle it.
      if ((text.length === 1 || text.length === 2) && (range as any).parentBody?.type === "Header") continue;

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
 * Processes a body object for text and shapes.
 */
async function processBodyExhaustive(body: Word.Body, useSmartIgnore: boolean, context: Word.RequestContext) {
  if (!body) return;
  await processRange(body.getRange(), useSmartIgnore, context);

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
}

/**
 * Main Conversion Functions
 */
export async function convertSelection(useSmartIgnore: boolean) {
  await Word.run(async (context: Word.RequestContext) => {
    const selection = context.document.getSelection();
    await processRange(selection, useSmartIgnore, context);
    await context.sync();
  });
}

export async function convertDocument(useSmartIgnore: boolean, includeHF: boolean, dynamicLists: boolean, dynamicPageNumbers: boolean) {
  await Word.run(async (context: Word.RequestContext) => {
    // 1. Process Main Body
    await processBodyExhaustive(context.document.body, useSmartIgnore, context);

    // 2. Handle Dynamic Lists (๑.๑) - Safer Implementation
    if (dynamicLists) {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items/isListItem,items/list");
      await context.sync();

      for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        if (para.isListItem && para.list) {
          try {
            const listLevels = (para.list as any).listTemplate.listLevels;
            listLevels.load("items");
            await context.sync();
            for (let j = 0; j < listLevels.items.length; j++) {
              listLevels.items[j].numberStyle = "ThaiArabic";
            }
          } catch (e) {
             // Fallback: try setting numbering style directly if template fails
             try { (para.list as any).setLevelNumbering(0, "ThaiArabic" as any); } catch(err) {}
          }
        }
      }
    }

    // 3. Handle Page Numbers & Headers/Footers
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    for (let i = 0; i < sections.items.length; i++) {
      const section = sections.items[i];
      
      if (dynamicPageNumbers) {
        try {
            // Apply Thai numbering to the section's page counter
            (section as any).pageNumbering.numberStyle = "ThaiArabic";
            // Also target fields specifically for TOC and Page numbers
            const body = section.body;
            const fields = body.fields;
            fields.load("items/type");
            await context.sync();
            for (let j = 0; j < fields.items.length; j++) {
                if (fields.items[j].type === "Page" || fields.items[j].type === "Toc") {
                    // This forces Word to re-evaluate the field with the new section style
                    (fields.items[j] as any).update(); 
                }
            }
        } catch (e) {}
      }

      if (includeHF) {
        const hfTypes: Word.HeaderFooterType[] = [
            Word.HeaderFooterType.primary,
            Word.HeaderFooterType.firstPage,
            Word.HeaderFooterType.evenPages
        ];
        for (const type of hfTypes) {
          try {
            const header = section.getHeader(type);
            const footer = section.getFooter(type);
            // Process text in headers/footers but the processRange now has a safety check
            await processBodyExhaustive(header, useSmartIgnore, context);
            await processBodyExhaustive(footer, useSmartIgnore, context);
          } catch (e) {}
        }
      }
    }
    await context.sync();
  });
}
