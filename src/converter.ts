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
 * Core logic to process a range SURGICALLY to preserve all formatting.
 */
async function processRange(range: Word.Range, useSmartIgnore: boolean, context: Word.RequestContext) {
  try {
    // 1. First, search for dates only if using smart ignore
    if (useSmartIgnore) {
        const datePattern = "(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Oct|Nov|Dec)";
        const dateResults = range.search(datePattern, { matchWildcards: false });
        dateResults.load("items");
        await context.sync();
        
        // Date conversion is complex for surgical replacement, 
        // so we'll skip it for now to prioritize style preservation
        // and only do simple numeral replacement.
    }

    // 2. Surgical Numeral Replacement (Preserves Formatting)
    const searchPattern = useSmartIgnore ? "[a-zA-Z0-9]{1,}" : "[0-9]{1,}";
    const results = range.search(searchPattern, { matchWildcards: true });
    results.load("items");
    await context.sync();

    for (let i = 0; i < results.items.length; i++) {
      const blockRange = results.items[i];
      blockRange.load("text");
      // Load parentBody type to check for headers
      try { (blockRange as any).parentBody.load("type"); } catch(e) {}
      await context.sync();

      const text = blockRange.text;
      
      // Page number protection
      try {
          if ((text.length === 1 || text.length === 2) && 
              ((blockRange as any).parentBody?.type === "Header" || (blockRange as any).parentBody?.type === "Footer")) continue;
      } catch(e) {}

      if (useSmartIgnore) {
        if (/^[0-9]+$/.test(text)) {
          const thaiText = text.split('').map((char: string) => ARABIC_TO_THAI_MAP[char] || char).join('');
          // REPLACE ONLY THE SMALL RANGE
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
 * Processes a body object safely.
 */
async function processBodyExhaustive(body: Word.Body, useSmartIgnore: boolean, context: Word.RequestContext) {
  if (!body) return;
  try {
    await processRange(body.getRange(), useSmartIgnore, context);
  } catch(e) {}

  try {
    const shapes = body.shapes;
    shapes.load("items/body");
    await context.sync();
    for (let i = 0; i < shapes.items.length; i++) {
      const shape = shapes.items[i];
      if (shape && shape.body) {
        await processBodyExhaustive(shape.body, useSmartIgnore, context);
      }
    }
  } catch (e) {}

  try {
    const fields = body.fields;
    fields.load("items/result");
    await context.sync();
    for (let i = 0; i < fields.items.length; i++) {
      const field = fields.items[i];
      if (field && field.result) {
        await processRange(field.result, useSmartIgnore, context);
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

    // 2. Handle Dynamic Lists (๑.๑) - Silent failure check
    if (dynamicLists) {
      try {
          const body = context.document.body;
          const paragraphs = body.paragraphs;
          paragraphs.load("items/isListItem,items/list");
          await context.sync();

          for (let i = 0; i < paragraphs.items.length; i++) {
            const para = paragraphs.items[i];
            if (para.isListItem && para.list) {
              try {
                const listAny = para.list as any;
                listAny.load("listTemplate/listLevels");
                await context.sync();
                const levels = listAny.listTemplate.listLevels;
                levels.load("items");
                await context.sync();
                for (let j = 0; j < levels.items.length; j++) {
                  levels.items[j].numberStyle = "ThaiArabic";
                }
              } catch (e) {}
            }
          }
      } catch(e) {}
    }

    // 3. Handle Page Numbers & Headers/Footers
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    for (let i = 0; i < sections.items.length; i++) {
      const section = sections.items[i];
      
      if (dynamicPageNumbers) {
        try {
            (section as any).pageNumbering.load("numberStyle");
            await context.sync();
            (section as any).pageNumbering.numberStyle = "ThaiArabic";
            
            const fields = section.body.fields;
            fields.load("items/type");
            await context.sync();
            for (let j = 0; j < fields.items.length; j++) {
                const fType = fields.items[j].type.toString().toLowerCase();
                if (fType === "page" || fType === "toc") {
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
            await processBodyExhaustive(header, useSmartIgnore, context);
            const footer = section.getFooter(type);
            await processBodyExhaustive(footer, useSmartIgnore, context);
          } catch (e) {}
        }
      }
    }
    await context.sync();
  });
}
