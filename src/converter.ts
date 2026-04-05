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
    processedText = processedText.replace(dateRegex, (_, d1, month, d2, year) => {
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
 * Core logic to process a range SURGICALLY.
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
 * Main Conversion Functions
 */
export async function convertSelection(useSmartIgnore: boolean) {
  await Word.run(async (context: Word.RequestContext) => {
    const selection = context.document.getSelection();
    await processRange(selection, useSmartIgnore, context);
    await context.sync();
  });
}

export async function convertMainBody(useSmartIgnore: boolean) {
  await Word.run(async (context: Word.RequestContext) => {
    await processRange(context.document.body.getRange(), useSmartIgnore, context);
    await context.sync();
  });
}

/**
 * Advanced Elements Processing (Flattening)
 */
async function processDeepBody(body: Word.Body, useSmartIgnore: boolean, context: Word.RequestContext) {
  if (!body) return;

  // 1. Flatten Lists (๑.๑)
  const paragraphs = body.paragraphs;
  paragraphs.load("items/isListItem,items/listItem");
  await context.sync();

  for (let i = 0; i < paragraphs.items.length; i++) {
    const para = paragraphs.items[i];
    if (para.isListItem) {
      try {
        para.listItem.load("listString");
        await context.sync();
        const listString = para.listItem.listString;
        const thaiLabel = convertText(listString, false);
        para.detachFromList();
        para.insertText(thaiLabel + " ", "Start");
      } catch (e) {}
    }
  }

  // 2. Process text
  await processRange(body.getRange(), useSmartIgnore, context);

  // 3. Shapes
  const shapes = body.shapes;
  shapes.load("items/body");
  await context.sync();
  for (let i = 0; i < shapes.items.length; i++) {
    const shape = shapes.items[i];
    if (shape && shape.body) {
      await processDeepBody(shape.body, useSmartIgnore, context);
    }
  }

  // 4. Fields (captions, page numbers)
  const fields = body.fields;
  fields.load("items/result");
  await context.sync();
  for (let i = 0; i < fields.items.length; i++) {
    const field = fields.items[i];
    if (field && field.result) {
      await processRange(field.result, useSmartIgnore, context);
    }
  }
}

export async function flattenAdvancedElements(useSmartIgnore: boolean) {
  await Word.run(async (context: Word.RequestContext) => {
    // 1. Headers/Footers
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
          await processDeepBody(section.getHeader(type), useSmartIgnore, context);
          await processDeepBody(section.getFooter(type), useSmartIgnore, context);
        } catch (e) {}
      }
    }

    // 2. Shapes in main body
    await processDeepBody(context.document.body, useSmartIgnore, context);
    
    await context.sync();
  });
}
