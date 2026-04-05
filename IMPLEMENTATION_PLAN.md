# Implementation Plan: Thai Numeral Converter Word Add-in

## Objective
Build a modern MS Word Web Add-in (React + TypeScript) to convert Arabic numerals to Thai numerals. The core feature will be a "Smart Ignore" algorithm to safely skip numbers embedded in English words, URLs, and email addresses (e.g., `spin9`, `9arm`, `www.site123.com`).

## Directory & Repository
- **Local Path**: `/Users/vitchakorn/Documents/GitHub/thai-numeral-converter`
- **GitHub Repo Name**: `thai-numeral-converter`

## Proposed Solution & Architecture
- **Framework**: React + TypeScript (Vite).
- **UI System**: Microsoft Fluent UI React (matches native Word design).
- **Core API**: Word JavaScript API (`Office.js`).

### "Smart Ignore" Algorithm Design
Since standard Word wildcard searches lack advanced Regex Lookarounds, the conversion logic will be handled as follows to preserve formatting:
1. Traverse the document by `Paragraph` objects.
2. Extract the raw text of each paragraph.
3. Use a robust JavaScript Regular Expression to find valid number matches.
   - Regex concept: `/(?<![a-zA-Z])[0-9]+(?![a-zA-Z])/g` (Find digits that do NOT have English letters immediately before or after them).
   - Additional checks can be applied to detect if the string block is part of a URL or Email by splitting the paragraph into words and evaluating the context.
4. For each valid match, map its index back to the Word Range and use `range.insertText(thaiText, "Replace")` to swap the text while perfectly maintaining the original fonts, colors, and sizes.

## Implementation Steps

### Phase 1: Initialization & Repository Setup
1. Scaffolded project with Vite in `/Users/vitchakorn/Documents/GitHub/thai-numeral-converter`.
2. Configure project for Office Add-in (Manifest, Office.js types).
3. Initialize Git.

### Phase 2: UI Development (Task Pane)
1. Clean up the default template in `src/App.tsx`.
2. Build the UI using Fluent UI components:
   - **Header**: Application title and logo.
   - **Buttons**: "Convert Entire Document" and "Convert Selection".
   - **Toggles/Checkboxes**: "Smart Ignore (Keep English/URLs safe)".

### Phase 3: Core Conversion Logic (`Office.js`)
1. Create a `converter.ts` utility file.
2. Define the character mapping (0-9 to ๐-๙).
3. Implement `convertSelection(useSmartIgnore: boolean)`:
   - Get the user's current selection.
   - Split into paragraphs.
   - Apply the Smart Ignore regex logic.
   - Execute replacements.
4. Implement `convertDocument(useSmartIgnore: boolean)`:
   - Iterate through document body, headers, and footers.
   - Apply the same robust replacement logic.

### Phase 4: Testing & Polish
1. Test edge cases: `spin9`, `9arm`, `test123test`, `123,456.78`, and `1.1.1` (list numbering).
2. Update `README.md` with installation instructions via Manifest sideloading.

## Verification
- Numbers adjacent to English text remain unchanged.
- Formatted text (e.g., bold numbers) retains its formatting after conversion.
- No errors are thrown when converting documents with complex tables or headers.
