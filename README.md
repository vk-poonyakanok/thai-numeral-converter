# Thai Numeral Converter Word Add-in

A modern MS Word Web Add-in built with React + TypeScript to convert Arabic numerals to Thai numerals. Features a "Smart Ignore" algorithm to safely skip numbers embedded in English words, URLs, and email addresses (e.g., `spin9`, `9arm`, `www.site123.com`).

## Features
- **Convert Entire Document**: Scans the whole document body and converts all valid Arabic numerals.
- **Convert Selection**: Only converts numerals within the highlighted text.
- **Smart Ignore**: Intelligent detection to avoid converting numerals that are part of English text, websites, or emails.
- **Formatting Preservation**: Replaces only the numerals, keeping your fonts, colors, and styles intact.

## Tech Stack
- **Framework**: React 19 + TypeScript (Vite)
- **UI System**: Microsoft Fluent UI React
- **Core API**: Word JavaScript API (Office.js)

## Getting Started

### Prerequisites
- Node.js (v18+)
- Microsoft Word (Web or Desktop)

### Local Development

1. **Clone the repository:**
   ```bash
   git clone https://github.com/vitchakorn/thai-numeral-converter.git
   cd thai-numeral-converter
   ```

2. **Install dependencies:**
   ```bash
   npm install
   ```

3. **Start the development server:**
   ```bash
   npm run dev
   ```
   The add-in will be served at `https://localhost:5173/` (using self-signed certificates).

### Sideloading the Add-in

#### Word on the Web
1. Open a document in Word on the Web.
2. Go to **Insert** -> **Add-ins**.
3. Select **Manage My Add-ins** -> **Upload My Add-in**.
4. Upload the `manifest.xml` file located in the project root.

#### Word on Desktop (Mac/Windows)
Please refer to the [official Microsoft documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing) for your specific operating system.

## Project Structure
- `manifest.xml`: The Office Add-in manifest file.
- `src/converter.ts`: Core conversion logic using Office.js.
- `src/App.tsx`: Task pane UI built with Fluent UI.
- `src/test-converter.ts`: Offline test suite for the conversion logic.

## License
MIT
