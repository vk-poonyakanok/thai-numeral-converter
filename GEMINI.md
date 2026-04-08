# IT๙ Converter Project Mandates

This document contains foundational mandates and project-specific context for the **IT๙ Converter** Word Add-in. Adhere to these instructions over general defaults.

## Project Identity
- **Name**: IT๙ Converter (Short name: IT๙)
- **Developer**: Blue (Vitchakorn Poonyakanok)
- **Version**: 1.22.0

## Technical Context
- **Stack**: React 19 + TypeScript (Vite)
- **API**: Word JavaScript API (Office.js)
- **Deployment**: Hosted on GitHub Pages at `https://vk-poonyakanok.github.io/thai-numeral-converter/`
- **Sideloading Path (Mac)**: `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/manifest.xml`

## Assets & Icons
- **Partner Center Compliance**: To pass Microsoft Partner Center validation, the following icons must be maintained in the `public/` directory and correctly referenced in `manifest.xml`:
  - `icon-16.png` (16x16)
  - `icon-32.png` (32x32)
  - `icon-64.png` (64x64)
  - `icon-80.png` (80x80)
- **High Resolution**: Always use `icon-64.png` for the `HighResolutionIconUrl` in the manifest.

## Compliance & Documentation
- **Legal Pages**: Maintain `privacy.html` and `eula.html` in the `public/` directory.
- **Screenshots**: Microsoft Partner Center screenshots must be exactly **1366 x 768 px** and under **1024 KB**. Use high-quality resizing (Lanczos filter) to ensure text readability.

## Engineering Mandates
- **Formatting Preservation**: Always use **Surgical Replacement** (searching for specific digit ranges using `range.search("[0-9]{1,}", ...)`) instead of replacing entire paragraphs. This ensures bold, italic, and colored text remains intact.
- **Smart Ignore Logic**: Strictly maintain the regex `/(?<![a-zA-Z0-9])[0-9]+(?![a-zA-Z0-9])/g` to avoid converting numbers embedded in English words, URLs, or emails.
- **VBA Parity**: The "Flatten" features should mimic the behavior of the `UltimateThaiConverter` VBA script, specifically unlinking fields and freezing auto-lists into plain text Thai numerals.
- **Type Safety**: Maintain strict TypeScript standards. Use `any` casting only as a last resort for advanced Office.js properties not yet defined in the `@types/office-js` package.

## Development Workflow
- **Local Testing**: Always run `npm run build` locally before pushing to GitHub to catch potential CI failures early.
- **Deployment**: Pushing to the `main` branch triggers a GitHub Action that deploys to `gh-pages`.
