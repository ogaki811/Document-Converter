# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a browser-based document converter that transforms Excel and Word files into PNG images entirely client-side (no server upload). It uses React 19, Vite, and Tailwind CSS v4 with a premium dark-mode design.

**Key libraries:**
- `xlsx` (SheetJS) - Excel file parsing
- `mammoth` - Word (.docx) file parsing
- `html2canvas` - HTML to image conversion
- `framer-motion` - UI animations
- `lucide-react` - Icons

## Development Commands

```bash
# Install dependencies
npm install

# Start development server (usually http://localhost:5173)
npm run dev

# Build for production
npm run build

# Preview production build
npm run preview

# Lint code
npm run lint
```

## Architecture

### Single Component Design
The entire application lives in `src/App.jsx` with all conversion logic self-contained. This is intentional for simplicity - there's no need to split into multiple components unless complexity significantly increases.

### File Processing Flow
1. **File Upload** → User drops/selects files (validated for `.xlsx`, `.xls`, `.docx`)
2. **File State** → Each file gets status: `idle` → `converting` → `success`/`error`
3. **Conversion Pipeline**:
   - Excel: File → SheetJS parse → HTML table → html2canvas → PNG download
   - Word: File → Mammoth HTML conversion → html2canvas → PNG download
4. **Image Generation** → Uses isolated iframe to prevent Tailwind v4 color conflicts (oklch issue)

### Key Implementation Details

**Iframe Isolation (App.jsx:113-200)**
The `generateImage` function creates a temporary, hidden iframe to render HTML content before capturing with html2canvas. This isolation is critical because:
- Prevents Tailwind CSS v4 (which uses oklch colors) from affecting the rendered document
- Provides clean white background with simple Arial font styling
- iframe is positioned off-screen and removed after capture to avoid memory leaks

**File State Management**
Files array structure in state:
```javascript
{
  id: string,           // Random unique ID
  file: File,           // Original File object
  name: string,         // Display name
  type: 'excel'|'word', // Determines icon and conversion logic
  size: string,         // Formatted size (MB)
  status: 'idle'|'converting'|'success'|'error',
  errorMessage: string  // Error details if conversion fails
}
```

**First Sheet Only**
Excel conversion only processes the first sheet (`workbook.SheetNames[0]`). This is a deliberate simplification - multi-sheet support would require UI changes for sheet selection.

## Styling

- **Framework**: Tailwind CSS v4 with custom primary color palette
- **Design System**: Dark mode base (slate-950 bg), cyan-to-blue gradients for accents
- **Font**: Inter (defined in tailwind.config.js)
- **Animation**: Framer Motion for file list enter/exit animations

## Important Constraints

- **Word format**: Only `.docx` is supported (not legacy `.doc`)
- **Browser dependencies**: Requires modern browser with JavaScript enabled
- **Memory considerations**: Very large files may cause performance issues since all processing is client-side
- **Layout limitations**: Complex Excel sheets or Word documents with intricate layouts may not render perfectly (browser-based HTML rendering limitations)

## Future Considerations

If extending functionality, consider:
- Multi-sheet Excel support (requires sheet selector UI)
- Page-by-page Word conversion for long documents
- Progress indicators for large file processing
- Batch conversion ("Convert All" button)
- Output quality/resolution settings
