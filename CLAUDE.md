# CLAUDE.md - MIRAI Passport (みらいパスポート)

## Project Overview

MIRAI Passport is a Google Apps Script (GAS) web application for Japanese elementary school education. It generates AI-powered worksheets, manages student submissions, and integrates with a companion system called "みらいコンパス" (Mirai Compass) for class management.

**Current version:** v2.3.0 (Phase 5: AI Grading, Analytics, PDF/OCR)

## Architecture

- **Backend:** Google Apps Script (serverless, deployed as a GAS web app)
- **Frontend:** Single-page HTML app with vanilla JavaScript, Bootstrap 5
- **Database:** Google Sheets (via `SpreadsheetApp` API)
- **AI:** Google Gemini API (gemini-2.5-flash) for worksheet generation, rubric generation, and AI-assisted grading
- **No build system, no npm, no TypeScript** - all dependencies are loaded via CDN or built into GAS

## File Structure

```
MIRAI-Passport/
  code.gs        # Backend: GAS server functions (entry point, DB ops, AI, Compass sync)
  index.html     # HTML template: layout, modals, server-side template injection
  js.html        # Frontend: all client-side JavaScript (included via GAS template)
  css.html       # Styles: CSS variables, animations, responsive, print media queries
```

There are only 4 source files. This is a monolithic GAS project, not a typical Node.js app.

## Key Concepts

### Two Modes
The app runs in either **teacher mode** or **student mode**, controlled by the `mode` URL parameter:
- **Teacher:** Creates/edits worksheets, uses AI generation, views submissions, grades
- **Student:** Views worksheets, draws/writes responses, submits, interacts in "Plaza"

### Google Sheets as Database
Three sheets act as tables:
- `Worksheets`: taskId, unitName, stepTitle, htmlContent, lastUpdated, jsonSource, canvasJson, rubricHtml, isShared
- `Responses`: responseId, taskId, studentId, studentName, submittedAt, canvasImage, textContent, status, feedbackText, score, feedbackJson, canvasJson, isPublic, reactions
- `ImportQueue`: transactionId, dataJson, createdAt

### GAS Template System
- `index.html` uses `<?= ?>` and `<?!= ?>` for server-side template injection
- `css.html` and `js.html` are included via `<?!= include('css') ?>` / `<?!= include('js') ?>`
- Config values (mode, taskId, studentId, studentName) are injected at render time

## Code Organization

### code.gs (Backend)
| Section | Lines | Purpose |
|---------|-------|---------|
| Entry point & init | 14-73 | `doGet()`, `checkSetupStatus()`, `performInitialSetup()`, `ensureSheet()` |
| Config management | 79-94 | `saveUserConfig()`, `getUserConfig()` |
| Database operations | 100-170 | CRUD for Worksheets sheet |
| Student responses | 176-262 | CRUD for Responses sheet, peer reactions |
| AI & utilities | 268-325 | `callGeminiAPI()`, `generateSingleWorksheet()`, `generateRubricAI()` |
| AI grading & analytics | 327-390 | `generateWorksheetWithRubric()`, `generateFeedbackAI()`, `getClassAnalytics()` |
| Compass integration | 395-450 | Import queue processing, sync to Compass |

### js.html (Frontend)
| Object | Purpose |
|--------|---------|
| `Server` | Promise wrapper around `google.script.run` |
| `UI` | Loading overlay, toasts, batch progress |
| `Printer` | Unified A4 print system (opens new window) |
| `Modals` | Dashboard, grading, rubric, settings, import, analytics modals |
| `State` | Global state: currentTask, fabricCanvas, studentId, etc. |
| `App` | Main application logic: init, task selection, AI generation, AI grading, student submission, Plaza, PDF/OCR |
| `Editor` | Rich editing: context menus, table manipulation, resizing |

### css.html (Styles)
Organized into sections: theme variables, loading animation, app layout, editor/worksheet styles, subject-specific modes (kokugo/vertical writing), grid/graph paper, print media queries.

## Frontend Dependencies (CDN)

- Bootstrap 5.3.0 (CSS + JS bundle)
- Bootstrap Icons 1.10.5
- Fabric.js 5.3.1 (canvas drawing)
- PDF.js 2.16.105 (PDF import)
- Pako 2.1.0 (compression)
- SweetAlert2 11 (modal dialogs)
- Google Fonts: Zen Maru Gothic, M PLUS Rounded 1c, Poppins

## Coding Conventions

### Style
- **Module pattern:** Major components are plain objects (`Server`, `UI`, `App`, `State`, `Editor`, etc.)
- **No classes, no modules, no imports** - everything is in global scope within `<script>` tags
- **camelCase** for functions and variables
- **CONSTANT_CASE** for constants (`APP_NAME`, `DB_NAME`)
- **Section dividers:** `// ==================================================` with numbered section headers
- **Comments and UI strings:** All in Japanese

### Patterns
- Server calls use `Server.call('functionName', ...args)` which wraps `google.script.run`
- Mode is determined via URL params with `APP_CONFIG` fallback
- Student identity stored in `localStorage` under key `manabi_sid`
- Auto-save runs every 10 seconds in student mode
- Canvas data is serialized as JSON via Fabric.js `toJSON()`/`loadFromJSON()`

### Error Handling
- Server errors show SweetAlert2 dialogs via the `Server` wrapper
- Loading overlay shown/hidden around async operations
- `try/finally` blocks ensure `UI.hideLoading()` is always called

## Deployment

This is a Google Apps Script web app. There is no build step.

1. Code is deployed through the Google Apps Script editor or `clasp`
2. `doGet(e)` serves the HTML via `HtmlService.createTemplateFromFile('index')`
3. The web app URL is retrieved via `ScriptApp.getService().getUrl()`
4. `setXFrameOptionsMode(ALLOWALL)` enables embedding in iframes

## Git Conventions

- **Branch naming:** `claude/<description>-<id>` for feature work; `main` for production
- **Commit messages:** Mix of conventional commits (`fix:`, `feat:`, `refactor:`) with Japanese descriptions
- **No CI/CD pipeline** configured
- **No pre-commit hooks**

## Testing

There is no automated test framework. Testing is manual. The app provides a sample JSON loader (`App.loadSampleJson()`) for testing the import workflow.

## Common Tasks

### Adding a new server-side function
1. Add the function to `code.gs`
2. Call it from the frontend via `Server.call('functionName', ...args)`
3. The function is automatically exposed to `google.script.run`

### Modifying the UI
1. HTML structure changes go in `index.html`
2. Behavior/logic changes go in `js.html` (typically in the `App` object)
3. Style changes go in `css.html`
4. Use Bootstrap 5 classes for layout; custom CSS for worksheet-specific styling

### Working with the AI prompt
- The main prompt template is in `App.generateSingleWorksheet()` in `js.html` (client-side, ~lines 747-799)
- The server-side version is `generateSingleWorksheet()` in `code.gs` (~lines 286-306)
- Both produce HTML body content for worksheets; the AI output is cleaned by `App.cleanAIOutput()`

## Important Notes

- The Gemini API key is stored per-user via `PropertiesService.getUserProperties()`
- Canvas images are stored as base64 data URLs directly in Google Sheets cells
- The app uses `contenteditable` for WYSIWYG editing (not a framework-based editor)
- Print output targets A4 paper size (210mm x 297mm)
- Student mode hides the sidebar and changes the navbar color to green
