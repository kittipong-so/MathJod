# Math typing addon - Traditional Picture-Object Architecture

A Microsoft Word Task Pane Add-in that implements the classic **MathType picture-object workflow**:

1. **Hub Task Pane**: Simple, focused hub with a context-sensitive **"Edit Math"** button that automatically enables when an equation is selected in the Word document.
2. **Interactive Editor (Office Dialog API)**: When clicking "Edit Math" or "New Equation", a floating dialog window opens powered by the **MathLive** WYSIWYG math editor.
3. **Insertion & Persistence**: Equations are rasterized client-side into ultra-crisp 300+ DPI Base64 PNG images and inserted into the document using `context.document.getSelection().insertInlinePictureFromBase64()`. The source LaTeX is stored invisibly inside the image's `altTextDescription` attribute (`[MathType:LaTeX] <latex>`).
4. **The Editing Loop**: An `Office.EventType.DocumentSelectionChanged` listener continuously monitors document selections. When the user clicks an equation picture in the document, its hidden LaTeX is detected, the "Edit Math" button lights up, and clicking it re-opens the MathLive editor with the existing equation pre-populated for modification.

100% client-side (HTML, CSS, JS) suitable for hosting on **GitHub Pages**.

---

## 📁 Architecture & File Overview

```
mathtype_clone/
├── manifest.xml       # Add-in manifest with Dialog API permissions & ribbon button
├── index.html         # Task Pane Hub UI (selection status, Edit Math & New Equation)
├── app.js             # Hub logic, selection listener, MathJax/Canvas rasterizer & Word API
├── editor.html        # Floating Dialog with MathLive interactive visual math editor
├── assets/            # Add-in ribbon icons (16x16, 32x32, 80x80)
└── README.md          # Documentation
```

---

## 🔄 The Complete Workflow

```
[ Document Selection Listener ]
       │  (Office.EventType.DocumentSelectionChanged)
       ▼
[ Click Equation in Word ] ──► Read image.altTextDescription
       │                       (Tag: [MathType:LaTeX] <latex>)
       ▼
[ "Edit Math" Enabled in Hub ]
       │
       ▼  (Office.context.ui.displayDialogAsync)
[ Floating Dialog (editor.html) ]
       │  (MathLive Visual <math-field> Editor)
       ▼  (Office.context.ui.messageParent)
[ LaTeX Returned to Hub ]
       │
       ▼  (MathJax Vector SVG ──► HTML5 Canvas ──► 3x High-DPI PNG)
[ Base64 PNG + AltText ]
       │
       ▼  (range.insertInlinePictureFromBase64 / Replace)
[ Word Document Updated ]
```

---

## 🌐 GitHub Pages Deployment

1. Commit and push these files to your GitHub repository.
2. Enable GitHub Pages in your repo under **Settings > Pages** (branch `main`, folder `/`).
3. In `manifest.xml`, replace:
   `https://YOUR-GITHUB-USERNAME.github.io/YOUR-REPO-NAME/`
   with your live GitHub Pages URL.
4. Sideload `manifest.xml` into Word (Desktop via Shared Folder / `npx office-addin-debugging start manifest.xml`, or Word for Web via Upload My Add-in).
