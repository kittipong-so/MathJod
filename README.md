# Microsoft Word Task Pane Add-in for Math Equations

A client-side Microsoft Word Task Pane Add-in that allows users to type standard LaTeX mathematical equations, view real-time rendering previews via KaTeX, and insert them directly into Microsoft Word documents as native, fully-editable Word equations (OMML).

Designed to run **100% client-side** without any Node.js or backend servers, making it ideal for hosting on **GitHub Pages**.

---

## 🚀 How It Works

Word uses **Office Math Markup Language (OMML)** wrapped in **Office Open XML (OOXML)** for native math objects. Raw MathML or raw LaTeX strings cannot be injected directly via `insertOoxml()`.

This add-in bridges the web and Office math standards entirely in the browser:

```
[ LaTeX Input ]
       │
       ▼  (KaTeX client-side engine)
[ MathML (<math>) ] + [ Live KaTeX Rendered Preview ]
       │
       ▼  (Clean semantics + mathml2omml)
[ OMML (<m:oMath>) ]
       │
       ▼  (Wrap into Open Packaging Conventions Flat OPC envelope)
[ OOXML (<pkg:package>) ]
       │
       ▼  (Word.run + context.document.getSelection().insertOoxml)
[ Native Editable Word Equation ]
```

---

## 📁 Project Structure

```
mathtype_clone/
├── manifest.xml       # Office Add-in XML manifest with ribbons and permissions
├── index.html         # Task pane UI with KaTeX and Office.js CDN links
├── app.js             # Conversion, rendering, and Office.js injection logic
├── assets/            # App icons (16x16, 32x32, 80x80)
└── README.md          # Setup, deployment, and sideloading documentation
```

---

## 🌐 GitHub Pages Deployment

1. **Create a GitHub Repository**:
   - Push these files (`manifest.xml`, `index.html`, `app.js`, and `assets/`) to your repository (e.g. `https://github.com/<YOUR-USERNAME>/<YOUR-REPO-NAME>`).

2. **Enable GitHub Pages**:
   - Go to your repository **Settings** > **Pages**.
   - Under **Build and deployment** > **Branch**, select `main` (or `master`) and folder `/ (root)`.
   - Click **Save**.
   - Your site URL will look like: `https://<YOUR-USERNAME>.github.io/<YOUR-REPO-NAME>/index.html`.

3. **Update `manifest.xml`**:
   - Open [manifest.xml](file:///d:/mathtype_clone/manifest.xml) and replace all instances of:
     `https://YOUR-GITHUB-USERNAME.github.io/YOUR-REPO-NAME/`
     with your actual GitHub Pages URL.

---

## 💻 Sideloading into Microsoft Word

### Method A: Word for the Web (Fastest for testing)
1. Open any document in [Word for the Web](https://office.live.com/start/Word.aspx).
2. Go to **Insert** > **Add-ins** > **Upload My Add-in**.
3. Browse and select your updated `manifest.xml`.
4. The **MathType** button will appear under the **Insert** tab in the ribbon.

### Method B: Word Desktop (Windows)
1. Share the folder containing `manifest.xml` as a network share (e.g., `\\localhost\MyAddins`).
2. In Word, go to **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**.
3. Enter the network share path (e.g., `\\localhost\MyAddins`), click **Add Catalog**, and check **Show in Menu**.
4. Restart Word.
5. Go to **Insert** > **Get Add-ins** (or **My Add-ins**) > **Shared Folder**, select **MathType Clone**, and click **Add**.

---

## ⌨️ Features & Shortcuts

- **Live KaTeX Preview**: Real-time rendering as you type with instant syntax validation.
- **Math Symbol Palette**: Quick-insert buttons for fractions, powers, roots, integrals, sums, matrices, and Greek letters.
- **Display vs. Inline Mode**: Choose between centered display block equations (`<m:oMathPara>`) or inline equations (`<m:oMath>`).
- **Keyboard Shortcut**: Press `Ctrl + Enter` (or `Cmd + Enter` on macOS) inside the LaTeX box to immediately insert the equation into Word.
- **Standalone Preview**: Open `index.html` directly in any web browser to test formula rendering and conversion outside Word.
