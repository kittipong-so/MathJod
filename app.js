/**
 * MathType Clone - Microsoft Word Task Pane Add-in
 * Pure client-side LaTeX -> MathML -> OMML -> OOXML equation editor
 */

// Global converter instance
let mml2ommlConverter = null;
let isWordHost = false;

// DOM Elements
const latexInput = document.getElementById('latexInput');
const previewContainer = document.getElementById('previewContainer');
const errorContainer = document.getElementById('errorContainer');
const btnInsert = document.getElementById('btnInsert');
const insertBtnText = document.getElementById('insertBtnText');
const insertSpinner = document.getElementById('insertSpinner');
const insertIcon = document.getElementById('insertIcon');
const statusBanner = document.getElementById('statusBanner');
const presetSelect = document.getElementById('presetSelect');
const btnClear = document.getElementById('btnClear');
const wordStatusBadge = document.getElementById('wordStatusBadge');
const wordStatusText = document.getElementById('wordStatusText');
const mathModeRadios = document.querySelectorAll('input[name="mathMode"]');

/**
 * Initialize converter from CDN (with unpkg fallback)
 */
async function loadConverter() {
  if (mml2ommlConverter) return mml2ommlConverter;
  try {
    const mod = await import('https://cdn.jsdelivr.net/npm/mathml2omml/+esm');
    mml2ommlConverter = mod.mml2omml;
    return mml2ommlConverter;
  } catch (err) {
    console.warn('Failed loading mathml2omml from jsDelivr, attempting unpkg fallback...', err);
    try {
      const mod = await import('https://unpkg.com/mathml2omml/dist/index.js');
      mml2ommlConverter = mod.mml2omml;
      return mml2ommlConverter;
    } catch (fallbackErr) {
      console.error('Failed loading mathml2omml from all CDNs:', fallbackErr);
      throw new Error('MathML to OMML converter could not be loaded.');
    }
  }
}

/**
 * Clean KaTeX MathML output into pure MathML for OMML conversion
 */
function cleanKaTeXMathML(rawMathML) {
  let mathml = rawMathML;

  // Extract the root <math>...</math> tag
  const mathMatch = mathml.match(/<math[\s\S]*?<\/math>/i);
  if (mathMatch) {
    mathml = mathMatch[0];
  }

  // Strip KaTeX <annotation> tag containing raw TeX source
  mathml = mathml.replace(/<annotation[\s\S]*?<\/annotation>/gi, '');

  // Strip <semantics> wrapper if present
  mathml = mathml.replace(/<semantics>/gi, '').replace(/<\/semantics>/gi, '');

  // Ensure MathML namespace is properly declared
  if (!mathml.includes('xmlns="http://www.w3.org/1998/Math/MathML"')) {
    mathml = mathml.replace('<math', '<math xmlns="http://www.w3.org/1998/Math/MathML"');
  }

  return mathml.trim();
}

/**
 * Build a valid Office Open XML (OOXML) Flat OPC package string containing OMML
 */
function buildOoxmlPackage(ommlString, isDisplay = true) {
  // Wrap in <m:oMathPara> for display (centered block) or use <m:oMath> for inline
  const mathBlock = isDisplay
    ? `<m:oMathPara>${ommlString}</m:oMathPara>`
    : ommlString;

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <w:body>
          <w:p>
            ${mathBlock}
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
}

/**
 * Get selected math format mode (display or inline)
 */
function isDisplayMode() {
  const checked = document.querySelector('input[name="mathMode"]:checked');
  return checked ? checked.value === 'display' : true;
}

/**
 * Render LaTeX in real-time using KaTeX
 */
function updatePreview() {
  const latex = latexInput.value.trim();

  // If input is empty, show default placeholder
  if (!latex) {
    previewContainer.innerHTML = '<span class="preview-placeholder">Live preview will appear here...</span>';
    errorContainer.style.display = 'none';
    btnInsert.disabled = true;
    return;
  }

  try {
    if (typeof katex === 'undefined') {
      previewContainer.textContent = latex;
      return;
    }

    // Render KaTeX with displayMode setting
    katex.render(latex, previewContainer, {
      displayMode: isDisplayMode(),
      throwOnError: true
    });

    errorContainer.style.display = 'none';
    btnInsert.disabled = false;
  } catch (err) {
    // Show syntax error without breaking the interface
    errorContainer.textContent = err.message || 'LaTeX syntax error';
    errorContainer.style.display = 'block';

    // Attempt a tolerant render so user sees partial progress
    try {
      katex.render(latex, previewContainer, {
        displayMode: isDisplayMode(),
        throwOnError: false
      });
    } catch (_) {
      // Ignore fallback errors
    }
  }
}

/**
 * Show temporary status message banner
 */
function showStatus(message, type = 'success', duration = 4000) {
  statusBanner.textContent = message;
  statusBanner.className = `status-banner ${type}`;
  statusBanner.style.display = 'flex';

  if (duration > 0) {
    setTimeout(() => {
      statusBanner.style.display = 'none';
    }, duration);
  }
}

/**
 * Insert LaTeX snippet into textarea at current cursor position
 */
function insertSnippet(snippet) {
  const startPos = latexInput.selectionStart;
  const endPos = latexInput.selectionEnd;
  const currentValue = latexInput.value;

  const newValue = currentValue.substring(0, startPos) + snippet + currentValue.substring(endPos);
  latexInput.value = newValue;

  // Move cursor to sensible position inside brackets if present
  let newCursorPos = startPos + snippet.length;
  const firstBrace = snippet.indexOf('{');
  if (firstBrace !== -1) {
    newCursorPos = startPos + firstBrace + 1;
  }

  latexInput.focus();
  latexInput.setSelectionRange(newCursorPos, newCursorPos);
  updatePreview();
}

/**
 * Convert LaTeX to OMML via MathML and insert OOXML into Word
 */
async function insertEquationIntoWord() {
  const latex = latexInput.value.trim();
  if (!latex) return;

  // UI loading state
  btnInsert.disabled = true;
  insertSpinner.style.display = 'inline-block';
  insertIcon.style.display = 'none';
  insertBtnText.textContent = 'Inserting...';

  try {
    // 1. Generate MathML from LaTeX using KaTeX client-side
    const rawMathML = katex.renderToString(latex, {
      output: 'mathml',
      throwOnError: false
    });

    const cleanMathML = cleanKaTeXMathML(rawMathML);

    // 2. Convert MathML to OMML using mathml2omml
    const converter = await loadConverter();
    const omml = converter(cleanMathML);

    if (!omml || !omml.includes('<m:oMath')) {
      throw new Error('Failed to generate valid Office Math (OMML) from equation.');
    }

    // 3. Wrap OMML inside Word's Flat OPC OOXML package
    const ooxml = buildOoxmlPackage(omml, isDisplayMode());

    // 4. Insert into Word using Office.js
    if (isWordHost && typeof Word !== 'undefined' && Word.run) {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.insertOoxml(ooxml, Word.InsertLocation.replace);
        await context.sync();
      });

      showStatus('Equation successfully inserted into Word!', 'success');
    } else {
      // Browser preview mode (running outside Word)
      console.log('Generated OOXML:\n', ooxml);
      showStatus(
        'Preview mode: Word host not detected. Sideload this add-in inside Word to insert equations directly.',
        'error',
        6000
      );
    }
  } catch (err) {
    console.error('Error inserting equation:', err);
    showStatus(`Insertion failed: ${err.message || 'Unknown error'}`, 'error', 6000);
  } finally {
    // Restore UI state
    btnInsert.disabled = false;
    insertSpinner.style.display = 'none';
    insertIcon.style.display = 'inline-block';
    insertBtnText.textContent = 'Insert into Word';
  }
}

/**
 * Attach UI event listeners
 */
function setupEventListeners() {
  // Real-time preview update on typing
  latexInput.addEventListener('input', updatePreview);

  // Math mode radio toggle (display vs inline)
  mathModeRadios.forEach((radio) => {
    radio.addEventListener('change', updatePreview);
  });

  // Palette button clicks
  document.querySelectorAll('.math-btn').forEach((button) => {
    button.addEventListener('click', () => {
      const snippet = button.dataset.latex;
      if (snippet) insertSnippet(snippet);
    });
  });

  // Preset formula dropdown
  presetSelect.addEventListener('change', (e) => {
    if (e.target.value) {
      latexInput.value = e.target.value;
      updatePreview();
      e.target.selectedIndex = 0;
    }
  });

  // Clear button
  btnClear.addEventListener('click', () => {
    latexInput.value = '';
    latexInput.focus();
    updatePreview();
  });

  // Main insert button
  btnInsert.addEventListener('click', insertEquationIntoWord);

  // Keyboard shortcut: Ctrl + Enter to insert
  latexInput.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
      e.preventDefault();
      insertEquationIntoWord();
    }
  });
}

/**
 * Office.js Initialization
 */
Office.onReady((info) => {
  // Check if running inside Word
  if (info && info.host === Office.HostType.Word) {
    isWordHost = true;
    wordStatusBadge.className = 'badge badge-online';
    wordStatusText.textContent = 'Connected to Word';
  } else {
    isWordHost = false;
    wordStatusBadge.className = 'badge badge-offline';
    wordStatusText.textContent = 'Preview Mode';
    wordStatusBadge.title = 'Running in browser preview mode. Sideload in Word to insert equations.';
  }

  // Pre-load converter in background
  loadConverter().catch((err) => console.warn('Preloading converter warning:', err));

  // Setup UI events
  setupEventListeners();

  // Initial preview check
  updatePreview();
});
