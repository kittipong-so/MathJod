/**
 * MathType Clone - Task Pane Hub & Picture-Object Controller
 * Implements traditional MathType workflow with Office Dialog API,
 * MathLive interactive editing, and Base64 PNG + AltText persistence.
 */

const LATEX_ALT_PREFIX = '[MathType:LaTeX] ';

// State
let isWordHost = false;
let currentDialog = null;
let selectedEquationLatex = null;
let isUpdatingExisting = false;
let lastProcessedActionId = null;

// DOM Elements
const btnEditMath = document.getElementById('btnEditMath');
const btnNewMath = document.getElementById('btnNewMath');
const selectionCard = document.getElementById('selectionCard');
const selectionPill = document.getElementById('selectionPill');
const selectionDesc = document.getElementById('selectionDesc');
const selectedEquationBox = document.getElementById('selectedEquationBox');
const statusBanner = document.getElementById('statusBanner');
const wordStatusBadge = document.getElementById('wordStatusBadge');
const wordStatusText = document.getElementById('wordStatusText');
const renderSandbox = document.getElementById('renderSandbox');

/**
 * Extract LaTeX string from an image's Alt Text Description
 */
function extractLatexFromAltText(altText) {
  if (!altText) return null;

  // Check prefix format: [MathType:LaTeX] <latex>
  if (altText.startsWith(LATEX_ALT_PREFIX)) {
    return altText.substring(LATEX_ALT_PREFIX.length).trim();
  }

  // Check JSON format fallback
  if (altText.startsWith('{') && altText.includes('"latex"')) {
    try {
      const data = JSON.parse(altText);
      return data.latex || null;
    } catch (_) {}
  }

  return null;
}

/**
 * Display visual status banner in the Hub UI
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
 * Update the Hub UI based on whether an editable equation is selected in Word
 */
function updateSelectionUI(latex = null) {
  if (latex) {
    selectedEquationLatex = latex;
    btnEditMath.disabled = false;
    btnEditMath.title = 'Edit the currently selected equation';

    selectionCard.className = 'selection-card active';
    selectionPill.className = 'selection-pill active';
    selectionPill.textContent = 'Equation Selected';
    selectionDesc.textContent = 'Found editable MathType equation in document:';

    selectedEquationBox.textContent = latex;
    selectedEquationBox.style.display = 'block';
  } else {
    selectedEquationLatex = null;
    btnEditMath.disabled = true;
    btnEditMath.title = 'Select an equation in the document to edit';

    selectionCard.className = 'selection-card';
    selectionPill.className = 'selection-pill';
    selectionPill.textContent = 'No Equation';
    selectionDesc.textContent = 'Click on any equation picture in your Word document to select and edit it.';

    selectedEquationBox.textContent = '';
    selectedEquationBox.style.display = 'none';
  }
}

/**
 * Selection Event Listener: Checks if the user clicked on an equation picture
 */
async function onDocumentSelectionChanged() {
  if (!isWordHost || typeof Word === 'undefined') return;

  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const inlinePictures = selection.inlinePictures;
      inlinePictures.load('items/altTextDescription, items/altTextTitle');

      await context.sync();

      if (inlinePictures.items && inlinePictures.items.length > 0) {
        const pic = inlinePictures.items[0];
        const latex = extractLatexFromAltText(pic.altTextDescription);

        if (latex) {
          updateSelectionUI(latex);
          return;
        }
      }

      // No equation image is selected
      updateSelectionUI(null);
    });
  } catch (err) {
    console.warn('Selection check warning:', err);
  }
}

/**
 * Safely close the active dialog
 */
function closeDialogSafely() {
  if (currentDialog) {
    try {
      currentDialog.close();
      console.log('Dialog closed via currentDialog.close()');
    } catch (e) {
      console.warn('currentDialog.close() exception:', e);
    }
    currentDialog = null;
  }
}

/**
 * Open the interactive MathLive editor using Office Dialog API
 */
function openEditorDialog(initialLatex = null) {
  isUpdatingExisting = Boolean(initialLatex);

  // Construct absolute dialog URL (hosted alongside index.html)
  const dialogUrl = new URL('editor.html', window.location.href);
  if (initialLatex) {
    dialogUrl.searchParams.set('latex', encodeURIComponent(initialLatex));
  }

  console.log('Opening editor dialog URL:', dialogUrl.href);

  // Office Dialog API options
  // displayInIframe: true ensures in-app overlay on Word on the web (avoids popup blockers)
  const options = {
    height: 65,
    width: 60,
    displayInIframe: true
  };

  Office.context.ui.displayDialogAsync(dialogUrl.href, options, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error('Failed to open dialog via displayDialogAsync:', asyncResult.error.message);

      // Fallback: If displayInIframe: true failed or was rejected, try fallback with displayInIframe: false
      if (options.displayInIframe) {
        options.displayInIframe = false;
        Office.context.ui.displayDialogAsync(dialogUrl.href, options, (retryResult) => {
          if (retryResult.status === Office.AsyncResultStatus.Failed) {
            showStatus(`Failed to open equation editor: ${retryResult.error.message}`, 'error', 6000);
          } else {
            currentDialog = retryResult.value;
            currentDialog.addEventHandler(Office.EventType.DialogMessageReceived, onDialogMessage);
            currentDialog.addEventHandler(Office.EventType.DialogEventReceived, onDialogEvent);
          }
        });
        return;
      }

      showStatus(`Failed to open equation editor: ${asyncResult.error.message}`, 'error', 6000);
      return;
    }

    currentDialog = asyncResult.value;
    console.log('Dialog successfully opened');

    // Register message and event handlers
    currentDialog.addEventHandler(Office.EventType.DialogMessageReceived, onDialogMessage);
    currentDialog.addEventHandler(Office.EventType.DialogEventReceived, onDialogEvent);
  });
}

/**
 * Universal Dialog Message Processor (Deduplicated across channels)
 */
async function processDialogMessagePayload(rawPayload) {
  if (!rawPayload) return;

  let data;
  try {
    data = typeof rawPayload === 'string' ? JSON.parse(rawPayload) : rawPayload;
  } catch (e) {
    console.warn('Unable to parse dialog payload:', rawPayload);
    return;
  }

  // Deduplication guard
  const msgId = `${data.action}_${data.timestamp || ''}`;
  if (lastProcessedActionId && lastProcessedActionId === msgId) {
    return;
  }
  lastProcessedActionId = msgId;

  console.log('Processing action from dialog:', data.action);

  if (data.action === 'insert') {
    closeDialogSafely();
    await processAndInsertEquation(data.latex);
  } else if (data.action === 'cancel') {
    closeDialogSafely();
    showStatus('Equation edit cancelled.', 'success', 2000);
  }
}

/**
 * 1. Primary Channel: Office.EventType.DialogMessageReceived
 */
function onDialogMessage(arg) {
  console.log('DialogMessageReceived event triggered:', arg);
  const raw = arg && arg.message ? arg.message : arg;
  processDialogMessagePayload(raw);
}

/**
 * Handle dialog events (e.g. user clicked window close button 'X')
 */
function onDialogEvent(arg) {
  console.log('DialogEventReceived event triggered:', arg);
  currentDialog = null;
}

/**
 * 2. Secondary Channel: postMessage listener
 */
window.addEventListener('message', (event) => {
  if (event.data && typeof event.data === 'string' && event.data.includes('"action"')) {
    processDialogMessagePayload(event.data);
  }
});

/**
 * 3. Tertiary Channel: BroadcastChannel listener
 */
if (typeof BroadcastChannel !== 'undefined') {
  const channel = new BroadcastChannel('mathtype_bridge_channel');
  channel.onmessage = (event) => {
    console.log('BroadcastChannel message received:', event.data);
    processDialogMessagePayload(event.data);
  };
}

/**
 * 4. Quaternary Channel: LocalStorage Event Bus listener
 */
window.addEventListener('storage', (e) => {
  if (e.key === 'mathtype_dialog_bridge' && e.newValue) {
    console.log('LocalStorage bridge event received');
    processDialogMessagePayload(e.newValue);
    try {
      localStorage.removeItem('mathtype_dialog_bridge');
    } catch (_) {}
  }
});

/**
 * Wait for MathJax to be fully initialized
 */
async function ensureMathJaxReady() {
  if (window.MathJax && window.MathJax.tex2svgPromise) return;

  for (let i = 0; i < 50; i++) {
    await new Promise(r => setTimeout(r, 100));
    if (window.MathJax && window.MathJax.tex2svgPromise) return;
  }

  throw new Error('MathJax rendering engine is not yet ready. Please check your internet connection.');
}

/**
 * Convert LaTeX string into a high-DPI Base64 PNG image using MathJax SVG + Canvas
 */
async function renderLatexToBase64Png(latex) {
  await ensureMathJaxReady();

  // Generate vector SVG node from LaTeX
  const container = await window.MathJax.tex2svgPromise(latex, { display: true });
  const svg = container.querySelector('svg');
  if (!svg) throw new Error('Could not render SVG from LaTeX expression.');

  // Clone SVG to isolate it
  const clonedSvg = svg.cloneNode(true);
  clonedSvg.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
  clonedSvg.setAttribute('xmlns:xlink', 'http://www.w3.org/1999/xlink');

  // Compute dimensions from SVG attributes or viewBox
  const exWidth = parseFloat(clonedSvg.getAttribute('width')) || 6;
  const exHeight = parseFloat(clonedSvg.getAttribute('height')) || 2;

  // Conversion: ~10px per ex at standard font scaling
  const baseWidthPx = Math.max(Math.round(exWidth * 10), 30);
  const baseHeightPx = Math.max(Math.round(exHeight * 10), 16);

  // High-DPI scale (3x) for razor-sharp appearance in Word
  const renderScale = 3;
  const canvasWidth = baseWidthPx * renderScale;
  const canvasHeight = baseHeightPx * renderScale;

  clonedSvg.setAttribute('width', canvasWidth);
  clonedSvg.setAttribute('height', canvasHeight);

  // Serialize SVG to base64 Data URL
  const serializer = new XMLSerializer();
  const svgString = serializer.serializeToString(clonedSvg);
  const svgDataUrl = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svgString)));

  // Render SVG onto HTML5 Canvas to produce pure PNG
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.crossOrigin = 'anonymous';

    img.onload = () => {
      const canvas = document.createElement('canvas');
      canvas.width = canvasWidth;
      canvas.height = canvasHeight;
      const ctx = canvas.getContext('2d');

      // Draw vector SVG onto canvas
      ctx.drawImage(img, 0, 0, canvasWidth, canvasHeight);

      // Extract raw Base64 PNG string (without 'data:image/png;base64,' prefix for Word API)
      const dataUrl = canvas.toDataURL('image/png');
      const base64Png = dataUrl.replace(/^data:image\/png;base64,/, '');

      // Word dimensions are specified in points (1 pt = 0.75 px)
      const pointWidth = Math.round(baseWidthPx * 0.75);
      const pointHeight = Math.round(baseHeightPx * 0.75);

      resolve({
        base64: base64Png,
        pointWidth: pointWidth,
        pointHeight: pointHeight
      });
    };

    img.onerror = (e) => reject(new Error('Failed to rasterize equation into Canvas: ' + e));
    img.src = svgDataUrl;
  });
}

/**
 * Render equation into Base64 image and inject into Word with LaTeX in AltText
 */
async function processAndInsertEquation(latex) {
  if (!latex || !latex.trim()) {
    showStatus('Equation was empty. Nothing inserted.', 'error');
    return;
  }

  showStatus('Rendering and inserting equation...', 'success', 2000);

  try {
    // 1. Render LaTeX to high-DPI Base64 PNG
    const { base64, pointWidth, pointHeight } = await renderLatexToBase64Png(latex);

    // 2. Insert into Word document
    if (isWordHost && typeof Word !== 'undefined' && Word.run) {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();

        // Check if an existing picture is currently selected to replace
        const pictures = selection.inlinePictures;
        pictures.load('items/altTextDescription');
        await context.sync();

        let insertedPic;
        if (pictures.items.length > 0 && extractLatexFromAltText(pictures.items[0].altTextDescription)) {
          // Replace selected equation picture directly in place
          const targetPic = pictures.items[0];
          insertedPic = targetPic.insertInlinePictureFromBase64(base64, Word.InsertLocation.replace);
        } else {
          // Insert at cursor position
          insertedPic = selection.insertInlinePictureFromBase64(base64, Word.InsertLocation.replace);
        }

        // 3. Store source LaTeX invisibly inside Alt Text Description
        insertedPic.altTextTitle = 'MathType Equation';
        insertedPic.altTextDescription = `${LATEX_ALT_PREFIX}${latex}`;

        // Set typographical display dimensions
        insertedPic.width = pointWidth;
        insertedPic.height = pointHeight;
        insertedPic.lockAspectRatio = true;

        await context.sync();
      });

      showStatus(
        isUpdatingExisting ? 'Equation updated successfully!' : 'Equation inserted into document!',
        'success'
      );

      // Re-trigger selection check to sync Hub state
      setTimeout(onDocumentSelectionChanged, 300);
    } else {
      // Browser preview mode
      console.log('Base64 PNG Generated (first 100 chars):', base64.slice(0, 100));
      console.log('AltText LaTeX Stored:', `${LATEX_ALT_PREFIX}${latex}`);
      showStatus('Preview mode: Equation rendered. Sideload in Word to insert into document.', 'success', 5000);
    }
  } catch (err) {
    console.error('Failed to process and insert equation:', err);
    showStatus(`Failed: ${err.message || 'Unknown error'}`, 'error', 6000);
  }
}

/**
 * Setup Hub UI Event Listeners
 */
function setupEventListeners() {
  // "Edit Math" button: opens editor with current selection's LaTeX
  btnEditMath.addEventListener('click', () => {
    if (selectedEquationLatex) {
      openEditorDialog(selectedEquationLatex);
    }
  });

  // "New Equation" button: opens editor with blank/new template
  btnNewMath.addEventListener('click', () => {
    openEditorDialog(null);
  });
}

/**
 * Office.js Initialization
 */
Office.onReady((info) => {
  if (info && info.host === Office.HostType.Word) {
    isWordHost = true;
    wordStatusBadge.className = 'badge badge-online';
    wordStatusText.textContent = 'Connected';

    // Register DocumentSelectionChanged listener for the editing loop
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      onDocumentSelectionChanged,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error('Failed to register DocumentSelectionChanged:', result.error.message);
        }
      }
    );

    // Initial check in case an equation is already selected upon opening
    onDocumentSelectionChanged();
  } else {
    isWordHost = false;
    wordStatusBadge.className = 'badge badge-offline';
    wordStatusText.textContent = 'Preview Mode';
    wordStatusBadge.title = 'Running in browser preview mode. Sideload in Word to enable document selection.';
  }

  setupEventListeners();
});
