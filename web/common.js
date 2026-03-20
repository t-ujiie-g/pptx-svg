/**
 * Shared utilities for pptx-svg demo pages.
 */

/** Shorthand for getElementById */
export const $ = (id) => document.getElementById(id);

/** Display a status message */
export function setStatus(msg, type = 'info') {
  $('status-card').style.display = '';
  $('status').className = type;
  $('status').textContent = msg;
}

/** Set up file drop zone and input for PPTX loading */
export function setupDropZone(onFile) {
  const dropZone = $('drop-zone');
  const fileInput = $('file-input');

  dropZone.addEventListener('click', () => fileInput.click());
  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
  });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    if (e.dataTransfer.files[0]) onFile(e.dataTransfer.files[0]);
  });
  fileInput.addEventListener('change', (e) => {
    if (e.target.files[0]) onFile(e.target.files[0]);
  });

  return dropZone;
}

/** Initialize Wasm renderer with drop zone enable/disable */
export async function initRenderer(renderer) {
  const dropZone = $('drop-zone');
  dropZone.style.pointerEvents = 'none';
  dropZone.style.opacity = '0.5';
  setStatus('Loading WebAssembly module...');
  $('status-card').style.display = '';

  try {
    await renderer.init();
    setStatus('Ready. Drop a .pptx file.', 'ok');
    dropZone.style.pointerEvents = '';
    dropZone.style.opacity = '';
  } catch (err) {
    setStatus(`Init failed: ${err.message}`, 'error');
    console.error('[pptx] Init error:', err);
  }
}

/** Set up export button */
export function setupExportButton(renderer) {
  $('export-btn').addEventListener('click', async () => {
    $('export-btn').disabled = true;
    $('export-status').textContent = 'Exporting...';
    try {
      const buffer = await renderer.exportPptx();
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'exported.pptx';
      a.click();
      URL.revokeObjectURL(url);
      $('export-status').textContent = 'Exported!';
      $('export-status').style.color = '#4caf50';
    } catch (err) {
      $('export-status').textContent = `Error: ${err.message}`;
      $('export-status').style.color = '#f44336';
      console.error('[pptx] Export error:', err);
    } finally {
      $('export-btn').disabled = false;
    }
  });
}
