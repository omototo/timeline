// A non-fatal error banner, pinned to the bottom of the pane.
//
// Reconcile failures happen in async Excel.run calls, below React — an
// ErrorBoundary can't catch them. Rather than swallow them (which is how
// "nothing happens when I click preview" looked), surface the host's message
// here so a failed Excel write is visible and debuggable.

const BANNER_ID = 'timeline-error-banner';
const TEXT_ID = 'timeline-error-banner-text';

export function showErrorBanner(message: string): void {
  const doc = globalThis.document;
  let banner = doc.getElementById(BANNER_ID);
  if (!banner) {
    banner = doc.createElement('div');
    banner.id = BANNER_ID;
    banner.setAttribute('role', 'alert');
    Object.assign(banner.style, {
      position: 'fixed',
      left: '0',
      right: '0',
      bottom: '0',
      background: '#7f1d1d',
      color: '#fee2e2',
      font: '11px ui-monospace, monospace',
      padding: '8px 24px 8px 10px',
      zIndex: '9999',
      whiteSpace: 'pre-wrap',
      maxHeight: '45%',
      overflow: 'auto',
    });

    const text = doc.createElement('span');
    text.id = TEXT_ID;
    banner.appendChild(text);

    const close = doc.createElement('button');
    close.textContent = '×';
    close.setAttribute('aria-label', 'Dismiss error');
    Object.assign(close.style, {
      position: 'absolute',
      top: '4px',
      right: '6px',
      background: 'transparent',
      border: '0',
      color: '#fee2e2',
      cursor: 'pointer',
      fontSize: '14px',
    });
    close.addEventListener('click', () => banner?.remove());
    banner.appendChild(close);

    doc.body.appendChild(banner);
  }
  const text = doc.getElementById(TEXT_ID);
  if (text) {
    text.textContent = `Timeline: ${message}`;
  }
}
