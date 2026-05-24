import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { showModal, showInfoPopup, showWarningPopup } from './harness.js';

/**
 * showModal XSS-defense tests (task #31).
 *
 * Before the fix, showModal interpolated the caller-supplied title
 * and message directly into modal.innerHTML, so any future call
 * site that piped user input through (e.g. a task name) became an
 * XSS hole.
 *
 * Contract after the fix:
 *  - title and message render via textContent — HTML in the
 *    inputs appears as literal text, no <img>/<script>/event
 *    handlers ever execute.
 *  - The dialog has role="dialog", aria-modal="true", and
 *    aria-labelledby pointing at the title (bonus from a11y task #41).
 */

function cleanupModal() {
  const m = document.getElementById('custom-modal');
  if (m) m.remove();
}

describe('showModal (XSS defense)', () => {
  beforeEach(cleanupModal);
  afterEach(cleanupModal);

  it('renders a literal "<img>" string in the title — does NOT create an <img> element', () => {
    showModal('<img src=x onerror=alert(1)>', 'plain message', 'info');

    const modal = document.getElementById('custom-modal');
    expect(modal).not.toBeNull();
    // No <img> anywhere in the modal subtree.
    expect(modal.querySelector('img')).toBeNull();
    // The literal string IS rendered, just as text.
    expect(modal.textContent).toContain('<img src=x onerror=alert(1)>');
  });

  it('renders a literal "<script>" string in the message — does NOT create a <script> element', () => {
    showModal('Title', '<script>alert(1)</script>', 'warning');

    const modal = document.getElementById('custom-modal');
    expect(modal.querySelector('script')).toBeNull();
    expect(modal.textContent).toContain('<script>alert(1)</script>');
  });

  it('renders normal text correctly', () => {
    showModal('Heads up', 'Please review your tasks.', 'info');

    const modal = document.getElementById('custom-modal');
    expect(modal.textContent).toContain('Heads up');
    expect(modal.textContent).toContain('Please review your tasks.');
  });

  it('has dialog ARIA attributes', () => {
    showModal('Title', 'Body', 'info');

    const modal = document.getElementById('custom-modal');
    expect(modal.getAttribute('role')).toBe('dialog');
    expect(modal.getAttribute('aria-modal')).toBe('true');
    expect(modal.getAttribute('aria-labelledby')).toBe('custom-modal-title');
    expect(document.getElementById('custom-modal-title')).not.toBeNull();
  });

  it('renders an OK button that closes the modal on click', () => {
    showModal('Title', 'Body', 'info');

    const okBtn = document.getElementById('modal-ok-btn');
    expect(okBtn).not.toBeNull();
    okBtn.click();
    expect(document.getElementById('custom-modal')).toBeNull();
  });

  it('replaces (does not duplicate) an existing modal when called twice', () => {
    showModal('First', 'First body', 'info');
    showModal('Second', 'Second body', 'warning');

    expect(document.querySelectorAll('#custom-modal').length).toBe(1);
    expect(document.getElementById('custom-modal').textContent).toContain('Second');
    expect(document.getElementById('custom-modal').textContent).not.toContain('First');
  });
});

describe('showWarningPopup / showInfoPopup (still safe through showModal)', () => {
  beforeEach(cleanupModal);
  afterEach(cleanupModal);

  it('showWarningPopup with HTML payload renders as text', () => {
    showWarningPopup('<svg onload=alert(1)>');
    const modal = document.getElementById('custom-modal');
    expect(modal.querySelector('svg')).toBeNull();
    expect(modal.textContent).toContain('<svg onload=alert(1)>');
  });

  it('showInfoPopup with HTML payload renders as text', () => {
    showInfoPopup('Info', '<a href="javascript:alert(1)">click</a>');
    const modal = document.getElementById('custom-modal');
    expect(modal.querySelector('a')).toBeNull();
    expect(modal.textContent).toContain('<a href="javascript:alert(1)">click</a>');
  });
});
