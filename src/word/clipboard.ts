/** Call directly from a click handler to preserve browser user activation. */
export async function copyTranslation(field: HTMLTextAreaElement): Promise<boolean> {
  const text = field.value;
  const active = document.activeElement;
  const start = field.selectionStart;
  const end = field.selectionEnd;
  const direction = field.selectionDirection;
  const scrollTop = field.scrollTop;
  const scrollLeft = field.scrollLeft;
  const restore = () => {
    field.setSelectionRange(start, end, direction);
    field.scrollTop = scrollTop;
    field.scrollLeft = scrollLeft;
    if (active instanceof HTMLElement) active.focus({ preventScroll: true });
  };

  // Office Online embeds the pane in an iframe that may deny clipboard-write.
  // Try the selection-based command synchronously, before an await can lose
  // the user gesture. Keep this deprecated API only for host compatibility.
  field.focus({ preventScroll: true });
  field.select();
  let copied = false;
  try { copied = document.execCommand('copy'); } catch { /* Try Clipboard API. */ }
  if (copied) { restore(); return true; }

  try {
    await navigator.clipboard.writeText(text);
    restore();
    return true;
  } catch {
    // Leave the visible translation selected for keyboard or context-menu copy.
    field.focus({ preventScroll: true });
    field.select();
    return false;
  }
}
