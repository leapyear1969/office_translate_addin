const MIN_WIDTH = 330;
const STORAGE_KEY = 'word-translator.pane-width';

export function rememberPaneWidth() {
  const web = Office.context.platform === Office.PlatformType.OfficeOnline;
  const normalize = (width: number) => Math.round(Math.max(MIN_WIDTH, web ? Math.min(500, width) : width));
  const resize = (width: number) => {
    try { Office.extensionLifeCycle?.taskpane?.setWidth(width); }
    catch { /* Older hosts may not support changing the outer pane width. */ }
  };
  let saved = 0;
  try { saved = Number(localStorage.getItem(STORAGE_KEY)); } catch { /* Storage is optional. */ }
  if (Number.isFinite(saved) && saved > 0) resize(normalize(saved));
  else if (window.innerWidth > 0 && window.innerWidth < MIN_WIDTH) resize(MIN_WIDTH);

  let timer: ReturnType<typeof setTimeout> | undefined;
  const onResize = () => {
    clearTimeout(timer);
    const width = window.innerWidth;
    if (document.hidden || width <= 0) return;
    // Correct the host immediately; debounce only persistence, not the limit.
    if (width < MIN_WIDTH) { resize(MIN_WIDTH); return; }
    timer = setTimeout(() => {
      if (document.hidden) return;
      try { localStorage.setItem(STORAGE_KEY, String(normalize(width))); } catch { /* Storage is optional. */ }
    }, 250);
  };
  window.addEventListener('resize', onResize);
  return () => {
    clearTimeout(timer);
    window.removeEventListener('resize', onResize);
  };
}
