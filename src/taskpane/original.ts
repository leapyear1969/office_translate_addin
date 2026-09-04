import { showOriginalMessage } from '../shared/mail';

const status = document.getElementById('status')!;
const retry = document.getElementById('show-original') as HTMLButtonElement;
async function restore() {
  retry.disabled = true;
  status.textContent = '正在显示原文…';
  try {
    await showOriginalMessage();
    status.textContent = '已显示原文，可以关闭此面板。';
  } catch (error) {
    status.textContent = error instanceof Error ? error.message : '原文显示失败，请重试。';
  } finally { retry.disabled = false; }
}
retry.addEventListener('click', () => void restore());
Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) void restore();
});
