import { completeCommand } from './command-logic';
import { currentItem, notify, translateCurrentMessage } from '../shared/mail';
import { loadSettings } from '../shared/settings';

export function translateMessage(event: Office.AddinCommands.Event): void {
  const item = currentItem();
  void completeCommand(event, async () => {
    try { await translateCurrentMessage(loadSettings().target); }
    catch (error) { if (item && (error as { code?: string }).code !== 'consent_required') notify(item, error instanceof Error ? error.message : '翻译失败，请重试。', true); }
  });
}
Office.onReady(() => { Office.actions.associate('translateMessage', translateMessage); });
