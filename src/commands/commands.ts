import {
  completeCommand,
  DisplayedBodyLike,
  formatAsyncError,
  getDisplayedBody,
} from "./command-logic";

const STATUS_NOTIFICATION_KEY = "displayed-body-test-status";
const UNSUPPORTED_MESSAGE =
  "当前 Outlook / Office.js 环境未提供 DisplayedBody.setAsync。";
const NO_ORIGINAL_MESSAGE =
  "尚未保存原始正文，请先点击“测试原文翻译”。";
const TEST_HTML = `
<div style="padding:20px;font-family:Segoe UI,Microsoft YaHei,sans-serif;">
    <h2>翻译测试成功</h2>
    <p>如果你现在看到这段文字，说明当前 Outlook 支持 DisplayedBody.setAsync。</p>
    <hr>
    <p>原始邮件正文已经成功读取，但没有修改服务器中的原始邮件。</p>
</div>`;

let originalHtml: string | null = null;

function getItem(): Office.MessageRead | null {
  return (Office.context?.mailbox?.item as Office.MessageRead | undefined) ?? null;
}

export function dumpEnvironmentInfo(): void {
  const context = Office.context;
  const mailbox = context?.mailbox;
  const diagnostics = mailbox?.diagnostics;
  const item = mailbox?.item as any;

  console.log("Office.context.host", context?.host);
  console.log("Office.context.platform", context?.platform);
  console.log("Office.context.mailbox.diagnostics.hostName", diagnostics?.hostName);
  console.log(
    "Office.context.mailbox.diagnostics.hostVersion",
    diagnostics?.hostVersion,
  );
  console.log("Office.context.mailbox.diagnostics.OWAView", diagnostics?.OWAView);
  console.log("Office.context.mailbox.item.itemType", item?.itemType);
  console.log("Office.context.mailbox.item.display", item?.display);
  console.log("Office.context.mailbox.item.display?.body", item?.display?.body);
}

function logDisplayedBodyCapability(item: Office.MessageRead | null): void {
  const itemAny = item as any;

  console.log("Office.context.mailbox.item", item);
  console.log("item.display", itemAny?.display);
  console.log("item.display?.body", itemAny?.display?.body);
  console.log(
    "typeof item.display?.body?.setAsync",
    typeof itemAny?.display?.body?.setAsync,
  );
}

function showNotification(message: string): Promise<void> {
  console.error(message);

  const notificationMessages = (Office.context?.mailbox?.item as any)
    ?.notificationMessages;
  if (!notificationMessages?.replaceAsync) {
    return Promise.resolve();
  }

  return new Promise((resolve) => {
    notificationMessages.replaceAsync(
      STATUS_NOTIFICATION_KEY,
      {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: message.slice(0, 150),
      },
      (result: Office.AsyncResult<void>) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error(
            "notificationMessages.replaceAsync failed:",
            formatAsyncError(result.error),
          );
        }
        resolve();
      },
    );
  });
}

function getCurrentBodyHtml(item: Office.MessageRead): Promise<string> {
  return new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Html, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
        return;
      }

      const details = formatAsyncError(result.error);
      console.error("item.body.getAsync failed:", details);
      reject(new Error(`读取邮件正文失败：${details}`));
    });
  });
}

function setDisplayedBodyHtml(
  displayedBody: DisplayedBodyLike,
  html: string,
): Promise<void> {
  return new Promise((resolve, reject) => {
    displayedBody.setAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      (result: Office.AsyncResult<void>) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
          return;
        }

        const details = formatAsyncError(result.error);
        console.error("item.display.body.setAsync failed:", details);
        reject(new Error(`替换显示正文失败：${details}`));
      },
    );
  });
}

async function runTestTranslation(): Promise<void> {
  dumpEnvironmentInfo();
  const item = getItem();
  if (!item) {
    await showNotification("当前 Outlook 中没有可读取的邮件项目。");
    return;
  }

  originalHtml = await getCurrentBodyHtml(item);
  console.log("Original message body read successfully.");

  logDisplayedBodyCapability(item);
  const displayedBody = getDisplayedBody(item);
  if (!displayedBody) {
    await showNotification(UNSUPPORTED_MESSAGE);
    return;
  }

  await setDisplayedBodyHtml(displayedBody, TEST_HTML);
  console.log("DisplayedBody.setAsync succeeded; the displayed body was replaced.");
}

async function runRestoreOriginal(): Promise<void> {
  dumpEnvironmentInfo();
  if (originalHtml === null) {
    await showNotification(NO_ORIGINAL_MESSAGE);
    return;
  }

  const item = getItem();
  logDisplayedBodyCapability(item);
  const displayedBody = getDisplayedBody(item);
  if (!displayedBody) {
    await showNotification(UNSUPPORTED_MESSAGE);
    return;
  }

  await setDisplayedBodyHtml(displayedBody, originalHtml);
  console.log("DisplayedBody.setAsync succeeded; the original body was restored.");
}

function handleCommandError(error: unknown): Promise<void> {
  const message = error instanceof Error ? error.message : String(error);
  console.error("Command failed:", error);
  return showNotification(message);
}

export function testTranslation(event: Office.AddinCommands.Event): void {
  void completeCommand(event, async () => {
    try {
      await runTestTranslation();
    } catch (error) {
      await handleCommandError(error);
    }
  });
}

export function restoreOriginal(event: Office.AddinCommands.Event): void {
  void completeCommand(event, async () => {
    try {
      await runRestoreOriginal();
    } catch (error) {
      await handleCommandError(error);
    }
  });
}

Office.onReady(() => {
  Office.actions.associate("testTranslation", testTranslation);
  Office.actions.associate("restoreOriginal", restoreOriginal);
});
