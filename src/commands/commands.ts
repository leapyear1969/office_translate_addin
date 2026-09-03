import {
  completeCommand,
  DisplayedBodyLike,
  formatAsyncError,
  getDisplayedBody,
} from "./command-logic";

const STATUS_NOTIFICATION_KEY = "displayed-body-test-status";
const ORIGINAL_BODY_STORAGE_PREFIX = "displayed-body-test:original:";
const ORIGINAL_BODY_CACHE_TTL_MS = 60 * 60 * 1000;
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
let originalItemId: string | null = null;
let originalItemReference: Office.MessageRead | null = null;

function getItem(): Office.MessageRead | null {
  return (Office.context?.mailbox?.item as Office.MessageRead | undefined) ?? null;
}

function getItemId(item: Office.MessageRead): string | null {
  const itemId = (item as any).itemId;
  return typeof itemId === "string" && itemId.length > 0 ? itemId : null;
}

function getOriginalBodyStorageKey(item: Office.MessageRead): string | null {
  const itemId = getItemId(item);
  return itemId ? `${ORIGINAL_BODY_STORAGE_PREFIX}${itemId}` : null;
}

function cacheOriginalBody(item: Office.MessageRead, html: string): void {
  const key = getOriginalBodyStorageKey(item);
  if (!key || typeof localStorage === "undefined") {
    return;
  }

  try {
    localStorage.setItem(key, JSON.stringify({ html, savedAt: Date.now() }));
  } catch (error) {
    console.warn("Unable to cache the original body in localStorage:", error);
  }
}

function loadCachedOriginalBody(item: Office.MessageRead): string | null {
  const key = getOriginalBodyStorageKey(item);
  if (!key || typeof localStorage === "undefined") {
    return null;
  }

  try {
    const stored = localStorage.getItem(key);
    if (stored === null) {
      return null;
    }

    const parsed = JSON.parse(stored) as { html?: unknown; savedAt?: unknown };
    if (
      typeof parsed.html !== "string" ||
      typeof parsed.savedAt !== "number" ||
      Date.now() - parsed.savedAt > ORIGINAL_BODY_CACHE_TTL_MS
    ) {
      localStorage.removeItem(key);
      return null;
    }

    return parsed.html;
  } catch (error) {
    console.warn("Unable to read the original body from localStorage:", error);
    try {
      localStorage.removeItem(key);
    } catch (removeError) {
      console.warn("Unable to remove invalid original body cache:", removeError);
    }
    return null;
  }
}

function removeCachedOriginalBody(item: Office.MessageRead): void {
  const key = getOriginalBodyStorageKey(item);
  if (!key || typeof localStorage === "undefined") {
    return;
  }

  try {
    localStorage.removeItem(key);
  } catch (error) {
    console.warn("Unable to remove the original body from localStorage:", error);
  }
}

function purgeExpiredOriginalBodyCaches(): void {
  if (typeof localStorage === "undefined") {
    return;
  }

  try {
    const keys: string[] = [];
    for (let index = 0; index < localStorage.length; index += 1) {
      const key = localStorage.key(index);
      if (key?.startsWith(ORIGINAL_BODY_STORAGE_PREFIX)) {
        keys.push(key);
      }
    }

    for (const key of keys) {
      try {
        const stored = localStorage.getItem(key);
        const parsed = stored
          ? (JSON.parse(stored) as { html?: unknown; savedAt?: unknown })
          : null;
        const isValid =
          typeof parsed?.html === "string" &&
          typeof parsed.savedAt === "number" &&
          Date.now() - parsed.savedAt <= ORIGINAL_BODY_CACHE_TTL_MS;

        if (!isValid) {
          localStorage.removeItem(key);
        }
      } catch (error) {
        console.warn("Unable to inspect an original body cache:", error);
        try {
          localStorage.removeItem(key);
        } catch (removeError) {
          console.warn("Unable to remove an invalid body cache:", removeError);
        }
      }
    }
  } catch (error) {
    console.warn("Unable to sweep original body caches:", error);
  }
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
  purgeExpiredOriginalBodyCaches();
  dumpEnvironmentInfo();
  const item = getItem();
  if (!item) {
    await showNotification("当前 Outlook 中没有可读取的邮件项目。");
    return;
  }

  originalHtml = await getCurrentBodyHtml(item);
  originalItemId = getItemId(item);
  originalItemReference = item;
  cacheOriginalBody(item, originalHtml);
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
  purgeExpiredOriginalBodyCaches();
  dumpEnvironmentInfo();
  const item = getItem();
  if (!item) {
    await showNotification("当前 Outlook 中没有可读取的邮件项目。");
    return;
  }

  const currentItemId = getItemId(item);
  const inMemoryBodyMatchesItem = currentItemId
    ? originalItemId === currentItemId
    : originalItemReference === item;

  if (!inMemoryBodyMatchesItem) {
    originalHtml = loadCachedOriginalBody(item);
    originalItemId = originalHtml === null ? null : currentItemId;
    originalItemReference = originalHtml === null ? null : item;
  }

  if (originalHtml === null) {
    await showNotification(NO_ORIGINAL_MESSAGE);
    return;
  }

  logDisplayedBodyCapability(item);
  const displayedBody = getDisplayedBody(item);
  if (!displayedBody) {
    await showNotification(UNSUPPORTED_MESSAGE);
    return;
  }

  await setDisplayedBodyHtml(displayedBody, originalHtml);
  removeCachedOriginalBody(item);
  originalHtml = null;
  originalItemId = null;
  originalItemReference = null;
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
