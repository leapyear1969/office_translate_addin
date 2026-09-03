type CommandHandler = (event: { completed(): void }) => void;

interface OfficeMockOptions {
  originalHtml?: string;
  includeDisplayedBody?: boolean;
  itemId?: string;
  bodyStatus?: "succeeded" | "failed";
  displayStatus?: "succeeded" | "failed";
  itemPresent?: boolean;
  notificationStatus?: "succeeded" | "failed";
}

function installOfficeMock(options: OfficeMockOptions = {}) {
  const handlers = new Map<string, CommandHandler>();
  const setAsync = jest.fn(
    (
      _html: string,
      _options: unknown,
      callback: (result: {
        status: string;
        error?: { code: number; name: string; message: string };
      }) => void,
    ) =>
      callback(
        options.displayStatus === "failed"
          ? {
              status: "failed",
              error: {
                code: 5002,
                name: "DisplayFailed",
                message: "Cannot set displayed body",
              },
            }
          : { status: "succeeded" },
      ),
  );
  const replaceAsync = jest.fn(
    (
      _key: string,
      _message: unknown,
      callback?: (result: { status: string }) => void,
    ) =>
      callback?.(
        options.notificationStatus === "failed"
          ? ({
              status: "failed",
              error: {
                code: 5003,
                name: "NotificationFailed",
                message: "Cannot show notification",
              },
            } as any)
          : { status: "succeeded" },
      ),
  );
  const item: Record<string, unknown> = {
    itemId: options.itemId ?? "message-id-1",
    itemType: "message",
    body: {
      getAsync: jest.fn(
        (
          _coercionType: string,
          callback: (result: { status: string; value: string }) => void,
        ) => {
          if (options.bodyStatus === "failed") {
            callback({
              status: "failed",
              value: "",
              error: {
                code: 5001,
                name: "BodyReadFailed",
                message: "Cannot read body",
              },
            } as any);
            return;
          }

          callback({
            status: "succeeded",
            value: options.originalHtml ?? "<p>Original</p>",
          });
        },
      ),
    },
    notificationMessages: { replaceAsync },
  };

  if (options.includeDisplayedBody !== false) {
    item.display = { body: { setAsync } };
  }

  (globalThis as any).Office = {
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    CoercionType: { Html: "html" },
    MailboxEnums: {
      ItemNotificationMessageType: {
        InformationalMessage: "informationalMessage",
        ErrorMessage: "errorMessage",
      },
    },
    context: {
      host: "Outlook",
      platform: "PC",
      mailbox: {
        diagnostics: {
          hostName: "Outlook",
          hostVersion: "16.0",
          OWAView: "OneColumn",
        },
        item: options.itemPresent === false ? undefined : item,
      },
    },
    onReady: (callback: () => void) => callback(),
    actions: {
      associate: (name: string, handler: CommandHandler) => handlers.set(name, handler),
    },
  };

  return { handlers, item, replaceAsync, setAsync };
}

function invoke(handler: CommandHandler): Promise<jest.Mock> {
  return new Promise((resolve) => {
    let completed: jest.Mock;
    completed = jest.fn((): void => resolve(completed));
    handler({ completed });
  });
}

describe("Outlook commands", () => {
  beforeEach(() => {
    jest.resetModules();
    const values = new Map<string, string>();
    (globalThis as any).localStorage = {
      getItem: (key: string) => values.get(key) ?? null,
      setItem: (key: string, value: string) => values.set(key, value),
      removeItem: (key: string) => values.delete(key),
      key: (index: number) => Array.from(values.keys())[index] ?? null,
      get length() {
        return values.size;
      },
    };
  });

  afterEach(() => {
    delete (globalThis as any).Office;
    delete (globalThis as any).localStorage;
    jest.restoreAllMocks();
  });

  it("reads the original HTML and replaces the displayed body", async () => {
    const { handlers, setAsync } = installOfficeMock({
      originalHtml: "<p>Mail body</p>",
    });
    jest.spyOn(console, "log").mockImplementation(() => undefined);
    await import("./commands");

    await invoke(handlers.get("testTranslation")!);

    expect(setAsync).toHaveBeenCalledTimes(1);
    expect(setAsync.mock.calls[0][0]).toContain("翻译测试成功");
    expect(setAsync.mock.calls[0][1]).toEqual({ coercionType: "html" });
  });

  it("restores the HTML cached by the translation test", async () => {
    const { handlers, setAsync } = installOfficeMock({
      originalHtml: "<p>Saved original</p>",
    });
    jest.spyOn(console, "log").mockImplementation(() => undefined);
    await import("./commands");

    await invoke(handlers.get("testTranslation")!);
    await invoke(handlers.get("restoreOriginal")!);

    expect(setAsync).toHaveBeenNthCalledWith(
      2,
      "<p>Saved original</p>",
      { coercionType: "html" },
      expect.any(Function),
    );
  });

  it("restores cached HTML after Outlook creates a new command runtime", async () => {
    const firstRuntime = installOfficeMock({
      originalHtml: "<p>Cross-runtime original</p>",
    });
    jest.spyOn(console, "log").mockImplementation(() => undefined);
    await import("./commands");
    await invoke(firstRuntime.handlers.get("testTranslation")!);

    jest.resetModules();
    const secondRuntime = installOfficeMock();
    await import("./commands");
    await invoke(secondRuntime.handlers.get("restoreOriginal")!);

    expect(secondRuntime.setAsync).toHaveBeenCalledWith(
      "<p>Cross-runtime original</p>",
      { coercionType: "html" },
      expect.any(Function),
    );
  });

  it("never restores one message body into another message in the same runtime", async () => {
    const firstMessage = installOfficeMock({
      itemId: "message-a",
      originalHtml: "<p>Private body A</p>",
    });
    jest.spyOn(console, "log").mockImplementation(() => undefined);
    jest.spyOn(console, "error").mockImplementation(() => undefined);
    await import("./commands");
    await invoke(firstMessage.handlers.get("testTranslation")!);

    const secondMessage = installOfficeMock({ itemId: "message-b" });
    const completed = await invoke(firstMessage.handlers.get("restoreOriginal")!);

    expect(secondMessage.setAsync).not.toHaveBeenCalled();
    expect(secondMessage.replaceAsync).toHaveBeenCalledWith(
      "displayed-body-test-status",
      expect.objectContaining({
        message: "尚未保存原始正文，请先点击“测试原文翻译”。",
      }),
      expect.any(Function),
    );
    expect(completed).toHaveBeenCalledTimes(1);
  });

  it("deletes the client cache after restoring the original body", async () => {
    const { handlers } = installOfficeMock({
      itemId: "message-cleanup",
      originalHtml: "<p>Delete after restore</p>",
    });
    jest.spyOn(console, "log").mockImplementation(() => undefined);
    await import("./commands");

    await invoke(handlers.get("testTranslation")!);
    expect(localStorage.getItem("displayed-body-test:original:message-cleanup"))
      .not.toBeNull();
    await invoke(handlers.get("restoreOriginal")!);

    expect(localStorage.getItem("displayed-body-test:original:message-cleanup"))
      .toBeNull();
  });

  it("sweeps expired body caches for other messages when a command runs", async () => {
    const prefix = "displayed-body-test:original:";
    localStorage.setItem(
      `${prefix}expired-message`,
      JSON.stringify({
        html: "<p>Expired private body</p>",
        savedAt: Date.now() - 61 * 60 * 1000,
      }),
    );
    localStorage.setItem(
      `${prefix}valid-message`,
      JSON.stringify({ html: "<p>Still valid</p>", savedAt: Date.now() }),
    );
    const { handlers } = installOfficeMock({ itemId: "current-message" });
    jest.spyOn(console, "log").mockImplementation(() => undefined);
    await import("./commands");

    await invoke(handlers.get("testTranslation")!);

    expect(localStorage.getItem(`${prefix}expired-message`)).toBeNull();
    expect(localStorage.getItem(`${prefix}valid-message`)).not.toBeNull();
  });

  it("shows a clear notification when DisplayedBody.setAsync is unavailable", async () => {
    const { handlers, replaceAsync } = installOfficeMock({
      includeDisplayedBody: false,
    });
    jest.spyOn(console, "log").mockImplementation(() => undefined);
    jest.spyOn(console, "error").mockImplementation(() => undefined);
    await import("./commands");

    await invoke(handlers.get("testTranslation")!);

    expect(replaceAsync).toHaveBeenCalledWith(
      "displayed-body-test-status",
      expect.objectContaining({
        message: "当前 Outlook / Office.js 环境未提供 DisplayedBody.setAsync。",
      }),
      expect.any(Function),
    );
  });

  it("asks the user to run the translation test before restoring", async () => {
    const { handlers, replaceAsync, setAsync } = installOfficeMock();
    jest.spyOn(console, "log").mockImplementation(() => undefined);
    jest.spyOn(console, "error").mockImplementation(() => undefined);
    await import("./commands");

    await invoke(handlers.get("restoreOriginal")!);

    expect(setAsync).not.toHaveBeenCalled();
    expect(replaceAsync).toHaveBeenCalledWith(
      "displayed-body-test-status",
      expect.objectContaining({
        message: "尚未保存原始正文，请先点击“测试原文翻译”。",
      }),
      expect.any(Function),
    );
  });

  it("logs body read failures and completes the command exactly once", async () => {
    const { handlers, replaceAsync, setAsync } = installOfficeMock({
      bodyStatus: "failed",
    });
    jest.spyOn(console, "log").mockImplementation(() => undefined);
    jest.spyOn(console, "error").mockImplementation(() => undefined);
    await import("./commands");

    const completed = await invoke(handlers.get("testTranslation")!);

    expect(setAsync).not.toHaveBeenCalled();
    expect(replaceAsync).toHaveBeenCalledWith(
      "displayed-body-test-status",
      expect.objectContaining({
        message: expect.stringContaining("code=5001"),
      }),
      expect.any(Function),
    );
    expect(completed).toHaveBeenCalledTimes(1);
  });

  it("logs displayed-body failures and completes the command exactly once", async () => {
    const { handlers, replaceAsync } = installOfficeMock({
      displayStatus: "failed",
    });
    jest.spyOn(console, "log").mockImplementation(() => undefined);
    jest.spyOn(console, "error").mockImplementation(() => undefined);
    await import("./commands");

    const completed = await invoke(handlers.get("testTranslation")!);

    expect(replaceAsync).toHaveBeenCalledWith(
      "displayed-body-test-status",
      expect.objectContaining({
        message: expect.stringContaining("code=5002"),
      }),
      expect.any(Function),
    );
    expect(completed).toHaveBeenCalledTimes(1);
  });

  it("completes exactly once when no current item exists", async () => {
    const { handlers, setAsync } = installOfficeMock({ itemPresent: false });
    jest.spyOn(console, "log").mockImplementation(() => undefined);
    jest.spyOn(console, "error").mockImplementation(() => undefined);
    await import("./commands");

    const completed = await invoke(handlers.get("testTranslation")!);

    expect(setAsync).not.toHaveBeenCalled();
    expect(completed).toHaveBeenCalledTimes(1);
  });

  it("logs notification callback failures and still completes exactly once", async () => {
    const { handlers } = installOfficeMock({
      includeDisplayedBody: false,
      notificationStatus: "failed",
    });
    jest.spyOn(console, "log").mockImplementation(() => undefined);
    const errorLog = jest
      .spyOn(console, "error")
      .mockImplementation(() => undefined);
    await import("./commands");

    const completed = await invoke(handlers.get("testTranslation")!);

    expect(errorLog).toHaveBeenCalledWith(
      "notificationMessages.replaceAsync failed:",
      "code=5003, name=NotificationFailed, message=Cannot show notification",
    );
    expect(completed).toHaveBeenCalledTimes(1);
  });

  it("treats an unreadable client cache as missing and completes once", async () => {
    const { handlers, replaceAsync } = installOfficeMock({ itemId: "broken-cache" });
    (globalThis as any).localStorage = {
      getItem: () => {
        throw new Error("read blocked");
      },
      setItem: jest.fn(),
      removeItem: () => {
        throw new Error("remove blocked");
      },
    };
    jest.spyOn(console, "log").mockImplementation(() => undefined);
    jest.spyOn(console, "error").mockImplementation(() => undefined);
    jest.spyOn(console, "warn").mockImplementation(() => undefined);
    await import("./commands");

    const completed = await invoke(handlers.get("restoreOriginal")!);

    expect(replaceAsync).toHaveBeenCalledWith(
      "displayed-body-test-status",
      expect.objectContaining({ message: NO_ORIGINAL_MESSAGE_FOR_TEST }),
      expect.any(Function),
    );
    expect(completed).toHaveBeenCalledTimes(1);
  });
});

const NO_ORIGINAL_MESSAGE_FOR_TEST =
  "尚未保存原始正文，请先点击“测试原文翻译”。";
