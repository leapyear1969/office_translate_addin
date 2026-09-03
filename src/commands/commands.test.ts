type CommandHandler = (event: { completed(): void }) => void;

interface OfficeMockOptions {
  originalHtml?: string;
  includeDisplayedBody?: boolean;
}

function installOfficeMock(options: OfficeMockOptions = {}) {
  const handlers = new Map<string, CommandHandler>();
  const setAsync = jest.fn(
    (
      _html: string,
      _options: unknown,
      callback: (result: { status: string }) => void,
    ) => callback({ status: "succeeded" }),
  );
  const replaceAsync = jest.fn(
    (
      _key: string,
      _message: unknown,
      callback?: (result: { status: string }) => void,
    ) => callback?.({ status: "succeeded" }),
  );
  const item: Record<string, unknown> = {
    itemType: "message",
    body: {
      getAsync: jest.fn(
        (
          _coercionType: string,
          callback: (result: { status: string; value: string }) => void,
        ) =>
          callback({
            status: "succeeded",
            value: options.originalHtml ?? "<p>Original</p>",
          }),
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
        item,
      },
    },
    onReady: (callback: () => void) => callback(),
    actions: {
      associate: (name: string, handler: CommandHandler) => handlers.set(name, handler),
    },
  };

  return { handlers, item, replaceAsync, setAsync };
}

function invoke(handler: CommandHandler): Promise<void> {
  return new Promise((resolve) => handler({ completed: resolve }));
}

describe("Outlook commands", () => {
  beforeEach(() => {
    jest.resetModules();
  });

  afterEach(() => {
    delete (globalThis as any).Office;
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
});
