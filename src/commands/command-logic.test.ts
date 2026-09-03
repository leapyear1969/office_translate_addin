import {
  completeCommand,
  formatAsyncError,
  getDisplayedBody,
} from "./command-logic";

describe("getDisplayedBody", () => {
  it("returns null when the item does not expose display.body.setAsync", () => {
    expect(getDisplayedBody(undefined)).toBeNull();
    expect(getDisplayedBody({ display: {} })).toBeNull();
    expect(getDisplayedBody({ display: { body: {} } })).toBeNull();
  });

  it("returns the body object when setAsync is callable", () => {
    const body = { setAsync: jest.fn() };

    expect(getDisplayedBody({ display: { body } })).toBe(body);
  });
});

it("formats every Office async error field", () => {
  expect(
    formatAsyncError({ code: 9001, name: "TestError", message: "failed" }),
  ).toBe("code=9001, name=TestError, message=failed");
});

it("formats missing Office async error fields safely", () => {
  expect(formatAsyncError(undefined)).toBe(
    "code=unknown, name=unknown, message=unknown",
  );
});

it("completes a command once after successful work", async () => {
  const event = { completed: jest.fn() };

  await completeCommand(event, async () => undefined);

  expect(event.completed).toHaveBeenCalledTimes(1);
});

it("completes a command once even when work rejects", async () => {
  const event = { completed: jest.fn() };

  await expect(
    completeCommand(event, async () => {
      throw new Error("boom");
    }),
  ).rejects.toThrow("boom");
  expect(event.completed).toHaveBeenCalledTimes(1);
});
