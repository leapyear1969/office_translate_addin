# Outlook DisplayedBody Preview API Test Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a minimal Outlook Office.js add-in that reads a received message body, detects the preview `DisplayedBody.setAsync` API, temporarily replaces the displayed body, and restores the cached original HTML.

**Architecture:** An XML add-in-only manifest exposes two `ExecuteFunction` buttons on `MessageReadCommandSurface`. A UI-less commands page loads Global Office.js Beta and a TypeScript bundle; the bundle keeps original HTML in runtime memory, wraps Office callbacks in promises, performs structural capability checks, and guarantees command completion through `finally`.

**Tech Stack:** TypeScript, Office.js Preview types, HTML, XML Office Add-in manifest, webpack 5, webpack-dev-server, Jest with ts-jest, office-addin-dev-certs, office-addin-manifest, Node.js/npm.

---

## File map

- `package.json`: scripts and pinned development dependencies.
- `tsconfig.json`: strict TypeScript compilation with DOM and Office preview types.
- `webpack.config.js`: bundle/copy pipeline and trusted local HTTPS dev server.
- `jest.config.js`: Node test environment and TypeScript transform.
- `src/commands/commands.html`: UI-less command host that loads Global Office.js Beta.
- `src/commands/commands.ts`: command association and Outlook interactions.
- `src/commands/command-logic.ts`: Office-independent capability and error formatting helpers.
- `src/commands/command-logic.test.ts`: unit tests for helpers and command-completion wrapper.
- `src/assets/icon-16.png`, `icon-32.png`, `icon-80.png`: manifest command icons.
- `manifest.xml`: Message Read commands and local HTTPS resources.
- `README.md`: setup, HTTPS, CDN, sideloading, debugging, and result interpretation.
- `TEST_CHECKLIST.md`: executable manual checklist and feedback template.

### Task 1: Toolchain and failing logic tests

**Files:**
- Create: `package.json`
- Create: `tsconfig.json`
- Create: `jest.config.js`
- Create: `src/commands/command-logic.test.ts`

- [ ] **Step 1: Create the npm project and strict TypeScript configuration**

Define scripts `build`, `dev-server`, `test`, `typecheck`, and `validate`, plus exact dependencies for webpack, Jest, Office preview types, certificates, and manifest validation. Configure `target: ES2019`, `module: commonjs`, `strict: true`, `types: ["office-js-preview", "jest"]`, and exclude `dist`.

- [ ] **Step 2: Write failing helper tests**

Create tests covering all structural branches:

```typescript
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
  expect(formatAsyncError({ code: 9001, name: "TestError", message: "failed" }))
    .toBe("code=9001, name=TestError, message=failed");
});

it("completes a command once even when work rejects", async () => {
  const event = { completed: jest.fn() };
  await expect(completeCommand(event, async () => { throw new Error("boom"); }))
    .rejects.toThrow("boom");
  expect(event.completed).toHaveBeenCalledTimes(1);
});
```

- [ ] **Step 3: Run tests and verify the missing module failure**

Run: `npm install && npm test -- --runInBand`

Expected: Jest fails because `src/commands/command-logic.ts` does not exist.

- [ ] **Step 4: Commit the failing tests and toolchain**

```bash
git add package.json package-lock.json tsconfig.json jest.config.js src/commands/command-logic.test.ts
git commit -m "test: define displayed body command behavior"
```

### Task 2: Pure command helpers

**Files:**
- Create: `src/commands/command-logic.ts`
- Test: `src/commands/command-logic.test.ts`

- [ ] **Step 1: Implement the smallest typed helpers**

```typescript
export interface CommandEventLike {
  completed(): void;
}

export interface AsyncErrorLike {
  code?: string | number;
  name?: string;
  message?: string;
}

export interface DisplayedBodyLike {
  setAsync: (...args: unknown[]) => void;
}

export function getDisplayedBody(item: unknown): DisplayedBodyLike | null {
  const candidate = item as {
    display?: { body?: { setAsync?: unknown } };
  } | null | undefined;
  const body = candidate?.display?.body;
  return body && typeof body.setAsync === "function"
    ? body as DisplayedBodyLike
    : null;
}

export function formatAsyncError(error: AsyncErrorLike | null | undefined): string {
  return `code=${String(error?.code ?? "unknown")}, name=${error?.name ?? "unknown"}, message=${error?.message ?? "unknown"}`;
}

export async function completeCommand(
  event: CommandEventLike,
  work: () => Promise<void>,
): Promise<void> {
  try {
    await work();
  } finally {
    event.completed();
  }
}
```

- [ ] **Step 2: Run unit tests**

Run: `npm test -- --runInBand`

Expected: all helper tests pass.

- [ ] **Step 3: Run type checking**

Run: `npm run typecheck`

Expected: TypeScript exits with code 0.

- [ ] **Step 4: Commit the helpers**

```bash
git add src/commands/command-logic.ts
git commit -m "feat: add displayed body command helpers"
```

### Task 3: Office command runtime

**Files:**
- Create: `src/commands/commands.ts`
- Create: `src/commands/commands.html`

- [ ] **Step 1: Add the UI-less Global Beta command host**

Create valid HTML whose body is empty and whose head contains:

```html
<script src="https://appsforoffice.microsoft.com/lib/beta/hosted/office.js"></script>
```

Webpack injects the local `commands.js` bundle after the Office.js script.

- [ ] **Step 2: Implement environment diagnostics and notifications**

Add `dumpEnvironmentInfo()` using guarded access to host, platform, diagnostics, item type, `display`, and `display.body`. Add `showNotification(message, isError)` backed by `item.notificationMessages.replaceAsync`, with console fallback and full `AsyncResult.error` logging.

- [ ] **Step 3: Implement callback-to-promise Office operations**

Implement these exact operations:

```typescript
function getCurrentBodyHtml(item: Office.MessageRead): Promise<string>;
function setDisplayedBodyHtml(displayedBody: DisplayedBodyLike, html: string): Promise<void>;
```

Each callback compares status with `Office.AsyncResultStatus.Succeeded`; otherwise it logs `code`, `name`, and `message` via `formatAsyncError` and rejects.

- [ ] **Step 4: Implement `testTranslation`**

The command must call `dumpEnvironmentInfo`, read HTML using `item.body.getAsync(Office.CoercionType.Html, ...)`, save it to module-level `originalHtml`, then inspect and log:

```typescript
console.log("Office.context.mailbox.item", item);
console.log("item.display", itemAny.display);
console.log("item.display?.body", itemAny.display?.body);
console.log("typeof item.display?.body?.setAsync", typeof itemAny.display?.body?.setAsync);
```

When supported, set the requested fixed success HTML with `coercionType: Office.CoercionType.Html`. When unsupported, show exactly `当前 Outlook / Office.js 环境未提供 DisplayedBody.setAsync。`.

- [ ] **Step 5: Implement `restoreOriginal`**

If `originalHtml === null`, show exactly `尚未保存原始正文，请先点击“测试原文翻译”。`; otherwise repeat capability detection and set the cached HTML.

- [ ] **Step 6: Associate both functions after Office initialization**

```typescript
Office.onReady(() => {
  Office.actions.associate("testTranslation", testTranslation);
  Office.actions.associate("restoreOriginal", restoreOriginal);
});
```

Both functions call `completeCommand`; their work catches errors only to log and notify, while `completeCommand` guarantees exactly one `event.completed()` call.

- [ ] **Step 7: Run tests and type checking**

Run: `npm test -- --runInBand && npm run typecheck`

Expected: all tests pass and TypeScript exits with code 0.

- [ ] **Step 8: Commit the runtime**

```bash
git add src/commands/commands.ts src/commands/commands.html
git commit -m "feat: implement Outlook displayed body commands"
```

### Task 4: Webpack, manifest, and assets

**Files:**
- Create: `webpack.config.js`
- Create: `manifest.xml`
- Create: `src/assets/icon-16.png`
- Create: `src/assets/icon-32.png`
- Create: `src/assets/icon-80.png`

- [ ] **Step 1: Configure production and HTTPS development builds**

Use `ts-loader`, `HtmlWebpackPlugin` with `inject: "body"`, and `CopyWebpackPlugin` for assets. Obtain trusted certificate options with `office-addin-dev-certs.getHttpsServerOptions()` and configure port 3000, CORS headers, and `https` server mode.

- [ ] **Step 2: Generate minimal valid PNG icons**

Create opaque square PNGs at exactly 16×16, 32×32, and 80×80 pixels, using one restrained blue fill and a white `T` glyph or geometric mark.

- [ ] **Step 3: Create the XML manifest**

Use `MailApp` schema 1.1, `ReadWriteItem`, a stable minimum requirement `Mailbox 1.3`, Message/Read activation, `VersionOverridesV1_0`, `MessageReadCommandSurface`, one group, two `ExecuteFunction` buttons, `FunctionFile resid="Commands.Url"`, localhost command/icon resources, and concise Chinese labels.

- [ ] **Step 4: Build the add-in**

Run: `npm run build`

Expected: webpack exits successfully and produces `dist/commands.html`, `dist/commands.js`, and all three icons.

- [ ] **Step 5: Validate the manifest**

Run: `npm run validate`

Expected: `office-addin-manifest validate manifest.xml` reports the manifest is valid. If the validator reports environment-only store-policy warnings, record them separately; schema errors must be fixed.

- [ ] **Step 6: Verify the Global Beta reference survived the build**

Run: `rg -F "https://appsforoffice.microsoft.com/lib/beta/hosted/office.js" dist/commands.html`

Expected: exactly one matching script URL.

- [ ] **Step 7: Commit build configuration, manifest, and icons**

```bash
git add webpack.config.js manifest.xml src/assets
git commit -m "build: add Outlook manifest and HTTPS bundle"
```

### Task 5: Operator documentation and final verification

**Files:**
- Create: `README.md`
- Create: `TEST_CHECKLIST.md`
- Modify: `.gitignore`

- [ ] **Step 1: Write README setup and architecture sections**

Document `npm install`, `npx office-addin-dev-certs install`, `npm run dev-server`, `https://localhost:3000/commands.html`, the manifest `SourceLocation`/`FunctionFile`/Commands URL mapping, and why no task pane exists.

- [ ] **Step 2: Document all three Office.js URLs and the chosen test URL**

List Global stable, Global Beta, and 21Vianet stable URLs. State explicitly that Microsoft documents no 21Vianet Beta endpoint, this Demo uses Global Beta, and inability to load that CDN is distinct from the API object being absent after successful initialization.

- [ ] **Step 3: Document sideloading and client-specific caveats**

Use current Microsoft documentation links and distinguish classic Outlook on Windows, new Outlook on Windows, and Outlook on the web. Explain that sovereign-cloud administrators may need centralized deployment and network allowlisting, and avoid claiming that every tenant exposes end-user sideload UI.

- [ ] **Step 4: Write the exact manual test and result template**

Include the seven requested steps plus fields for client, version, Office.js URL, `item.display`, `item.display.body`, `setAsync`, error code, and error message.

- [ ] **Step 5: Create `TEST_CHECKLIST.md`**

Provide checkboxes for prerequisites, certificate trust, server response, manifest installation, read-mode activation, replacement, restoration, console capture, and the final pass/fail classification.

- [ ] **Step 6: Ignore generated files without touching user-owned `.local`**

Add only:

```gitignore
node_modules/
dist/
.DS_Store
npm-debug.log*
```

- [ ] **Step 7: Run the complete verification suite**

Run: `npm test -- --runInBand && npm run typecheck && npm run build && npm run validate`

Expected: unit tests, strict type check, webpack build, and manifest schema validation all pass.

- [ ] **Step 8: Inspect final scope and commit**

Run: `git status --short` and verify `.local` remains untouched and untracked.

```bash
git add README.md TEST_CHECKLIST.md .gitignore
git commit -m "docs: add 21Vianet Outlook test guide"
```

