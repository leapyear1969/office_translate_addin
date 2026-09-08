# Outlook Notification Toggle Implementation Plan

> **For agentic workers:** Execute inline in this session using test-first steps. The user has approved the single-action design; do not modify unrelated Word/server changes. Commit/push only when explicitly requested, and do not deploy.

**Goal:** Switch the mail-body notification between restoring the original and translating it to the saved target language.

**Architecture:** Reuse the existing notification key and `Translation.Options` command. Share the single-action notification/fallback implementation, dispatch both context actions through existing pane operations, and suppress automatic inspection after an explicit startup action.

**Tech Stack:** TypeScript, Office.js Mailbox APIs, Jest/ts-jest/jsdom, Webpack.

---

### Task 1: Specify and test the notification cycle

**Files:** `src/commands/commands.test.ts`, `src/taskpane/taskpane.test.ts`.

- [x] Extend the Office mock in command tests with `roamingSettings.get`, returning undefined for default settings.
- [x] Add restore success assertions against the existing notification key:
  ```ts
  expect(item.notificationMessages.replaceAsync).toHaveBeenLastCalledWith(
    'mail-translation-status', expect.objectContaining({
      message: '已显示原文。',
      actions: [{ actionText: '将邮件翻译为：中文（简体）', actionType: 'showTaskPane',
        commandId: 'Translation.Options', contextData: JSON.stringify({ action: 'translateMessage' }) }],
    }), expect.any(Function));
  ```
- [x] Test saved Japanese language, unsupported notification fallback, failed restoration, and selection changes during display completion.
- [x] Test `setup({ action: 'translateMessage' }, 'never', 'ja')` and JSON-string contexts: exactly one `/api/translate` with `to: 'ja'`, one display replacement, and one show-original action.
- [x] Test the same launch in `always` mode: no redundant `/api/detect` or `/api/translate` after explicit translation.
- [x] In a running pane, restore then feed its actual notification context back through `itemHandlers.contextChanged`; verify the resulting translation changes the action back to show-original. Repeat in the other direction.
- [x] Test saving target preferences after restoration, stale item notification handlers, in-flight duplicate clicks, and pending sign-in finishing after a notification action.
- [x] Run `npm test -- --runInBand src/commands/commands.test.ts src/taskpane/taskpane.test.ts`; verify failures identify the missing translate notification/dispatcher behavior.

### Task 2: Implement single-action notifications

**File:** `src/shared/mail.ts`.

- [x] Import `LANGUAGES` and `loadSettings` from `./settings`.
- [x] Generalize the existing translation-complete notifier to a private `notifyMessageAction(item, message, actionText, action, fallback)` helper. It checks the current item, writes exactly one action with `JSON.stringify({ action })`, and uses ordinary notification fallback on action rejection.
- [x] Keep the translation completion text and show-original payload unchanged.
- [x] Add `notifyOriginalDisplayed(item, target?)` with label `` `将邮件翻译为：${LANGUAGES[target || loadSettings().target]}` `` and action `translateMessage`. Its fallback says to use that action in translation options.
- [x] Call `notifyOriginalDisplayed(item, target)` only after successful original body replacement; the pane supplies the last committed target, not roaming settings being saved.
- [x] Run the command tests and ensure notification success/fallback and body-safety regressions pass.

### Task 3: Dispatch translate contexts without duplicate automatic work

**File:** `src/taskpane/taskpane.ts`.

- [x] Dispatch both supported actions through existing `restoreOriginal()` / `translateItem(currentItem())` functions; return a boolean indicating a recognized launch action. Preserve invalid-context handling.
- [x] Return that boolean from `readInitializationContext()`; use `false` when unsupported, unsuccessful, or stale. Capture the initial `actionEpoch`; if a newer notification operation has run, ignore the delayed startup context and return `true` to suppress automatic inspection.
- [x] Give `signIn` an optional inspection flag. Initialize with `readInitializationContext().then(handled => signIn(false, !handled))`, guarded by the startup `loginEpoch` so it cannot supersede a new item's login. Capture `actionEpoch` when sign-in starts and inspect only if no newer explicit action has superseded it.
- [x] After saving preferences, call `refreshOriginalNotification()`: read `notificationMessages.getAllAsync`, find `mail-translation-status`, and update only an existing restored-original notification. Check selected item, operation version and committed target before writing. Do not use a local original-display boolean: ribbon commands can change the display independently, and a user may dismiss the notification.
- [x] Pass the committed target into restoration. If settings commit while restoration is pending, update the final notification to the newly committed target after display succeeds. Test save failure rollback as well as successful concurrent save.
- [x] Run both focused test suites until all pass. Existing sign-out, consent, no-item, language settings, and async item-switch protections must remain intact.

### Task 4: Verify and document

**File:** `TEST_CHECKLIST.md`.

- [x] Add manual checks for the two directions, one action only, saved target labels, fallback, and fresh/pinned task panes without duplicate auto-translation.
- [x] Run `npm test -- --runInBand` and `npm run typecheck`.
- [x] Build with `npx webpack --mode production --env host=outlook --output-path <new temporary build directory>` to avoid cleaning/replacing a user's existing dist directory or regenerating manifests.
- [x] Review `git diff --check` and the changed Outlook files only; report unrelated pre-existing failures if any. Do not claim real Outlook validation without running it.

## Verification outcome

- Initial red run: 12 new cases failed for missing notification/dispatcher behavior; 38 passed.
- First implementation: 50 focused tests passed; 217 full-suite tests, global typecheck and a temporary production build passed.
- Review reproduced and fixed three edge cases: ribbon notifications being overwritten, uncommitted target language during failed saves, and delayed startup contexts replaying an operation. Added regression coverage also preserves dismissed notifications and new-item login.
- Final focused run: 57 tests passed. Independent review confirmed the three fixes.
- Final non-Word run: 14 suites / 124 tests passed. A TypeScript check using the project's compiler flags on the changed Outlook source/tests and their imports passed.
- Final whole-project run was blocked by concurrently changed Word files: `src/word/document.test.ts:440` expected cancellation after a layout-only edit, and `src/word/ooxml.test.ts:2` imported the not-yet-exported `ooxmlStructure`. Global `tsc --noEmit` failed on the latter; the chained final production rebuild was therefore skipped. Those files were not modified for this task.
- At implementation handoff, `git diff --check` passed with Git's LF-to-CRLF conversion warnings. No commit, deployment, manifest change, or real Outlook validation was performed at that stage.
- The user subsequently requested a GitHub push. After the separate Word work was committed, the pre-push recheck passed: 17 suites / 229 tests, global `tsc --noEmit`, and a temporary production Webpack build. The Outlook-only change is to be committed on `feat/outlook-notification-toggle`; backup files remain excluded. This does not deploy the add-in.
