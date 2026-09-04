# Mail translation implementation plan

**Goal:** Translate the complete Outlook message in place, provide a two-item menu and a translation preferences pane, authenticate the current user with the new China-cloud Entra app.

**Architecture:** Office command and task pane share settings and translation orchestration. An Express backend validates Office SSO tokens, uses MSAL OBO for Graph User.Read, and calls Azure China Translator. Secrets live only in the server .env. HTML text nodes are translated in bounded batches without translating attributes, CSS, URLs or embedded images. Results are applied only when the original message is still selected.

**Tech stack:** TypeScript, Office.js preview, Webpack, Express, MSAL Node, jose, Cheerio, Jest.

- [x] Add tests for HTML preservation, batching, failures, token validation, protected API routes, settings and message-switch races; run the failing tests.
- [x] Implement backend configuration, JWT validation, OBO profile retrieval, translation and local HTTPS hosting.
- [x] Implement shared settings and authenticated frontend requests; replace test commands with translation.
- [x] Implement the reference-style preferences pane, save/load, language management, account display and open-pane automatic behavior.
- [x] Replace manifest controls with a two-item translation menu; add v1.1 SSO registration and task pane pinning.
- [x] Run tests, typecheck, build and manifest validation; inspect the pane in a browser and test Azure using non-private sample content.
- [x] Document local startup, Entra prerequisites, the limited scope of open-pane automation and server migration. Keep user-owned .local unchanged and untracked.

Local SSO defaults: tenant d7125684-0e28-40b5-aba2-ea9580f2a201, client 59a13c0d-6f19-4c66-b8fd-2fa80d0186bc, origin https://localhost:3000. Application ID URI provisionally api://localhost:3000/59a13c0d-6f19-4c66-b8fd-2fa80d0186bc pending user confirmation. CLIENT_SECRET must be supplied by user. Manual translation always overrides automatic preferences.

Verification: 28 Jest tests pass; typecheck/build/manifest validation pass. Browser inspection confirms the reference-style settings layout and the explicit Outlook-only state outside Office. Tenant metadata responds HTTP 200. Application ID URI was confirmed by user; CLIENT_SECRET is now present in .env. After the user set TRANSLATOR_REGION=chinanorth3, real language detection returned en with confidence 1 and real translation returned Chinese while preserving table styles, links and CID image references. The earlier 401001 error is resolved. No live Outlook SSO/Graph end-to-end success is claimed. Final frontend bundle scan found no configured secrets.
