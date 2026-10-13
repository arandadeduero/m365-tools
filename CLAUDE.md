# m365-users

Node.js CLI that manages the Microsoft 365 users of one specific tenant (Ayuntamiento de Aranda de Duero, `arandadeduero.es`) through Microsoft Graph. It is run by hand by an administrator against the **production** tenant.

## Safety

- There is no sandbox tenant. Never run a CLI command that talks to Graph (`import`, `fix …`, `sync-check`, `reset-password`, `edit`, `list`, …) to "try something". Verify behaviour with tests.
- `config.json` (real tenant and client IDs) and `real-excel/` (real employee data) are gitignored and private. Don't read, print or commit them. Use `config.example.json` and workbooks built in tests instead.
- Every command that writes in bulk shows a preview and asks for confirmation (twice for `fix …` and `sync-check --disable-left-workers`). Keep that pattern for any new bulk write.

## Commands

```bash
npm test                 # ava, 220-odd tests, no network, a few seconds
npm run test:coverage    # c8 + ava; reports every file under src/
npx ava test/import.test.js              # one file
npx ava --match '*sync-check*'           # by test title
node src/index.js validate <file.xlsx>   # the only command that never calls Graph
```

Requires Node 24+ (`import … with { type: 'json' }`, `import.meta.dirname`); `.nvmrc` pins the major and CI reads it. ESM only, no build step, no linter or formatter configured.

## Layout

| Path | Role |
|---|---|
| `src/index.js` | commander definitions only. Each action does `loadConfig` → `initAuth` → call the command → `handleError`. |
| `src/auth.js` | MSAL device-code login. Token cache is encrypted to `~/.config/m365-users/msal-cache.json`. |
| `src/graph.js` | The only module that talks to Graph. One exported function per operation. |
| `src/utils/cache.js` | In-process cache used by `graph.js`. |
| `src/commands/*.js` | One file per command: prompts, output, orchestration. |
| `src/utils/excel.js` | Reads the `ACTIVOS` (required) and `CONCEJALES` (optional) sheets into string rows. |
| `src/utils/excel-importer.js` | Maps the Spanish employee-file headers to Graph user fields. |
| `src/validators/excel-rules.js` | Offline data-quality rules for `validate`. Add a rule by pushing to `RULES`. |
| `test/helpers/` | Fake Graph, fake prompts, temp HOME, workbook builder. |

## How the code is organised

**Graph access.** Commands never build requests themselves; they call `graph.js`. Most functions use the Graph SDK client; the three `fetch*Map` functions POST to `/$batch` with raw `fetch` in chunks of 20. When adding a `$filter` with user input, escape single quotes by doubling them (`value.replace(/'/g, "''")`), as the existing functions do.

**Cache.** `cache.js` lives for the process and has no TTL and no reset. Any new write function in `graph.js` must call the matching `invalidate*` helper, otherwise later reads in the same command return stale data. `listAllUsers` seeds the user cache, so a following `getUser`/`findUserByUpn` costs nothing.

**`listAllUsers` only returns enabled accounts** (or only disabled ones with `onlyDisabled`), and always drops `#EXT#` guests. Callers that say "all users" (`sync-check`, `validate-users`, `import`) therefore only see enabled accounts.

**Two sources for the domain.** `config.domain` (via `getDomain()`) completes short usernames and is enforced by `import`. `src/constants.js` hard-codes `arandadeduero.es` and is what `validate`, `sync-check` and `reset-password` use. Other tenant-specific values are hard-coded too: auto-assigned group names in `cache.js`, room-account prefixes (`sala-`, `salapz`) in `sync-check.js`, the `SPECIAL_ACCOUNTS` allowed to be cloud-only (in `constants.js`), SharePoint URLs in `reset-password.js`, target timezone in `fix-timezone.js`.

**Exit codes.** `validate`, `sync-check`, `reset-password` and `validate-users` call `process.exit` themselves (0 = clean, 1 = issues or abort). Other commands just return and let `index.js` handle thrown errors.

**Employee file.** Dates are `dd/mm/yyyy` strings; `excel.js` renders real Date cells the same way. `sync-check`, the validators and `rowToUserPayload` (hire date) parse that format explicitly; never hand these strings to `new Date()`, which reads them as `mm/dd/yyyy`. Header matching is case-insensitive in `excel-importer.js` but exact in `excel-rules.js` and `sync-check.js` (`ID responsable`, `Fecha de Baja`, `e_mail`, `Trabajador`).

## Conventions

- Code, comments and most CLI output are in English. `README.md` and the `validate-users` output are in Spanish; keep each file in the language it already uses.
- User-visible changes to a command need the matching README section updated, including its output example and the row in the "Resumen de comandos" table. README examples use made-up people only; never paste output from the real tenant or from `real-excel/`.
- Prompt types must be ones the installed `inquirer` registers (`select`, not the old `list`). Tests replace `inquirer.prompt`, so only `test/prompt-types.test.js` catches a wrong type.
- Console output goes through `chalk`; long operations use an `ora` spinner; prompts use `inquirer.prompt`.
- Underscore-prefixed exports (`_setFakeToken`, `_clearFakeToken`) are test hooks.

## Testing

Tests run commands end to end against a fake tenant; nothing reaches the network.

```js
import { cleanupHome } from './helpers/home.js';          // first: before any src import
import { mockGraph, mockPrompts, runCommand, writes } from './helpers/graph-mock.js';
import { writeWorkbook, dmy } from './helpers/excel.js';

const { initAuth, _setFakeToken } = await import('../src/auth.js');
```

- `helpers/home.js` points `os.homedir()` at a temp directory. It must be imported before anything from `src/`, because `auth.js` resolves its cache path at import time and `logout()` deletes that file. This only works with `"workerThreads": false` in the ava config: inside a worker thread, setting `process.env.HOME` does not change `os.homedir()`. Don't remove that setting.
- `mockGraph(t, handler)` replaces `globalThis.fetch` and records every request. Unhandled requests answer 404. Don't return 429/503/504: the Graph SDK retries those with real delays.
- `writes(calls)` is the list of state-changing requests. For anything that can modify the tenant, assert the exact writes, and assert `[]` for the declined/dry-run paths.
- `mockPrompts(t, [...])` answers `inquirer.prompt` calls in order; an unexpected prompt throws.
- `runCommand(fn)` captures output (ANSI stripped) and turns `process.exit(n)` into `{ exitCode: n }`.
- Tests that use these helpers must be `test.serial` (they patch globals).
- The cache cannot be reset, so give each test its own user IDs/UPNs. Each test file runs in its own process.

## CI and dependencies

- `.github/workflows/ci.yml` runs `npm test` on every push to `main` and on every pull request.
- Dependabot (`.github/dependabot.yml`) opens one grouped PR per ecosystem (npm, GitHub Actions) for minor and patch updates and a separate PR per major. Its CI may have run against an older base: before merging, run the suite with the PR merged onto current `main`.
- `@azure/msal-node` handles login and the session cache, which no test exercises end to end. Read the release notes before accepting a major.

## Known issues

- `findUserByUpn` returns `null` on any Graph error and does not escape quotes in the UPN, so `import` can mistake a failed lookup for "user does not exist" and attempt a create.
- `import` never prints the password it generates for a new account; the administrator has to run `reset-password` afterwards.
- `edit.js`, `search.js`, `assign-manager.js`, `validate-users.js` and the interactive `fix org-chart` loop have no tests yet.
