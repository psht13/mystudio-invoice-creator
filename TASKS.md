# MyStudio Invoice Creator Audit and Codex Refactor Prompts

Repository: `psht13/mystudio-invoice-creator`  
Audit type: static repository inspection through GitHub file access  
Primary goal: refactor toward a clean onion/layer architecture with strong test coverage while preserving current functionality exactly.

## Executive summary

This is a small React + Vite application that generates half-month time-tracking invoices as `.xlsx` files. The README says the app preserves the original Excel formatting by loading workbook templates from `public/templates/` and filling only dynamic values such as name, period, day rows, and totals.

The current implementation is understandable and compact, but it is not yet layered. The main business logic and Excel-writing logic live together in `src/invoice.ts`, while UI state, browser fetch, template selection, workbook generation, Blob download, validation messages, and page rendering live together in `src/App.tsx`.

The biggest refactor risk is not architectural complexity. The biggest risk is accidentally changing workbook behavior. The current output depends on exact date formatting, row placement, row collapsing rules, formula cells, template choice, total-hour calculation, and filename generation. Codex should lock those behaviors down with characterization tests before moving code.

## Files inspected

- `README.md`
- `package.json`
- `eslint.config.js`
- `tsconfig.json`
- `tsconfig.app.json`
- `tsconfig.node.json`
- `vite.config.ts`
- `src/main.tsx`
- `src/App.tsx`
- `src/invoice.ts`
- `tests/invoice.test.ts`
- latest commit metadata and patch context for the OOO range fix

## Current functionality to preserve

These are the non-negotiable behavior constraints for the refactor.

- The app remains a React + Vite app.
- It still generates `.xlsx` files from the existing templates in `public/templates/`.
- The generated workbook keeps the original template formatting.
- Template selection remains:
  - run day `1` uses `/templates/template-1st.xlsx`
  - run day `16` uses `/templates/template-16th.xlsx`
- Default person name remains `Pavlo Yurchenko`.
- `?name=Your%20Name` still pre-fills the name input.
- Empty or whitespace-only names still block generation and show the existing error message.
- Generated filename format remains:
  - `<sanitized name> Time Tracking - <M_D_YYYY>.xlsx`
- Period logic remains:
  - run day `16`: current month day 1 through day 15
  - run day `1`: previous month day 16 through previous month end
- Date display remains `M/D/YYYY`.
- Workday default hours remain `8`.
- Weekends are automatic and non-editable.
- Editable statuses remain `WORK`, `HOLIDAY`, `VACATION`, and `OOO`.
- Status labels remain `Work`, `Holiday`, `Vacation`, `OOO`, and `Weekend`.
- Only `WORK` days count toward total hours.
- Hours are sanitized the same way: non-finite or negative values become `0`; values are rounded to two decimals.
- Sheet rows still begin at row `15` and end at row `28`.
- Total formula remains `SUM(E15:E28)` in `E29`.
- Workbook cells still receive:
  - period label in `D5`
  - person name in `D8`
  - day labels in column `B`
  - work hours in column `E`
- Unused data rows are still cleared.
- Consecutive `WEEKEND`, `HOLIDAY`, and `VACATION` ranges still collapse into date ranges.
- Consecutive `OOO` days remain separate rows.
- UI labels, button text, and notice text stay unchanged unless a test explicitly captures and approves a change.

## Current strengths

The repo already has several good foundations:

- Small and readable codebase.
- Strict TypeScript is enabled in the app config.
- ESLint is already present.
- Vitest is already present.
- Invoice logic is already mostly isolated in `src/invoice.ts`.
- Existing tests cover:
  - default invoice selection
  - half-month period ranges
  - sheet entry generation
  - total-hour calculation
  - template-backed workbook generation
  - workbook cell output and total formula result

Those tests are especially valuable because they already check generated workbook contents with ExcelJS.

## Main architecture issues

### 1. `src/App.tsx` does too much

`App.tsx` currently owns:

- URL query-param parsing for the default name.
- UI state for person name, month, run day, overrides, loading state, and notices.
- Day status updates.
- Hour parsing and sanitization calls.
- Reset behavior.
- Template URL selection.
- Browser `fetch`.
- Workbook generation.
- Blob URL download.
- UI rendering.

That is manageable now, but it makes UI tests, service tests, and behavior preservation harder.

### 2. `src/invoice.ts` mixes domain logic and infrastructure

`invoice.ts` contains clean pure invoice logic, but it also imports `exceljs` and writes workbook cells. That means the domain layer has an infrastructure dependency.

For onion architecture, date/period/day-entry logic should be pure domain or application logic. ExcelJS workbook writing should live in an infrastructure adapter.

### 3. No explicit application layer

There is no use-case/service layer that says, “given form state and a template, produce an invoice download.” The UI directly orchestrates everything. A clean application service would make behavior easier to test without React or browser APIs.

### 4. No ports for browser and Excel side effects

The app has side effects that should be abstracted:

- template loading
- workbook writing
- browser file download
- current date/time
- object URL lifecycle

Ports should make these behaviors testable and keep React components thin.

### 5. Coverage gate is missing

`package.json` has `test`, `lint`, and `build`, but no coverage script or coverage threshold. Add coverage after characterization tests. Vitest supports coverage providers such as `v8` and supports line/function/branch/statement thresholds.

Useful references:

- https://main.vitest.dev/guide/coverage
- https://main.vitest.dev/config/coverage

### 6. CI appears absent

I did not find a `.github/workflows/ci.yml` or `.github/workflows/test.yml` through direct file checks. Add CI once the quality scripts exist.

### 7. Existing tests are good but narrow

The current tests focus on invoice logic and generated workbook cells. They do not appear to cover:

- React UI behavior.
- URL name prefill.
- empty-name error behavior.
- template fetch failures.
- download behavior.
- reset-period behavior.
- entry capacity validation from the UI.
- invalid month fallback behavior.
- filename sanitization.
- leap-year and year-boundary date cases.
- direct `generateWorkbookFromTemplate` overflow behavior.
- accessibility basics such as form labels and notice announcements.

## Efficiency notes

There are no major performance problems for the current app size. Still, these are worth tightening after behavior is locked down:

- Keep `buildSheetEntries` O(n), which it already is.
- Avoid moving ExcelJS into initial UI-render code paths if lazy imports are practical and behavior-safe.
- Split expensive workbook generation from UI state logic.
- Keep derived values memoized in React, but move the derivation into a hook or application service.
- Avoid direct repeated allocation of temporary structures where simple pure functions can reuse computed values.
- Avoid binary snapshot comparison of XLSX files because workbook output bytes can vary. Compare semantic workbook contents instead: cell values, formulas, styles, merged ranges, row heights, column widths, and sheet names where necessary.
- Consider validating `entries.length <= DATA_ROW_CAPACITY` in the application service, not just in UI, so direct service callers cannot overflow template rows.

## Proposed onion/layer architecture

Target dependency direction:

```text
domain -> application -> infrastructure/adapters -> main
```

Suggested structure:

```text
src/
  main.tsx

  domain/
    invoice/
      constants.ts
      types.ts
      dates.ts
      periods.ts
      day-records.ts
      sheet-entries.ts
      filenames.ts
      hours.ts

  application/
    ports/
      clock.ts
      template-loader.ts
      workbook-writer.ts
      file-downloader.ts
    services/
      invoice-selection-service.ts
      invoice-preview-service.ts
      invoice-generation-service.ts

  infrastructure/
    excel/
      exceljs-workbook-writer.ts
    browser/
      fetch-template-loader.ts
      browser-file-downloader.ts

  adapters/
    react/
      App.tsx
      components/
        InvoiceForm.tsx
        SummaryCards.tsx
        DayTable.tsx
        ActionsBar.tsx
        Notice.tsx
      hooks/
        useInvoiceForm.ts
      presenters/
        invoice-view-model.ts

tests/
  domain/
  application/
  infrastructure/
  adapters/
```

Rules:

- `domain` must not import React, ExcelJS, Vite, DOM APIs, `fetch`, or browser download APIs.
- `application` may import domain and application ports, but not React, ExcelJS, DOM APIs, or browser globals.
- `infrastructure` implements application ports with ExcelJS and browser APIs.
- `adapters/react` maps UI events to application services and renders view models.
- `main.tsx` should be the composition root.

## Recommended test strategy

### Characterization first

Before refactoring, add tests that describe current output and UI behavior. These tests should fail if Codex accidentally changes functionality.

### Coverage target

Aim for at least:

```text
90% lines
90% branches
90% functions
90% statements
```

For this repo, meaningful coverage matters more than percentage. The high-value areas are:

- period/date boundaries
- workbook cell output
- row collapsing behavior
- UI generate/reset behavior
- template load failures
- browser download adapter behavior
- application service orchestration

### Testing tools

Keep Vitest. Add:

```text
@vitest/coverage-v8
@testing-library/react
@testing-library/user-event
@testing-library/jest-dom
jsdom
```

React Testing Library documentation recommends testing React components through user-facing behavior rather than implementation details:

- https://testing-library.com/docs/react-testing-library/intro/
- https://testing-library.com/docs/user-event/intro/

## Codex prompts

Use these prompts in order. They are intentionally split into small, reviewable steps so functionality stays intact.

---

## Prompt 1 - Add characterization tests before touching architecture

```text
You are working in https://github.com/psht13/mystudio-invoice-creator.

Goal: add characterization tests that lock down current behavior before any refactor. Do not change production behavior.

Important invariant:
No functionality changes are allowed. Preserve the generated workbook behavior, UI strings, file names, date labels, template selection, row positions, formula, and status behavior exactly.

Tasks:
1. Add or extend tests for pure invoice functions in `src/invoice.ts`:
   - `formatDate`
   - `formatFileDate`
   - `formatHours`
   - `sanitizeHours`
   - `sanitizeFileName`
   - `buildInvoiceFileName`
   - `parseMonthValue`
   - `getDefaultInvoiceSelection`
   - `getPeriod`
   - `listPeriodDays`
   - `buildDayRecords`
   - `calculateTotalWorkHours`
   - `buildSheetEntries`
2. Include edge cases:
   - leap year February
   - non-leap February
   - January run day `1`
   - December to January transition
   - invalid `monthValue`
   - negative hours
   - NaN and Infinity hours
   - decimal rounding
   - empty filename name
   - invalid filename characters
   - consecutive weekend ranges
   - consecutive holiday ranges
   - consecutive vacation ranges
   - consecutive OOO days staying separate
   - mixed non-work statuses not merging
3. Add workbook characterization tests using the existing templates:
   - `template-1st.xlsx`
   - `template-16th.xlsx`
4. Workbook tests should assert semantic contents, not binary equality:
   - worksheet name selection behavior
   - `D5`
   - `D8`
   - rows `B15:E28`
   - `E29` formula and result
   - cleared unused rows
   - template formatting is still present for representative cells
5. Do not refactor yet.
6. Run:
   - npm run lint
   - npm test
   - npm run build
7. Return a summary of new tests and any surprising existing behavior that was preserved.
```

---

## Prompt 2 - Add coverage, CI, and formatting gates

```text
You are working in https://github.com/psht13/mystudio-invoice-creator.

Goal: add quality gates while preserving behavior.

Important invariant:
No runtime behavior changes. Do not edit generated invoice logic or UI behavior unless only required for testability and covered by exact tests.

Tasks:
1. Add Vitest coverage using the V8 provider.
2. Install the needed dev dependency:
   - @vitest/coverage-v8
3. Add a coverage config with thresholds:
   - lines: 90
   - branches: 90
   - functions: 90
   - statements: 90
4. Ensure tests include relevant source files and exclude:
   - test files
   - build output
   - config files unless intentionally tested
5. Add scripts:
   - `test:coverage`
   - `typecheck`
   - `format`
   - `ci`
6. `ci` should run:
   - typecheck
   - lint
   - test:coverage
   - build
7. Add Prettier if not already present.
8. Add `.prettierrc` with simple defaults consistent with current code style.
9. Add `.github/workflows/ci.yml` using Node 22 and `npm ci`.
10. Do not weaken existing TypeScript settings.
11. Run:
   - npm run ci
12. Return:
   - commands run
   - coverage numbers
   - files changed
   - any excluded files and why
```

---

## Prompt 3 - Freeze browser/UI behavior with component tests

```text
You are working in https://github.com/psht13/mystudio-invoice-creator.

Goal: add UI behavior tests before splitting `App.tsx`.

Important invariant:
No UI text or behavior changes. Preserve all current labels, buttons, notices, default values, and the query-param name prefill.

Tasks:
1. Add React component testing with:
   - @testing-library/react
   - @testing-library/user-event
   - @testing-library/jest-dom
   - jsdom
2. Configure Vitest to support React component tests.
3. Add tests for `App`:
   - default name is `Pavlo Yurchenko`
   - `?name=Jane%20Doe` pre-fills the name
   - blank query param falls back to default name
   - invoice month input updates summary cards
   - run day buttons update invoice date and pay period
   - day status select changes a workday to holiday/vacation/OOO
   - hours input changes total work hours
   - negative hours sanitizes to `0`
   - reset period removes overrides only for the current period
   - empty name blocks generation and shows the existing error
   - generate button shows `Generating...` while generation is in progress
   - template fetch failure shows the existing error behavior
4. Mock `fetch`, `URL.createObjectURL`, `URL.revokeObjectURL`, and link click behavior.
5. Do not inspect private React state. Test through visible UI and browser-observable effects.
6. Run:
   - npm run ci
7. Return:
   - tested behavior list
   - remaining UI branches not covered
```

---

## Prompt 4 - Split pure domain logic out of `src/invoice.ts`

```text
You are working in https://github.com/psht13/mystudio-invoice-creator.

Goal: move pure invoice logic into a domain layer without changing behavior.

Important invariant:
All existing tests and characterization tests must pass unchanged. Do not change output strings, dates, total hours, row labels, or filenames.

Target files:
- `src/domain/invoice/constants.ts`
- `src/domain/invoice/types.ts`
- `src/domain/invoice/dates.ts`
- `src/domain/invoice/periods.ts`
- `src/domain/invoice/day-records.ts`
- `src/domain/invoice/sheet-entries.ts`
- `src/domain/invoice/hours.ts`
- `src/domain/invoice/filenames.ts`
- `src/domain/invoice/index.ts`

Tasks:
1. Move pure types and constants out of `src/invoice.ts`.
2. Move pure functions out of `src/invoice.ts`.
3. Keep backward-compatible exports from `src/invoice.ts` initially so existing imports do not break.
4. Ensure the new domain files do not import:
   - react
   - exceljs
   - vite
   - DOM/browser APIs
5. Preserve exact behavior.
6. Add or update tests to import from the new domain modules.
7. Run:
   - npm run ci
8. Return:
   - new module map
   - old exports still kept for compatibility
   - confirmation that no domain module imports infrastructure/UI dependencies
```

---

## Prompt 5 - Extract ExcelJS workbook writing into infrastructure

```text
You are working in https://github.com/psht13/mystudio-invoice-creator.

Goal: move ExcelJS-specific workbook generation out of the domain layer.

Important invariant:
Generated workbook semantic output must remain exactly the same.

Tasks:
1. Create an application port:
   - `src/application/ports/workbook-writer.ts`
2. Create an ExcelJS implementation:
   - `src/infrastructure/excel/exceljs-workbook-writer.ts`
3. Move `generateWorkbookFromTemplate` implementation into the ExcelJS infrastructure adapter.
4. Keep a temporary compatibility export from `src/invoice.ts` if needed.
5. Preserve:
   - worksheet lookup by `Invoice` with fallback to first worksheet
   - `fullCalcOnLoad = true`
   - `D5` period label
   - `D8` person name
   - clearing `B15:E28`
   - writing entries to `B` and `E`
   - `E29` formula object with result
   - returned `ArrayBuffer` shape
   - `Template worksheet is missing` error
6. Add tests against the workbook-writer port and ExcelJS implementation using both templates.
7. Run:
   - npm run ci
8. Return:
   - files moved
   - compatibility exports kept
   - generated workbook assertions that prove behavior stayed intact
```

---

## Prompt 6 - Add an application invoice-generation service

```text
You are working in https://github.com/psht13/mystudio-invoice-creator.

Goal: create a framework-free application service that coordinates invoice generation.

Important invariant:
The React UI should produce the same workbook, filename, notices, and errors as before.

Tasks:
1. Create application ports:
   - `TemplateLoader`
   - `WorkbookWriter`
   - `FileDownloader`
   - `Clock` if useful for default selection tests
2. Create `src/application/services/invoice-generation-service.ts`.
3. The service should accept simple DTOs:
   - personName
   - monthValue
   - runDay
   - overrides
4. The service should:
   - validate non-empty person name
   - compute period
   - list period days
   - build day records
   - build sheet entries
   - validate `entries.length <= DATA_ROW_CAPACITY`
   - choose template path exactly as before
   - build period label exactly as before
   - build file name exactly as before
   - ask `TemplateLoader` for the template
   - ask `WorkbookWriter` for the workbook buffer
   - ask `FileDownloader` to download it
   - return a success or error result with the same notice text currently shown by the UI
5. Keep browser-specific behavior out of this service.
6. Add fake implementations for the ports in tests.
7. Add service tests for:
   - success path
   - empty name
   - capacity limit
   - template load failure
   - workbook writer failure
   - downloader failure
   - template path selection for both run days
8. Refactor `App.tsx` to use this service without changing UI behavior.
9. Run:
   - npm run ci
10. Return:
   - service API
   - covered scenarios
   - proof that UI text remains unchanged
```

---

## Prompt 7 - Add browser infrastructure adapters

```text
You are working in https://github.com/psht13/mystudio-invoice-creator.

Goal: move browser side effects into small infrastructure adapters.

Important invariant:
Download behavior and template URLs must remain unchanged.

Tasks:
1. Create:
   - `src/infrastructure/browser/fetch-template-loader.ts`
   - `src/infrastructure/browser/browser-file-downloader.ts`
2. `FetchTemplateLoader` should:
   - call `fetch(templatePath)`
   - throw `Failed to load template file: ${templatePath}` when response is not OK
   - return `arrayBuffer()`
3. `BrowserFileDownloader` should preserve current behavior:
   - create Blob with MIME type `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
   - create object URL
   - create temporary anchor
   - set `href`
   - set `download`
   - append to document body
   - click link
   - remove link
   - revoke object URL
4. Add tests with mocked browser APIs.
5. Refactor `App.tsx` or the composition module to inject these adapters.
6. Run:
   - npm run ci
7. Return:
   - adapter APIs
   - tested browser side effects
```

---

## Prompt 8 - Split React UI into components and a hook

```text
You are working in https://github.com/psht13/mystudio-invoice-creator.

Goal: make the React adapter thin and maintainable by splitting UI code.

Important invariant:
No UI behavior or styling changes. Keep class names, labels, status options, notice text, and button text unchanged.

Suggested files:
- `src/adapters/react/App.tsx`
- `src/adapters/react/hooks/useInvoiceForm.ts`
- `src/adapters/react/components/InvoiceControls.tsx`
- `src/adapters/react/components/SummaryCards.tsx`
- `src/adapters/react/components/DayTable.tsx`
- `src/adapters/react/components/ActionsBar.tsx`
- `src/adapters/react/components/Notice.tsx`

Tasks:
1. Move form state and derived data into `useInvoiceForm`.
2. Move JSX sections into focused components.
3. Keep existing CSS classes exactly the same.
4. Keep current DOM structure as close as practical so component tests still pass.
5. Keep `src/App.tsx` as a compatibility re-export if needed, or update imports carefully.
6. Do not introduce global state.
7. Do not change generated workbook behavior.
8. Run:
   - npm run ci
9. Return:
   - before/after component tree
   - any DOM changes made and why
```

---

## Prompt 9 - Add dependency-boundary enforcement

```text
You are working in https://github.com/psht13/mystudio-invoice-creator.

Goal: enforce the onion architecture so future changes do not collapse layers again.

Important invariant:
Do not change runtime behavior.

Tasks:
1. Add a lightweight dependency boundary test or ESLint rule.
2. Enforce:
   - `src/domain/**` cannot import React, ExcelJS, Vite, DOM/browser infrastructure, or application services.
   - `src/application/**` cannot import React, ExcelJS implementations, DOM/browser infrastructure, or adapter components.
   - `src/infrastructure/**` can import application ports and domain, but not React components.
   - `src/adapters/react/**` can import application services and domain view models, but should not import ExcelJS directly.
3. Add tests or lint checks that fail on boundary violations.
4. Run:
   - npm run ci
5. Return:
   - boundary rules
   - examples of allowed and disallowed imports
```

---

## Prompt 10 - Harden validation while preserving current behavior

```text
You are working in https://github.com/psht13/mystudio-invoice-creator.

Goal: make validation more explicit without changing user-visible behavior.

Important invariant:
Do not change user-visible text or generated workbook output. Any new validation must preserve current outcomes unless a test documents an approved bug fix.

Tasks:
1. Create explicit domain/application validation helpers for:
   - person name presence
   - run day value
   - month value parsing
   - hours sanitization
   - sheet row capacity
2. Preserve current fallback behavior for invalid month values.
3. Preserve current `sanitizeHours` behavior.
4. Preserve current `sanitizeFileName` behavior.
5. Ensure `entries.length > DATA_ROW_CAPACITY` returns the same error text currently used in `App.tsx`.
6. Add tests for validation helpers and service results.
7. Run:
   - npm run ci
8. Return:
   - validations added
   - behavior intentionally preserved
   - any potential future bug fixes deferred because they would change behavior
```

---

## Prompt 11 - Optimize safe hot paths and bundle behavior

```text
You are working in https://github.com/psht13/mystudio-invoice-creator.

Goal: apply safe optimizations after tests and layers are in place.

Important invariant:
No functional changes and no generated workbook changes.

Tasks:
1. Keep invoice calculations O(n).
2. Avoid recomputing period days, day records, entries, and totals unnecessarily in React.
3. Consider lazy-loading the ExcelJS workbook writer only when generating a workbook, if it does not change behavior.
4. Ensure lazy loading does not break tests, build, or download flow.
5. Avoid adding dependencies unless justified.
6. Add or update tests to prove:
   - generated workbook output is unchanged
   - errors are unchanged
   - UI loading state is unchanged
7. Measure bundle impact if possible using Vite build output.
8. Run:
   - npm run ci
9. Return:
   - optimizations applied
   - why each is behavior-safe
   - build output comparison if available
```

---

## Prompt 12 - Add final docs and cleanup compatibility exports

```text
You are working in https://github.com/psht13/mystudio-invoice-creator.

Goal: finish the refactor with clean docs and no dead compatibility code.

Important invariant:
Functionality remains intact.

Tasks:
1. Remove temporary compatibility exports only after all imports have migrated.
2. Update README project layout.
3. Add `docs/architecture.md` explaining:
   - domain layer
   - application layer
   - infrastructure layer
   - React adapter layer
   - composition root
   - dependency rules
4. Add a testing section documenting:
   - npm run lint
   - npm test
   - npm run test:coverage
   - npm run build
   - npm run ci
5. Document behavior-preservation tests, especially workbook tests.
6. Confirm default name and query-param prefill behavior remain documented.
7. Run:
   - npm run ci
8. Return:
   - final folder tree
   - coverage numbers
   - commands run
   - any intentional non-changes made to preserve behavior
```

## Final acceptance criteria

The refactor is successful when all of these are true:

```text
npm run ci passes
coverage >= 90% lines, branches, functions, and statements
domain has no React, ExcelJS, Vite, DOM, or browser API imports
application has no React, ExcelJS implementation, DOM, or browser API imports
ExcelJS lives in infrastructure
browser fetch/download logic lives in infrastructure
React components are thin adapters over application services
generated workbook semantic tests pass for both templates
current UI behavior tests pass
README and architecture docs describe the new structure
no generated invoice behavior changes unless explicitly approved by tests
```

## Suggested review checklist for your second audit

When you send the result for the next audit, I would check:

- Did Codex preserve current workbook cell values, formulas, row clearing, and template selection?
- Did it add coverage before major refactoring?
- Did it avoid moving behavior and architecture in one huge commit?
- Does `domain` remain pure?
- Does `application` avoid React, ExcelJS, and browser globals?
- Do browser side effects have small tests?
- Are UI tests based on visible behavior, not private state?
- Is coverage meaningful rather than only percentage-driven?
- Are all old imports cleaned up?
- Does CI run the same commands developers are expected to run locally?
