# MyStudio Invoice Creator

React + Vite app for generating half-month time-tracking invoices as `.xlsx` files.
It preserves the original Excel formatting by loading template files from `public/templates/` and filling only dynamic values (name, period, day rows, and totals).

## Project layout

```text
.
|-- docs/
|   `-- architecture.md
|-- public/
|   `-- templates/
|       |-- template-1st.xlsx
|       `-- template-16th.xlsx
|-- src/
|   |-- adapters/react/        React components, hooks, and UI state
|   |-- application/           Invoice generation use case and ports
|   |-- domain/invoice/        Pure invoice rules, dates, rows, totals, names
|   |-- infrastructure/        Browser and ExcelJS implementations
|   |-- main.tsx               React browser entrypoint
|   |-- App.css
|   `-- index.css
|-- tests/
|   |-- App.test.tsx
|   |-- browser-adapters.test.ts
|   |-- dependency-boundaries.test.ts
|   |-- invoice-generation-service.test.ts
|   `-- invoice.test.ts
`-- vitest.config.ts
```

See [docs/architecture.md](docs/architecture.md) for the layer responsibilities and dependency rules.

## Run

```bash
npm install
npm run dev
```

## Testing

```bash
npm run lint
npm test
npm run test:coverage
npm run build
npm run ci
```

`npm run ci` is the release gate. It runs TypeScript, lint, coverage, and the production build in sequence.

The tests intentionally lock behavior while the code is split into layers:

- `tests/invoice.test.ts` covers pure invoice rules and workbook behavior, including both templates, generated rows, formulas, preserved styles, first-worksheet fallback, missing worksheet errors, and lazy writer parity with the direct ExcelJS writer.
- `tests/invoice-generation-service.test.ts` covers the application use case, validation messages, selected template paths, sanitized hours, filename generation, and port call order.
- `tests/App.test.tsx` covers the React adapter behavior, including default name, `?name=` query-param prefill, period controls, editable day overrides, reset behavior, loading state, and browser download errors.
- `tests/browser-adapters.test.ts` covers browser fetch and download adapters.
- `tests/dependency-boundaries.test.ts` checks that source imports stay within the documented architecture boundaries.

## Change default name (for forks)

The default value in the **Name** field is set in `src/adapters/react/hooks/useInvoiceForm.ts`.
You can also prefill it per-link by opening the app with `?name=Your%20Name`, for example:

```text
http://localhost:5173/?name=Jane%20Doe
```

Find this line:

```ts
const DEFAULT_PERSON_NAME = 'Pavlo Yurchenko'
```

Replace `'Pavlo Yurchenko'` with your preferred default.
