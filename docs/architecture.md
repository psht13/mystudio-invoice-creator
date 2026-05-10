# Architecture

The app is split into small layers so invoice behavior can be tested without React, browser APIs, or ExcelJS. The dependency direction is inward: UI and infrastructure can depend on application/domain code, but domain and application code cannot depend on UI, browser, build-tool, or concrete workbook implementations.

## Domain layer

Location: `src/domain/invoice/`

The domain layer contains pure invoice rules:

- invoice period selection and half-month date ranges
- day record creation, weekend detection, and override handling
- work-hour sanitization and total calculation
- sheet-entry labels and non-work range grouping
- invoice filename/date formatting
- domain validation helpers

This layer must stay framework-free and side-effect-free. It must not import React, browser APIs, Vite, ExcelJS, application services, infrastructure implementations, or React adapters.

## Application layer

Location: `src/application/`

The application layer contains the invoice generation use case and the ports it needs:

- `services/invoice-generation-service.ts` coordinates validation, domain calculations, template loading, workbook writing, and file download.
- `services/invoice-generation-validation.ts` adapts domain validation into user-facing generation notices.
- `ports/` defines `TemplateLoader`, `WorkbookWriter`, and `FileDownloader`.

The application layer depends on the domain layer and port interfaces only. It must not import React, browser APIs, ExcelJS, infrastructure implementations, React adapters, or browser entrypoints.

## Infrastructure layer

Location: `src/infrastructure/`

The infrastructure layer implements application ports:

- `browser/fetch-template-loader.ts` loads template files with `fetch`.
- `browser/browser-file-downloader.ts` downloads the generated workbook through a temporary browser anchor.
- `excel/exceljs-workbook-writer.ts` writes workbook contents with ExcelJS.
- `excel/lazy-exceljs-workbook-writer.ts` lazily imports the ExcelJS writer so the heavy dependency is loaded only when generation starts.

Infrastructure may depend on domain constants/helpers and application ports. It must not depend on React components or application services.

## React adapter layer

Location: `src/adapters/react/`

The React adapter turns application/domain state into the browser UI:

- `App.tsx` renders the page structure.
- `components/` contains presentational controls, summary cards, day rows, actions, and notices.
- `hooks/useInvoiceForm.ts` owns UI state, derives invoice previews from the domain layer, and calls the invoice generation use case.

React adapters may depend on domain rules and application services. They must not import ExcelJS directly.

## Composition root

Locations: `src/main.tsx` and `src/adapters/react/hooks/useInvoiceForm.ts`

`src/main.tsx` is the browser entrypoint that mounts React. The invoice generation composition happens in `useInvoiceForm.ts`, where `InvoiceGenerationService` is wired with the browser template loader, lazy ExcelJS workbook writer, and browser file downloader.

Keeping composition at the UI edge preserves the inner dependency rules while still using concrete browser and workbook implementations when a user generates a file.

## Dependency rules

- Domain imports no outer layers and no framework, browser, build-tool, or workbook packages.
- Application imports domain code and application port types; it does not import infrastructure implementations.
- Infrastructure implements application ports and can use concrete browser or ExcelJS APIs.
- React adapters call domain/application code and stay free of direct ExcelJS imports.
- `src/main.tsx` and the React hook composition code are the outer edge where concrete implementations are selected.
- `tests/dependency-boundaries.test.ts` enforces the main import boundaries.
