# MyStudio Invoice Creator

React + Vite app for generating half-month time-tracking invoices as `.xlsx` files.
It preserves the original Excel formatting by loading template files from `public/templates/` and filling only dynamic values (name, period, day rows, and totals).

## Run

```bash
npm install
npm run dev
```

## Build and checks

```bash
npm run lint
npm test
npm run build
```

## Change default name (for forks)

The default value in the **Name** field is set in `src/App.tsx`.
You can also prefill it per-link by opening the app with `?name=Your%20Name`, for example:

```text
http://localhost:5173/?name=Jane%20Doe
```

Find this line:

```ts
const DEFAULT_PERSON_NAME = 'Pavlo Yurchenko'
```

Replace `'Pavlo Yurchenko'` with your preferred default.
