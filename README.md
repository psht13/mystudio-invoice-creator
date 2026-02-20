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

Find this line:

```ts
const [personName, setPersonName] = useState('Pavlo Yurchenko')
```

Replace `'Pavlo Yurchenko'` with your preferred default.
