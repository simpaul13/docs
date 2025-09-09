## Quick repo snapshot

- Purpose: small Express app that fills a DOCX template using `docxtemplater` and serves the generated file to users.
- Run-time: Node.js app started from `server.js` (listens on port 3000).
- Key files: `server.js`, `public/index.html`, `templates/default-template.docx`, `package.json`.

## Big picture (what an agent must know)

- `server.js` is the single-server process. It: serves static files from `public/`, exposes three routes (`/`, `/download-template`, `/generate`), receives file uploads via `multer`, renders templates with `pizzip` + `docxtemplater`, and returns a generated DOCX buffer.
- The template data shape (set in `doc.setData(...)`) is: `{ name, date, orderId, table }` where `table` is an array of rows. Each row in current code has keys: `col1`, `col2`, `col3`, `col4`.
- The app supports either the built-in `templates/default-template.docx` or a user-uploaded `.docx` (multipart form field `template`). Uploaded files are stored in `uploads/` (temporary) and removed after generation.

## Important patterns & conventions in this codebase

- Placeholder names used in templates: `name`, `date`, `orderId`, and `table` (table loops). Keep the exact names when editing templates or changing server code.
- Docxtemplater options used: `{ paragraphLoop: true, linebreaks: true }` — expect multi-paragraph table loops and linebreak handling.
- File cleanup: uploaded files are deleted synchronously (`fs.unlinkSync(req.file.path)`) after processing. If you change upload handling, keep or intentionally replace this behavior to avoid orphaned files.
- Static UI: `public/index.html` posts the form to `/generate` with fields `name`, `date`, `orderId` and file input `template`.

## How to run / developer workflow (explicit)

1. Install deps:

```powershell
npm install
```

2. Start in development (auto-reload):

```powershell
npm run dev
```

3. Start production:

```powershell
npm start
```

4. Open the app: http://localhost:3000

Notes: `package.json` exposes `start` -> `node server.js` and `dev` -> `nodemon server.js`.

## Debugging & common fixes

- Template rendering errors: `doc.render()` is wrapped in a try/catch; logged errors usually indicate mismatched placeholder names or invalid template structure. Check the template for the placeholders listed above.
- If uploads accumulate, check that `uploads/` exists and that `fs.unlinkSync` runs. When changing to asynchronous file handling, ensure you still remove temp files.
- To change port or environment behavior, modify `PORT` in `server.js` or export `PORT` before starting; the code currently hardcodes `const PORT = 3000`.

## Integration points & external deps

- Uses these npm packages: `express`, `multer`, `pizzip`, `docxtemplater`, and `faker` (unused in server but present in `package.json`). If adding features that rely on external services, there are no current hooks—network calls should be added in `server.js` or new modules.

## Where to look for examples

- UI/example payload: `public/index.html` — shows exact field names and form encoding (`multipart/form-data`).
- Server flow: `server.js` — all request handling, template reading, rendering options, table generation, and file download headers.
- Default template: `templates/default-template.docx` — update placeholders here to match any server-side changes.

## Short goals for an agent (prioritized)

1. Preserve placeholder names when editing templates or server data keys.
2. Keep the synchronous cleanup or replace with safe async cleanup and test for orphaned files.
3. When modifying routes, update `public/index.html` and `templates/*` examples to remain consistent.

---
If any of these sections are unclear or you want the agent to adopt a different pattern (e.g., async upload cleanup, config via ENV), tell me which area to expand and I will iterate.

## Table loop (cost estimate) example

Use this example when your DOCX template contains a repeating cost-estimate table (see `templates/default-template.docx`).

- DOCX: edit the single row that should repeat. Put `{#table}` before the first tag in the row and `{/table}` after the last tag in the same row. Example row (cells left→right):

	`{#table}{index}` | `{Description}` | `{qty}` | `{unit_price}` | `{sub_total}{/table}`

	Make sure each tag is contiguous in Word (Word may split tags across runs when styling).

- Server-side data shape: pass `table` as an array of objects where keys match tags exactly (case-sensitive). Example snippet to compute subtotals, VAT and totals:

	```js
	const rows = [
		{ Description: 'Widget A', qty: 2, unit_price: 50.0 },
		{ Description: 'Service B', qty: 1, unit_price: 75.0 }
	];

	const table = rows.map((r, i) => {
		const sub = (r.qty || 0) * (r.unit_price || 0);
		return {
			index: i + 1,
			Description: r.Description || '',
			qty: r.qty || 0,
			unit_price: (r.unit_price || 0).toFixed(2),
			sub_total: sub.toFixed(2)
		};
	});

	const subtotal = table.reduce((s, row) => s + parseFloat(row.sub_total || 0), 0);
	const discount_percent = parseFloat(req.body.discount_percent || 0);
	const discount_flat = parseFloat(req.body.discount_flat || 0);
	const discount_from_percent = (subtotal * (discount_percent / 100));
	const discounted = subtotal - discount_from_percent - discount_flat;
	const VAT_RATE = parseFloat(req.body.vat_rate || 0.12);
	const VAT = discounted * VAT_RATE;
	const total = discounted + VAT;

	doc.setData({
		table,
		sub_total: subtotal.toFixed(2),
		discount_percent: discount_percent.toString(),
		discount_flat: discount_flat.toFixed(2),
		VAT: VAT.toFixed(2),
		total: total.toFixed(2),
		// other placeholders...
	});
	```

Notes:
- Keep `{#table}` and `{/table}` in the same row.
- Format numbers as strings on the server for consistent decimals/currency symbols.
- If tags are split in Word, use Find/Replace to rejoin them into a single run.
