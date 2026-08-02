# OfficeEditor Web Demo

This React/Vite client is the local demonstration UI for `OfficeEditor.Api`. It is not a separately supported product or a production-ready hosted service.

## Run locally

From the repository root, start the API and client together:

```bash
make dev
```

Then open <http://localhost:5173/>. To run only the client:

```bash
cd OfficeEditor.Web.Client
npm install
npm run dev
```

The client expects the local API at `http://localhost:5001`.

## Demo screens

- **Render** — render whitelisted reference decks to PNG or SVG with server timings.
- **Generate** — edit the demo title/theme, generate PPTX, preview slides, and download the deck.
- **Any render** — upload a PPTX and render it locally.
- **Compare** — compare OfficeEditor rendering with an optional locally installed LibreOffice/Poppler toolchain.

## Checks

```bash
npm run lint
npm run build
npm audit
```

## Security boundary

This client and `OfficeEditor.Api` are local demo surfaces. They do not provide production authentication, tenant isolation, quotas, or sandboxing. Do not expose them directly to the public internet without an application-specific security boundary. See [`../SECURITY.md`](../SECURITY.md).
