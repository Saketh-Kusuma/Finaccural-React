# FinAccrual React Add-in

This is an isolated React implementation of the Excel task-pane add-in. It runs on port `3001`, uses a distinct Office add-in ID, and leaves `../Frontend` and `../Backend` untouched.

## Run locally

```powershell
npm install
npm run dev-server
npm start
```

Sideload `manifest.xml` if you prefer to register the add-in manually. The original add-in remains available on port `3000` with its original manifest.

## Build

```powershell
npm run build
```

The React implementation keeps the original task-pane stylesheet and visual class system, while moving view selection, sign-in, plan selection, checkout handoff, OAuth handoff, dashboard state, and session persistence into React components.
