# Org Explorer

Single-page Microsoft 365 org chart navigator for Factor1. Embeddable in SharePoint via Vercel.

## Local preview

```bash
npx serve . -l 8080
```

- Demo data: http://localhost:8080/?demo=1
- Live Entra sign-in: http://localhost:8080/

## Entra app

- Application (client) ID is set in `index.html` (`CONFIG.clientId`)
- Redirect URI (SPA): `http://localhost:8080/` for local, then your Vercel URL in production
- Delegated permissions: `User.Read`, `User.Read.All` (admin consent required)

## Deploy

Connect this repo to Vercel. After deploy, add the Vercel origin as a SPA redirect URI in Entra ID.
