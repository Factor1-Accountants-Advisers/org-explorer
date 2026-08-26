# Org Explorer

Microsoft 365 org chart navigator for Factor1. Hierarchy is **Company → Department → Team → Role**. Team is stored in Exchange **CustomAttribute1** (`onPremisesExtensionAttributes.extensionAttribute1` in Graph). Embeddable in SharePoint via Vercel.

Repo: [Factor1-Accountants-Advisers/org-explorer](https://github.com/Factor1-Accountants-Advisers/org-explorer)

## Local preview

View-only (and demo admin, which does not call Graph):

```bash
npm run preview
```

- Demo data: http://localhost:8080/?demo=1
- Live Entra sign-in: http://localhost:8080/

Admin writes against Microsoft 365 need the API:

```bash
npm run dev
```

That runs `vercel dev` (sign in to Vercel the first time). Copy the production env vars into `.env.local` so `/api/admin` can acquire an app token.

## Entra app

Use the existing app registration (`CONFIG.clientId` in `app.js`).

Delegated (SPA, everyone):

- `User.Read`
- `User.Read.All` (admin consent)

Application (server, admin writes only):

- `User.ReadWrite.All`
- `GroupMember.Read.All`

Both application permissions need **admin consent**. Create a client secret and store it in Vercel as `ENTRA_CLIENT_SECRET`. Do not put the secret in the frontend.

Redirect URIs (SPA):

- `http://localhost:8080/` for `npm run preview`
- `http://localhost:3000/` for `vercel dev` (check the port it prints)
- The production Vercel origin, e.g. `https://org-explorer.vercel.app/`

### Admin group

1. In Entra ID create a security group, e.g. **Org Explorer Admins**.
2. Add those people under **Members** (not only Owners).
3. Copy the group **Object ID** from Entra → Groups → the group → Overview. Names and mail nicknames will not work.
4. In the Vercel project that owns the live URL, set `ADMIN_GROUP_ID` exactly (case-sensitive) for **Production**, then **Redeploy**. Existing deployments do not receive new variables.

### Team / CustomAttribute1

Graph can write `extensionAttribute1` only for **cloud-only** users (`onPremisesSyncEnabled` is false or null). Hybrid or Exchange-mastered mailboxes return an error in the edit drawer; those must be updated in Exchange. Confirm CustomAttribute1 is unused in the tenant before the first live save.

## Deploy to Vercel

1. Import `Factor1-Accountants-Advisers/org-explorer` (framework **Other**, root of the repo).
2. Set environment variables for Production (and Preview if you want admin there):

| Variable | Value |
| --- | --- |
| `ENTRA_CLIENT_ID` | Same as `CONFIG.clientId` in `app.js` |
| `ENTRA_CLIENT_SECRET` | App registration client secret |
| `ENTRA_TENANT_ID` | `factor1.com.au` or the tenant GUID |
| `ADMIN_GROUP_ID` | Org Explorer Admins object ID |

3. Deploy, then add the Vercel origin as an SPA redirect URI in Entra ID.
4. Grant admin consent for the application permissions if you have not already.
5. After changing env vars, **Redeploy** the Production deployment. New variables are not picked up by an already-running deploy.
