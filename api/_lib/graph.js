export const GRAPH = "https://graph.microsoft.com/v1.0";

export function graphMessage(err, fallback) {
  return err?.error?.message || err?.error_description || fallback;
}

export function env(name) {
  if (process.env[name]) return process.env[name];
  const found = Object.keys(process.env).find((key) => key.toLowerCase() === name.toLowerCase());
  return found ? process.env[found] : "";
}

export async function getAppToken() {
  const tenant = env("ENTRA_TENANT_ID");
  const clientId = env("ENTRA_CLIENT_ID");
  const clientSecret = env("ENTRA_CLIENT_SECRET");
  if (!tenant || !clientId || !clientSecret) {
    throw new Error("Server is missing Entra app credentials.");
  }

  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    grant_type: "client_credentials",
    scope: "https://graph.microsoft.com/.default",
  });

  const res = await fetch(
    `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body,
    }
  );
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(graphMessage(data, "Could not acquire an app token."));
  }
  return data.access_token;
}
