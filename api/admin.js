const GRAPH = "https://graph.microsoft.com/v1.0";

function json(res, status, body) {
  res.status(status).json(body);
}

function normalizeField(value) {
  if (value == null) return "";
  return String(value).trim();
}

function graphMessage(err, fallback) {
  return err?.error?.message || err?.error_description || fallback;
}

async function getAppToken() {
  const tenant = process.env.ENTRA_TENANT_ID;
  const clientId = process.env.ENTRA_CLIENT_ID;
  const clientSecret = process.env.ENTRA_CLIENT_SECRET;
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

async function callerFromUserToken(authHeader) {
  if (!authHeader?.startsWith("Bearer ")) return null;
  const token = authHeader.slice(7);
  const res = await fetch(`${GRAPH}/me?$select=id,displayName,userPrincipalName`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) return null;
  return res.json();
}

async function isGroupMember(appToken, userId, groupId) {
  const res = await fetch(`${GRAPH}/users/${encodeURIComponent(userId)}/checkMemberGroups`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${appToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ groupIds: [groupId] }),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(graphMessage(data, "Could not verify admin group membership."));
  }
  return Array.isArray(data.value) && data.value.includes(groupId);
}

function parseBody(req) {
  if (req.body == null || req.body === "") return {};
  if (typeof req.body === "object") return req.body;
  try {
    return JSON.parse(req.body);
  } catch {
    throw new Error("Request body must be JSON.");
  }
}

export default async function handler(req, res) {
  res.setHeader("Cache-Control", "no-store");

  if (req.method === "OPTIONS") {
    res.status(204).end();
    return;
  }

  const groupId = process.env.ADMIN_GROUP_ID;
  if (!groupId) {
    json(res, 500, { error: "ADMIN_GROUP_ID is not configured." });
    return;
  }

  try {
    const caller = await callerFromUserToken(req.headers.authorization);
    if (!caller?.id) {
      json(res, 401, { error: "Sign in required." });
      return;
    }

    const appToken = await getAppToken();
    const isAdmin = await isGroupMember(appToken, caller.id, groupId);

    if (req.method === "GET") {
      json(res, 200, { isAdmin, userId: caller.id });
      return;
    }

    if (req.method === "PATCH") {
      if (!isAdmin) {
        json(res, 403, { error: "You are not in the Org Explorer Admins group." });
        return;
      }

      const body = parseBody(req);
      const userId = typeof body.userId === "string" ? body.userId.trim() : "";
      if (!userId) {
        json(res, 400, { error: "userId is required." });
        return;
      }

      const companyName = normalizeField(body.companyName);
      const department = normalizeField(body.department);
      const jobTitle = normalizeField(body.jobTitle);
      const team = normalizeField(body.team);

      const patch = {
        companyName: companyName || null,
        department: department || null,
        jobTitle: jobTitle || null,
        onPremisesExtensionAttributes: {
          extensionAttribute1: team || null,
        },
      };

      const patchRes = await fetch(`${GRAPH}/users/${encodeURIComponent(userId)}`, {
        method: "PATCH",
        headers: {
          Authorization: `Bearer ${appToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(patch),
      });

      if (!patchRes.ok) {
        const err = await patchRes.json().catch(() => ({}));
        let message = graphMessage(err, patchRes.statusText);
        if (/extensionAttribute|onPremises|source of authority/i.test(message)) {
          message += " Graph can only write CustomAttribute1 for cloud-only users. Hybrid or Exchange-mastered mailboxes must be updated in Exchange.";
        }
        json(res, patchRes.status, { error: message });
        return;
      }

      json(res, 200, {
        ok: true,
        userId,
        companyName: companyName || null,
        department: department || null,
        jobTitle: jobTitle || null,
        team: team || null,
      });
      return;
    }

    res.setHeader("Allow", "GET, PATCH");
    json(res, 405, { error: "Method not allowed." });
  } catch (err) {
    json(res, 500, { error: err.message || "Admin API failed." });
  }
}
