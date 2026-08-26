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

function env(name) {
  if (process.env[name]) return process.env[name];
  const found = Object.keys(process.env).find((key) => key.toLowerCase() === name.toLowerCase());
  return found ? process.env[found] : "";
}

function envPresent() {
  return {
    ADMIN_GROUP_ID: Boolean(env("ADMIN_GROUP_ID")),
    ENTRA_CLIENT_ID: Boolean(env("ENTRA_CLIENT_ID")),
    ENTRA_CLIENT_SECRET: Boolean(env("ENTRA_CLIENT_SECRET")),
    ENTRA_TENANT_ID: Boolean(env("ENTRA_TENANT_ID")),
  };
}

function similarEnvKeys() {
  return Object.keys(process.env)
    .filter((key) => /admin|group|entra/i.test(key))
    .sort();
}

async function getAppToken() {
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

async function callerFromUserToken(authHeader) {
  if (!authHeader?.startsWith("Bearer ")) return null;
  const token = authHeader.slice(7);
  const res = await fetch(`${GRAPH}/me?$select=id,displayName,userPrincipalName`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) return null;
  return res.json();
}

async function graphJson(token, path, options = {}) {
  const res = await fetch(`${GRAPH}${path}`, {
    method: options.method || "GET",
    headers: {
      Authorization: `Bearer ${token}`,
      ...(options.body ? { "Content-Type": "application/json" } : {}),
    },
    body: options.body ? JSON.stringify(options.body) : undefined,
  });
  const data = await res.json().catch(() => ({}));
  return { res, data };
}

async function checkMemberGroups(token, path, groupId) {
  const { res, data } = await graphJson(token, path, {
    method: "POST",
    body: { groupIds: [groupId] },
  });
  if (res.status === 401 || res.status === 403) return { result: null, status: res.status };
  if (!res.ok) {
    throw new Error(graphMessage(data, "Could not verify admin group membership."));
  }
  const wanted = normalizeGuid(groupId);
  const hit = (data.value || []).some((id) => normalizeGuid(id) === wanted);
  return { result: hit, status: res.status };
}

async function getMemberGroupsContains(token, path, groupId) {
  const { res, data } = await graphJson(token, path, {
    method: "POST",
    body: { securityEnabledOnly: false },
  });
  if (res.status === 401 || res.status === 403) return { result: null, status: res.status };
  if (!res.ok) return { result: null, status: res.status, error: graphMessage(data, res.statusText) };
  const wanted = normalizeGuid(groupId);
  const hit = (data.value || []).some((id) => normalizeGuid(id) === wanted);
  return { result: hit, status: res.status };
}

async function isDirectMember(token, groupId, userId) {
  const { res } = await graphJson(
    token,
    `/groups/${encodeURIComponent(groupId)}/members/${encodeURIComponent(userId)}?$select=id`
  );
  if (res.status === 404) return { result: false, status: 404 };
  if (res.status === 401 || res.status === 403) return { result: null, status: res.status };
  return { result: res.ok, status: res.status };
}

async function readGroup(token, groupId) {
  const { res, data } = await graphJson(
    token,
    `/groups/${encodeURIComponent(groupId)}?$select=id,displayName,groupTypes,securityEnabled,mailEnabled,visibility`
  );
  if (!res.ok) return null;
  return {
    id: data.id,
    displayName: data.displayName,
    groupTypes: data.groupTypes || [],
    securityEnabled: data.securityEnabled,
    mailEnabled: data.mailEnabled,
    visibility: data.visibility,
  };
}

async function isGroupMember(userToken, appToken, userId, groupId) {
  const checks = {};
  const group = (await readGroup(appToken, groupId)) || (await readGroup(userToken, groupId));

  checks.userCheckMemberGroups = await checkMemberGroups(userToken, "/me/checkMemberGroups", groupId);
  checks.appDirectMember = await isDirectMember(appToken, groupId, userId);
  checks.userGetMemberGroups = await getMemberGroupsContains(userToken, "/me/getMemberGroups", groupId);
  checks.appCheckMemberGroups = await checkMemberGroups(
    appToken,
    `/users/${encodeURIComponent(userId)}/checkMemberGroups`,
    groupId
  );
  checks.appGetMemberGroups = await getMemberGroupsContains(
    appToken,
    `/users/${encodeURIComponent(userId)}/getMemberGroups`,
    groupId
  );

  const isAdmin = Object.values(checks).some((item) => item.result === true);
  const evaluated = Object.values(checks).some((item) => item.result === true || item.result === false);
  if (!isAdmin && !evaluated) {
    throw new Error(
      "Could not evaluate group membership. Add delegated GroupMember.Read.All, grant admin consent, then sign out and sign in again."
    );
  }

  return { isAdmin, checks, group };
}

function normalizeGuid(id) {
  return String(id || "")
    .trim()
    .replace(/^['"]+|['"]+$/g, "")
    .replace(/[{}]/g, "")
    .toLowerCase();
}

async function graphBatch(token, requests) {
  const map = new Map();
  if (!requests.length) return map;
  const res = await fetch(`${GRAPH}/$batch`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ requests }),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(graphMessage(data, "Graph batch failed."));
  }
  (data.responses || []).forEach((item) => map.set(String(item.id), item));
  return map;
}

function batchError(item) {
  if (!item) return "No Graph response.";
  if (item.status >= 200 && item.status < 300) return null;
  let message = graphMessage(item.body, `Graph ${item.status}`);
  if (/extensionAttribute|onPremises|source of authority/i.test(message)) {
    message += " Graph can only write CustomAttribute1 for cloud-only users. Hybrid or Exchange-mastered mailboxes must be updated in Exchange.";
  }
  return message;
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

  const groupId = normalizeGuid(env("ADMIN_GROUP_ID"));
  if (!groupId) {
    json(res, 500, {
      error: "ADMIN_GROUP_ID is not available to this deployment. Add it on the Vercel project that serves this URL, tick Production, then Redeploy. Saving the variable on an old deployment is not enough.",
      vercelEnv: process.env.VERCEL_ENV || null,
      vercelTargetEnv: process.env.VERCEL_TARGET_ENV || null,
      envPresent: envPresent(),
      similarKeys: similarEnvKeys(),
    });
    return;
  }

  try {
    const caller = await callerFromUserToken(req.headers.authorization);
    if (!caller?.id) {
      json(res, 401, { error: "Sign in required." });
      return;
    }

    const userToken = req.headers.authorization.slice(7);
    const appToken = await getAppToken();
    const membership = await isGroupMember(userToken, appToken, caller.id, groupId);
    const isAdmin = membership.isAdmin;

    if (req.method === "GET") {
      json(res, 200, {
        isAdmin,
        userId: caller.id,
        group: membership.group,
        checks: membership.checks,
      });
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

    if (req.method === "POST") {
      if (!isAdmin) {
        json(res, 403, { error: "You are not in the Org Explorer Admins group." });
        return;
      }

      const body = parseBody(req);
      const updates = Array.isArray(body.updates) ? body.updates : null;
      if (!updates) {
        json(res, 400, { error: "updates array is required." });
        return;
      }
      if (updates.length > 20) {
        json(res, 400, { error: "Send at most 20 updates per request." });
        return;
      }

      const queued = [];
      const results = [];
      updates.forEach((item, index) => {
        const userId = typeof item?.userId === "string" ? item.userId.trim() : "";
        if (!userId) {
          results.push({ userId: "", ok: false, error: "userId is required." });
          return;
        }
        queued.push({
          id: String(index + 1),
          userId,
          companyName: normalizeField(item.companyName),
          department: normalizeField(item.department),
          jobTitle: normalizeField(item.jobTitle),
          team: normalizeField(item.team),
        });
      });

      const profileMap = await graphBatch(appToken, queued.map((item) => ({
        id: item.id,
        method: "PATCH",
        url: `/users/${item.userId}`,
        headers: { "Content-Type": "application/json" },
        body: {
          companyName: item.companyName || null,
          department: item.department || null,
          jobTitle: item.jobTitle || null,
        },
      })));
      const teamMap = await graphBatch(appToken, queued.map((item) => ({
        id: item.id,
        method: "PATCH",
        url: `/users/${item.userId}`,
        headers: { "Content-Type": "application/json" },
        body: {
          onPremisesExtensionAttributes: {
            extensionAttribute1: item.team || null,
          },
        },
      })));

      queued.forEach((item) => {
        const profileError = batchError(profileMap.get(item.id));
        const teamError = batchError(teamMap.get(item.id));
        if (profileError && teamError) {
          results.push({ userId: item.userId, ok: false, error: profileError });
          return;
        }
        results.push({
          userId: item.userId,
          ok: true,
          companyName: item.companyName || null,
          department: item.department || null,
          jobTitle: item.jobTitle || null,
          team: item.team || null,
          ...(profileError ? { profileError } : {}),
          ...(teamError ? { teamError } : {}),
        });
      });

      json(res, 200, { ok: true, results });
      return;
    }

    res.setHeader("Allow", "GET, PATCH, POST");
    json(res, 405, { error: "Method not allowed." });
  } catch (err) {
    json(res, 500, { error: err.message || "Admin API failed." });
  }
}
