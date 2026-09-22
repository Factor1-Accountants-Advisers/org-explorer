import { createHash, timingSafeEqual } from "node:crypto";
import { GRAPH, env, getAppToken, graphMessage } from "./_lib/graph.js";

// Served at /agent (see vercel.json). Writes the whole org chart as Markdown so
// AI agents can read every person and their reporting line in one fetch.

const USER_SELECT = [
  "id", "displayName", "givenName", "surname", "mail", "userPrincipalName",
  "jobTitle", "department", "officeLocation", "companyName", "employeeType",
  "onPremisesExtensionAttributes", "accountEnabled", "userType",
].join(",");

function sameSecret(a, b) {
  const ha = createHash("sha256").update(String(a)).digest();
  const hb = createHash("sha256").update(String(b)).digest();
  return timingSafeEqual(ha, hb);
}

function providedKey(req) {
  const auth = req.headers.authorization || "";
  if (auth.startsWith("Bearer ")) return auth.slice(7).trim();
  const key = req.query?.key;
  return typeof key === "string" ? key.trim() : "";
}

async function graphGet(token, path) {
  const res = await fetch(path.startsWith("http") ? path : `${GRAPH}${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(graphMessage(data, `Graph request failed (${res.status}).`));
  return data;
}

function clean(value) {
  return value == null ? "" : String(value).trim();
}

function toPerson(u) {
  const ext = u.onPremisesExtensionAttributes || {};
  return {
    id: u.id,
    name: clean(u.displayName),
    givenName: clean(u.givenName),
    surname: clean(u.surname),
    email: clean(u.mail) || clean(u.userPrincipalName),
    userPrincipalName: clean(u.userPrincipalName),
    company: clean(u.companyName),
    department: clean(u.department),
    team: clean(ext.extensionAttribute1),
    role: clean(u.jobTitle),
    location: clean(ext.extensionAttribute2),
    officeLocation: clean(u.officeLocation),
    employeeType: clean(u.employeeType),
    managerId: u.manager?.id || null,
  };
}

async function loadPeople(token) {
  const people = [];
  let path = `/users?$select=${USER_SELECT}&$expand=manager($select=id)&$top=999`;
  while (path) {
    const data = await graphGet(token, path);
    (data.value || []).forEach((u) => {
      const guest = (u.userType || "").toLowerCase() === "guest";
      if (u.id && u.accountEnabled !== false && !guest && clean(u.displayName)) {
        people.push(toPerson(u));
      }
    });
    path = data["@odata.nextLink"] || "";
  }

  const byId = new Map(people.map((p) => [p.id, p]));
  // A manager outside the listed people (disabled, guest) leaves the person at the top.
  people.forEach((p) => {
    if (p.managerId && !byId.has(p.managerId)) p.managerId = null;
  });
  return { people, byId };
}

function byName(a, b) {
  return a.name.localeCompare(b.name, "en", { sensitivity: "base" });
}

function buildTree(people, byId) {
  const children = new Map();
  people.forEach((p) => {
    if (!p.managerId) return;
    if (!children.has(p.managerId)) children.set(p.managerId, []);
    children.get(p.managerId).push(p);
  });
  children.forEach((list) => list.sort(byName));

  const hasOrgDetails = (p) => p.role || p.department || p.company || children.has(p.id);
  const tops = people.filter((p) => !p.managerId).sort(byName);
  const roots = tops.filter(hasOrgDetails);
  const unplaced = tops.filter((p) => !hasOrgDetails(p));

  // People caught in a manager loop never reach a root; surface them as extra roots.
  const reached = new Set();
  const mark = (p) => {
    if (reached.has(p.id)) return;
    reached.add(p.id);
    (children.get(p.id) || []).forEach(mark);
  };
  tops.forEach(mark);
  const looped = people.filter((p) => !reached.has(p.id)).sort(byName);
  looped.forEach((p) => {
    if (reached.has(p.id)) return;
    roots.push(p);
    mark(p);
  });

  return { children, roots, unplaced, byId };
}

function summary(p) {
  const bits = [p.role, [p.company, p.department, p.team].filter(Boolean).join(" / ")].filter(Boolean);
  return bits.length ? `${p.name} — ${bits.join(" · ")}` : p.name;
}

function treeLines(tree) {
  const lines = [];
  const seen = new Set();
  const walk = (p, depth) => {
    const indent = "  ".repeat(depth);
    if (seen.has(p.id)) {
      lines.push(`${indent}- ${p.name} (reporting loop — see People)`);
      return;
    }
    seen.add(p.id);
    lines.push(`${indent}- ${summary(p)}`);
    (tree.children.get(p.id) || []).forEach((c) => walk(c, depth + 1));
  };
  tree.roots.forEach((p) => walk(p, 0));
  return lines;
}

function personBlock(p, tree) {
  const manager = p.managerId ? tree.byId.get(p.managerId) : null;
  const reports = tree.children.get(p.id) || [];
  const field = (label, value) => (value ? `- ${label}: ${value}` : null);
  return [
    `### ${p.name}`,
    field("Role", p.role),
    field("Company", p.company),
    field("Department", p.department),
    field("Team", p.team),
    field("Location", p.location),
    field("Office", p.officeLocation && p.officeLocation !== p.location ? p.officeLocation : ""),
    field("Employee type", p.employeeType),
    field("Email", p.email),
    field("UPN", p.userPrincipalName && p.userPrincipalName !== p.email ? p.userPrincipalName : ""),
    `- Reports to: ${manager ? `${manager.name}${manager.email ? ` <${manager.email}>` : ""}` : "nobody (top of the org)"}`,
    `- Direct reports (${reports.length}): ${reports.length ? reports.map((r) => r.name).join(", ") : "none"}`,
    `- ID: ${p.id}`,
  ].filter(Boolean).join("\n");
}

function renderMarkdown(people, tree) {
  const now = new Date().toISOString();
  const sorted = [...people].sort(byName);
  const out = [
    "# Factor1 Org Chart",
    "",
    `Generated ${now} from Microsoft 365 (Entra ID). ${people.length} active people.`,
    "",
    "Hierarchy is Company → Department → Team → Role. Reporting lines are the Microsoft 365 manager.",
    "Team is Exchange CustomAttribute1 and Location is CustomAttribute2.",
    "Machine-readable version: add `format=json` to this URL.",
    "",
    "## Reporting tree",
    "",
    "Each line is `Name — Role · Company / Department / Team`. Indented people report to the line above them.",
    "",
    ...treeLines(tree),
    "",
  ];

  if (tree.unplaced.length) {
    out.push(
      "## Accounts with no role or reporting line",
      "",
      "Active accounts with no manager, no direct reports, and no role, department or company. Often shared or service accounts.",
      "",
      ...tree.unplaced.map((p) => `- ${p.name}${p.email ? ` <${p.email}>` : ""}`),
      ""
    );
  }

  out.push("## People", "", "Everyone, A–Z, with full details.", "");
  sorted.forEach((p) => out.push(personBlock(p, tree), ""));
  return out.join("\n");
}

function renderJson(people, tree) {
  return {
    generatedAt: new Date().toISOString(),
    count: people.length,
    people: [...people].sort(byName).map((p) => ({
      ...p,
      managerName: p.managerId ? tree.byId.get(p.managerId)?.name || null : null,
      directReportIds: (tree.children.get(p.id) || []).map((r) => r.id),
    })),
  };
}

export default async function handler(req, res) {
  res.setHeader("Cache-Control", "no-store");
  res.setHeader("X-Robots-Tag", "noindex, nofollow");

  if (req.method !== "GET" && req.method !== "HEAD") {
    res.setHeader("Allow", "GET, HEAD");
    res.status(405).send("Method not allowed.\n");
    return;
  }

  const accessKey = env("AGENT_ACCESS_KEY");
  if (!accessKey) {
    res.status(503).setHeader("Content-Type", "text/plain; charset=utf-8");
    res.send("The agent view is not configured. Set AGENT_ACCESS_KEY on the Vercel project, then redeploy.\n");
    return;
  }
  if (!sameSecret(providedKey(req), accessKey)) {
    res.status(401).setHeader("Content-Type", "text/plain; charset=utf-8");
    res.send("Access key required. Use /agent?key=<key> or send Authorization: Bearer <key>.\n");
    return;
  }

  try {
    const token = await getAppToken();
    const { people, byId } = await loadPeople(token);
    const tree = buildTree(people, byId);

    if (req.query?.format === "json") {
      res.status(200).json(renderJson(people, tree));
      return;
    }
    res.status(200).setHeader("Content-Type", "text/markdown; charset=utf-8");
    res.send(renderMarkdown(people, tree));
  } catch (err) {
    res.status(500).setHeader("Content-Type", "text/plain; charset=utf-8");
    res.send(`Could not load the org chart: ${err.message || "unknown error"}\n`);
  }
}
