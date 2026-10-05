import { GRAPH, graphMessage } from "./graph.js";

// Loads every active member from Microsoft 365 with their manager, for /agent and the nightly backup.

const USER_SELECT = [
  "id", "displayName", "givenName", "surname", "mail", "userPrincipalName",
  "jobTitle", "department", "officeLocation", "companyName", "employeeType",
  "onPremisesExtensionAttributes", "accountEnabled", "userType",
].join(",");

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

export async function loadPeople(token) {
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

export function byName(a, b) {
  return a.name.localeCompare(b.name, "en", { sensitivity: "base" });
}

export function buildTree(people, byId) {
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

export function snapshot(people, tree) {
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
