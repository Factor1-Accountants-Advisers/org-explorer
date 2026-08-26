const CONFIG = {
  clientId: "3cc7e474-6af5-4f3e-8c5f-607c5cfc11b8",
  tenantId: "factor1.com.au",
  redirectUri: window.location.origin + "/",
      scopes: ["User.Read", "User.Read.All", "GroupMember.Read.All"],
  graphBase: "https://graph.microsoft.com/v1.0",
};

const USER_SELECT = [
  "id", "displayName", "givenName", "surname", "mail", "userPrincipalName",
  "jobTitle", "department", "officeLocation", "companyName", "employeeType",
  "onPremisesExtensionAttributes",
].join(",");

const REPORT_SELECT = USER_SELECT;

const DEMO_USERS = {
  "822b4129-2de9-4f18-a319-c336e8366bf4": {
    id: "822b4129-2de9-4f18-a319-c336e8366bf4",
    displayName: "Benjamin Bryant",
    givenName: "Benjamin", surname: "Bryant",
    mail: "BenjaminBryant@factor1.com.au",
    userPrincipalName: "BenjaminBryant@factor1.com.au",
    jobTitle: "Chief Innovation Officer",
    department: "Innovation & Systems",
    team: "Innovation Leadership",
    companyName: "Factor1 Group",
    employeeType: "Melbourne",
    managerId: null,
  },
  "05295543-6a39-479c-a5d9-6fc9db95e1ed": {
    id: "05295543-6a39-479c-a5d9-6fc9db95e1ed",
    displayName: "David Ahlhaus",
    givenName: "David", surname: "Ahlhaus",
    mail: "DavidAhlhaus@factor1.com.au",
    userPrincipalName: "DavidAhlhaus@factor1.com.au",
    jobTitle: "Innovation and Systems Lead",
    department: "Innovation & Systems",
    team: "Systems",
    companyName: "Factor1 Group",
    employeeType: "Melbourne",
    managerId: "822b4129-2de9-4f18-a319-c336e8366bf4",
  },
  "a1000001-0000-4000-8000-000000000001": {
    id: "a1000001-0000-4000-8000-000000000001",
    displayName: "Alex Chen",
    givenName: "Alex", surname: "Chen",
    mail: "AlexChen@factor1.com.au",
    userPrincipalName: "AlexChen@factor1.com.au",
    jobTitle: "Systems Analyst",
    department: "Innovation & Systems",
    team: "Systems",
    companyName: "Factor1 Group",
    employeeType: "Melbourne",
    managerId: "05295543-6a39-479c-a5d9-6fc9db95e1ed",
  },
  "a1000002-0000-4000-8000-000000000002": {
    id: "a1000002-0000-4000-8000-000000000002",
    displayName: "Sam Rivera",
    givenName: "Sam", surname: "Rivera",
    mail: "SamRivera@factor1.com.au",
    userPrincipalName: "SamRivera@factor1.com.au",
    jobTitle: "Automation Specialist",
    department: "Innovation & Systems",
    team: "Automation",
    companyName: "Factor1 Group",
    employeeType: "Brisbane",
    managerId: "05295543-6a39-479c-a5d9-6fc9db95e1ed",
  },
  "a1000003-0000-4000-8000-000000000003": {
    id: "a1000003-0000-4000-8000-000000000003",
    displayName: "Jordan Lee",
    givenName: "Jordan", surname: "Lee",
    mail: "JordanLee@factor1.com.au",
    userPrincipalName: "JordanLee@factor1.com.au",
    jobTitle: "Data Engineer",
    department: "Innovation & Systems",
    team: "Data",
    companyName: "Factor1 Group",
    employeeType: "Melbourne",
    managerId: "05295543-6a39-479c-a5d9-6fc9db95e1ed",
  },
  "b2000001-0000-4000-8000-000000000001": {
    id: "b2000001-0000-4000-8000-000000000001",
    displayName: "Taylor Morgan",
    givenName: "Taylor", surname: "Morgan",
    mail: "TaylorMorgan@factor1.com.au",
    userPrincipalName: "TaylorMorgan@factor1.com.au",
    jobTitle: "HR Business Partner",
    department: "People & Culture",
    team: "People Partners",
    companyName: "Factor1 Group",
    employeeType: "Melbourne",
    managerId: "822b4129-2de9-4f18-a319-c336e8366bf4",
  },
  "c3000001-0000-4000-8000-000000000001": {
    id: "c3000001-0000-4000-8000-000000000001",
    displayName: "Terry Chung",
    givenName: "Terry", surname: "Chung",
    mail: "TerryChung@factor1.com.au",
    userPrincipalName: "TerryChung@factor1.com.au",
    jobTitle: "Managing Director",
    department: "Firm Management",
    team: "Firm Office",
    companyName: "Factor1 Group",
    employeeType: "Melbourne",
    managerId: null,
  },
};

const params = new URLSearchParams(window.location.search);
// Live mode by default once Client ID is set. Use ?demo=1 for sample data.
const isDemoMode = () => params.get("demo") === "1";

let msalInstance = null;
let account = null;
let myUserId = null;
let branchStack = [];
let demoMode = false;
let isAdmin = false;
let companyFilter = "";
let departmentFilter = "";
let teamFilter = "";
let knownCompanies = new Set();
const departmentsByCompany = new Map();
const teamsByCompanyDept = new Map();
const knownRoles = new Set();
let editingUserId = null;
let searchTimer = null;
let animDirection = null;
let drawFrame = 0;
const photoCache = new Map();
const userCache = new Map();
const directReportCountCache = new Map();
let orgRootsCache = null;
let orgRootsPromise = null;
let fullOrgPeople = null;
let fullOrgManagerMap = null;
let fullOrgLoadPromise = null;
let ftMode = "employees"; // employees | structure
let ftPan = { x: 0, y: 0, scale: 1 };
let ftDragging = false;
let ftDragStart = null;
let ftFocusNodeId = null;
let ftLayoutBounds = { width: 0, height: 0 };
let ftPanZoomBound = false;

const FT = {
  PERSON_W: 132,
  PERSON_H: 78,
  COMPANY_W: 170,
  COMPANY_H: 56,
  DEPT_W: 150,
  DEPT_H: 50,
  TEAM_W: 150,
  TEAM_H: 50,
  JOB_W: 158,
  JOB_H: 50,
  H_GAP: 18,
  V_GAP: 88,
  MIN_SCALE: 0.18,
  MAX_SCALE: 2.4,
};

const els = {
  status: document.getElementById("status"),
  mainUi: document.getElementById("main-ui"),
  branchView: document.getElementById("branch-view"),
  fullTreeUi: document.getElementById("full-tree-ui"),
  treeViewport: document.getElementById("tree-viewport"),
  treeInner: document.getElementById("tree-inner"),
  branchSvg: document.getElementById("branch-svg"),
  pathSpine: document.getElementById("path-spine"),
  searchInput: document.getElementById("search-input"),
  searchResults: document.getElementById("search-results"),
  companyFilter: document.getElementById("company-filter"),
  deptFilter: document.getElementById("dept-filter"),
  teamFilter: document.getElementById("team-filter"),
  setupNote: document.getElementById("setup-note"),
  btnSignIn: document.getElementById("btn-signin"),
  btnSignOut: document.getElementById("btn-signout"),
  btnHome: document.getElementById("btn-home"),
  btnFullTree: document.getElementById("btn-full-tree"),
  btnAdmin: document.getElementById("btn-admin"),
  btnAdminBack: document.getElementById("btn-admin-back"),
  btnAdminCsv: document.getElementById("btn-admin-csv"),
  btnAdminCsvApply: document.getElementById("btn-admin-csv-apply"),
  adminCsvFile: document.getElementById("admin-csv-file"),
  adminUi: document.getElementById("admin-ui"),
  adminSearchInput: document.getElementById("admin-search-input"),
  adminResults: document.getElementById("admin-results"),
  editDrawer: document.getElementById("edit-drawer"),
  editDrawerPerson: document.getElementById("edit-drawer-person"),
  editForm: document.getElementById("edit-form"),
  editCompany: document.getElementById("edit-company"),
  editDepartment: document.getElementById("edit-department"),
  editTeam: document.getElementById("edit-team"),
  editRole: document.getElementById("edit-role"),
  editMessage: document.getElementById("edit-message"),
  btnEditSave: document.getElementById("btn-edit-save"),
  btnFtBack: document.getElementById("btn-ft-back"),
  btnFtEmployees: document.getElementById("btn-ft-employees"),
  btnFtStructure: document.getElementById("btn-ft-structure"),
  btnFtFocus: document.getElementById("btn-ft-focus"),
  ftViewport: document.getElementById("ft-viewport"),
  ftStage: document.getElementById("ft-stage"),
  ftSvg: document.getElementById("ft-svg"),
  ftNodes: document.getElementById("ft-nodes"),
  ftLoading: document.getElementById("ft-loading"),
  demoBanner: document.getElementById("demo-banner"),
  btnLive: document.getElementById("btn-live"),
};

// ── MSAL loader (alcdn URL 404s; use jsDelivr with fallback) ───────────────
function loadMsalLibrary() {
  if (window.msal) return Promise.resolve(window.msal);

  const sources = [
    "https://cdn.jsdelivr.net/npm/@azure/msal-browser@2.38.3/lib/msal-browser.min.js",
    "https://unpkg.com/@azure/msal-browser@2.38.3/lib/msal-browser.min.js",
  ];

  return new Promise((resolve, reject) => {
    let index = 0;

    const tryNext = () => {
      if (window.msal) return resolve(window.msal);
      if (index >= sources.length) {
        reject(new Error("Could not load Microsoft sign-in library. Check your network or ad blocker."));
        return;
      }

      const script = document.createElement("script");
      script.src = sources[index++];
      script.async = true;
      script.onload = () => {
        if (window.msal) resolve(window.msal);
        else tryNext();
      };
      script.onerror = tryNext;
      document.head.appendChild(script);
    };

    tryNext();
  });
}

// ── Auth ────────────────────────────────────────────────────────────────────
function getMsalConfig() {
  return {
    auth: {
      clientId: CONFIG.clientId,
      authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
      redirectUri: CONFIG.redirectUri,
      navigateToLoginRequestUrl: false,
    },
    cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: true },
  };
}

async function ensureMsal() {
  await loadMsalLibrary();
  if (!msalInstance) {
    msalInstance = new msal.PublicClientApplication(getMsalConfig());
    await msalInstance.initialize();
  }
}

async function initAuth() {
  await ensureMsal();
  const redirectResult = await msalInstance.handleRedirectPromise();
  account = redirectResult?.account || msalInstance.getAllAccounts()[0] || null;
  updateAuthUI();
  return !!account;
}

async function signIn() {
  await ensureMsal();
  await msalInstance.loginRedirect({ scopes: CONFIG.scopes });
}
async function signOut() {
  await ensureMsal();
  await msalInstance.logoutRedirect();
}

async function getToken() {
  if (!account) throw new Error("Not signed in");
  try {
    return (await msalInstance.acquireTokenSilent({ scopes: CONFIG.scopes, account })).accessToken;
  } catch {
    await msalInstance.acquireTokenRedirect({ scopes: CONFIG.scopes, account });
    return null;
  }
}

function updateAuthUI() {
  const signedIn = !!account;
  els.btnSignIn.classList.toggle("hidden", signedIn);
  els.btnSignOut.classList.toggle("hidden", !signedIn);
  els.btnHome.classList.toggle("hidden", !signedIn && !demoMode);
}

function showSetupNote() {
  els.setupNote.classList.remove("hidden");
  els.setupNote.innerHTML = `
    <strong>Preview / live</strong><br>
    Sample data: <a href="?demo=1">?demo=1</a> · Live Entra sign-in: <a href="./">open without demo</a><br>
    Local redirect URI must be exactly: <code>${CONFIG.redirectUri}</code>
  `;
}

// ── Data ────────────────────────────────────────────────────────────────────
function stripInternal(raw) {
  const { managerId, ...user } = raw;
  return user;
}

function demoRaw(userId) {
  const id = userId === "me" ? myUserId : userId;
  return DEMO_USERS[id];
}

function userTeam(user) {
  if (user?.team?.trim()) return user.team.trim();
  return user?.onPremisesExtensionAttributes?.extensionAttribute1?.trim() || "";
}

function orgKey(company, dept) {
  return `${company || ""}\0${dept || ""}`;
}

function cacheUser(user) {
  if (!user?.id) return;
  const team = userTeam(user);
  if (team) user.team = team;
  userCache.set(user.id, user);
}

function rememberOrgFields(user) {
  const team = userTeam(user);
  const company = user.companyName?.trim();
  const dept = user.department?.trim();
  const role = user.jobTitle?.trim();
  if (role) knownRoles.add(role);
  if (company) {
    knownCompanies.add(company);
    if (!departmentsByCompany.has(company)) departmentsByCompany.set(company, new Set());
    if (dept) {
      departmentsByCompany.get(company).add(dept);
      const key = orgKey(company, dept);
      if (!teamsByCompanyDept.has(key)) teamsByCompanyDept.set(key, new Set());
      if (team) teamsByCompanyDept.get(key).add(team);
    }
  }
}

function rebuildOrgCatalogs() {
  knownCompanies = new Set();
  departmentsByCompany.clear();
  teamsByCompanyDept.clear();
  knownRoles.clear();
  userCache.forEach(rememberOrgFields);
}

async function fetchUser(userId) {
  if (demoMode) {
    const raw = demoRaw(userId);
    if (!raw) throw new Error("Person not found.");
    const user = stripInternal(raw);
    cacheUser(user);
    return user;
  }
  if (userCache.has(userId)) return userCache.get(userId);
  const path = userId === "me" ? `/me?$select=${USER_SELECT}` : `/users/${userId}?$select=${USER_SELECT}`;
  const user = await graphGet(path);
  cacheUser(user);
  return user;
}

async function fetchManager(userId) {
  if (demoMode) {
    const raw = demoRaw(userId);
    if (!raw?.managerId) return null;
    return stripInternal(DEMO_USERS[raw.managerId]);
  }
  const base = userId === "me" ? "/me" : `/users/${userId}`;
  try {
    const mgr = await graphGet(`${base}/manager?$select=${REPORT_SELECT}`);
    if (mgr) cacheUser(mgr);
    return mgr;
  } catch { return null; }
}

async function fetchDirectReports(userId) {
  if (demoMode) {
    return sortPeople(Object.values(DEMO_USERS)
      .filter((u) => u.managerId === userId)
      .map(stripInternal));
  }
  const base = userId === "me" ? "/me" : `/users/${userId}`;
  const data = await graphGet(`${base}/directReports?$select=${REPORT_SELECT}&$top=100`);
  const reports = (data?.value || []).filter((u) => u.id);
  reports.forEach(cacheUser);
  return sortPeople(reports);
}

async function countDirectReports(userId) {
  if (directReportCountCache.has(userId)) return directReportCountCache.get(userId);
  if (demoMode) {
    const count = Object.values(DEMO_USERS).filter((u) => u.managerId === userId).length;
    directReportCountCache.set(userId, count);
    return count;
  }
  const base = userId === "me" ? "/me" : `/users/${userId}`;
  try {
    const data = await graphGet(`${base}/directReports?$select=id&$top=999`);
    const count = (data?.value || []).filter((u) => u.id).length;
    directReportCountCache.set(userId, count);
    return count;
  } catch {
    return 0;
  }
}

function sortPeople(users) {
  return [...users].sort((a, b) => {
    const companyA = (a.companyName || "").trim().toLowerCase();
    const companyB = (b.companyName || "").trim().toLowerCase();
    if (companyA !== companyB) {
      if (!companyA) return 1;
      if (!companyB) return -1;
      return companyA.localeCompare(companyB);
    }
    const deptA = (a.department || "").trim().toLowerCase();
    const deptB = (b.department || "").trim().toLowerCase();
    if (deptA !== deptB) {
      if (!deptA) return 1;
      if (!deptB) return -1;
      return deptA.localeCompare(deptB);
    }
    const teamA = userTeam(a).toLowerCase();
    const teamB = userTeam(b).toLowerCase();
    if (teamA !== teamB) {
      if (!teamA) return 1;
      if (!teamB) return -1;
      return teamA.localeCompare(teamB);
    }
    return (a.displayName || "").localeCompare(b.displayName || "", undefined, { sensitivity: "base" });
  });
}

function companyBadgeClass(company) {
  const name = (company || "").trim().toLowerCase();
  if (!name) return "";

  // Named brands (check specific Factor1 entities before the Group match)
  if (name.includes("shpp")) return "company-blue";
  if (name.includes("taxopia")) return "company-dark-green";
  if (name.includes("jmr")) return "company-light-green";
  if (name.includes("kp partner")) return "company-purple";
  if (name === "factor1 group" || name === "factor1") return "company-factor1";

  const palette = [
    "company-teal", "company-blue", "company-light-green", "company-purple",
    "company-slate", "company-rose", "company-amber", "company-cyan",
  ];
  let hash = 0;
  for (let i = 0; i < name.length; i++) hash = ((hash << 5) - hash) + name.charCodeAt(i);
  return palette[Math.abs(hash) % palette.length];
}

function orgTagsHtml(user) {
  const company = user.companyName?.trim();
  const dept = user.department?.trim();
  const team = userTeam(user);
  if (!company && !dept && !team) return "";
  const tags = [];
  if (company) {
    const tone = companyBadgeClass(company);
    tags.push(`<span class="node-tag node-company ${tone}">${escapeHtml(company)}</span>`);
  }
  if (dept) tags.push(`<span class="node-tag node-dept">${escapeHtml(dept)}</span>`);
  if (team) tags.push(`<span class="node-tag node-team">${escapeHtml(team)}</span>`);
  return `<div class="node-tags">${tags.join("")}</div>`;
}

async function graphGet(path) {
  const token = await getToken();
  if (!token) return null;
  const res = await fetch(`${CONFIG.graphBase}${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    if (res.status === 403) {
      throw new Error("Access denied — admin must grant User.Read.All on the app.");
    }
    throw new Error(err.error?.message || res.statusText);
  }
  return res.json();
}

let directoryPeople = null;
let directoryPeoplePromise = null;

async function fetchDirectoryPeople() {
  if (directoryPeople) return directoryPeople;
  if (directoryPeoplePromise) return directoryPeoplePromise;

  directoryPeoplePromise = (async () => {
    if (demoMode) {
      directoryPeople = sortPeople(Object.values(DEMO_USERS).map((raw) => {
        const user = stripInternal(raw);
        cacheUser(user);
        return user;
      }));
      return directoryPeople;
    }

    const users = [];
    let path = `/users?$select=${USER_SELECT},accountEnabled,userType&$top=999`;
    while (path) {
      const data = await graphGet(path);
      (data?.value || []).forEach((u) => {
        const guest = (u.userType || "").toLowerCase() === "guest";
        if (u.id && u.accountEnabled !== false && !guest && (u.displayName || "").trim()) {
          cacheUser(u);
          users.push(u);
        }
      });
      const next = data?.["@odata.nextLink"] || "";
      if (!next) break;
      path = next.startsWith(CONFIG.graphBase) ? next.slice(CONFIG.graphBase.length) : new URL(next).pathname + new URL(next).search;
    }
    directoryPeople = sortPeople(users);
    return directoryPeople;
  })().finally(() => {
    directoryPeoplePromise = null;
  });

  return directoryPeoplePromise;
}

async function graphBatch(requests) {
  const token = await getToken();
  if (!token || !requests.length) return [];
  const chunks = [];
  for (let i = 0; i < requests.length; i += 20) chunks.push(requests.slice(i, i + 20));
  const responses = [];
  for (const chunk of chunks) {
    const res = await fetch(`${CONFIG.graphBase}/$batch`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ requests: chunk }),
    });
    if (!res.ok) continue;
    const data = await res.json();
    responses.push(...(data.responses || []));
  }
  return responses;
}

async function fetchOrgRoots() {
  if (orgRootsCache) return orgRootsCache;
  if (orgRootsPromise) return orgRootsPromise;

  orgRootsPromise = (async () => {
    if (demoMode) {
      orgRootsCache = sortPeople(
        Object.values(DEMO_USERS).filter((u) => !u.managerId).map(stripInternal)
      );
      return orgRootsCache;
    }

    try {
      const data = await graphGet(`/users?$select=${USER_SELECT},accountEnabled&$top=999`);
      const users = (data?.value || []).filter((u) =>
        u.id && u.accountEnabled !== false && (u.displayName || "").trim()
      );
      users.forEach(cacheUser);

      const batchReqs = users.map((u, i) => ({
        id: String(i),
        method: "GET",
        url: `/users/${u.id}/manager?$select=id`,
      }));
      const batchRes = await graphBatch(batchReqs);
      const hasManager = new Set();
      batchRes.forEach((r) => {
        const body = typeof r.body === "string" ? JSON.parse(r.body || "{}") : (r.body || {});
        if (r.status >= 200 && r.status < 300 && body.id) {
          const user = users[Number(r.id)];
          if (user?.id) hasManager.add(user.id);
        }
      });

      orgRootsCache = sortPeople(users.filter((u) =>
        !hasManager.has(u.id) && (u.jobTitle || u.department || u.companyName)
      ));
      return orgRootsCache;
    } catch {
      orgRootsCache = [];
      return orgRootsCache;
    } finally {
      orgRootsPromise = null;
    }
  })();

  return orgRootsPromise;
}

async function graphSearch(query) {
  if (demoMode) {
    const q = query.toLowerCase();
    return Object.values(DEMO_USERS)
      .map(stripInternal)
      .filter((u) => u.displayName.toLowerCase().includes(q))
      .slice(0, 12);
  }
  const token = await getToken();
  if (!token) return [];
  const escaped = query.replace(/"/g, '\\"');
  const res = await fetch(
    `${CONFIG.graphBase}/users?$search="displayName:${escaped}"&$select=${REPORT_SELECT}&$top=15`,
    {
      headers: {
        Authorization: `Bearer ${token}`,
        ConsistencyLevel: "eventual",
      },
    }
  );
  if (!res.ok) return [];
  const data = await res.json();
  const users = data.value || [];
  users.forEach(cacheUser);
  return users;
}

async function buildBranchTo(userId) {
  const chain = [];
  let id = userId;
  const seen = new Set();
  while (id && !seen.has(id)) {
    seen.add(id);
    const user = await fetchUser(id);
    chain.unshift(user);
    const mgr = await fetchManager(id);
    id = mgr?.id || null;
  }
  return chain;
}

async function fetchPhotoUrl(userId) {
  if (demoMode) return null;
  if (photoCache.has(userId)) return photoCache.get(userId);
  const token = await getToken();
  if (!token) return null;
  const base = userId === myUserId ? "/me" : `/users/${userId}`;
  try {
    const res = await fetch(`${CONFIG.graphBase}${base}/photo/$value`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!res.ok) { photoCache.set(userId, null); return null; }
    const url = URL.createObjectURL(await res.blob());
    photoCache.set(userId, url);
    return url;
  } catch { return null; }
}

function orgMatches(user) {
  const company = (user.companyName || "").trim();
  const dept = (user.department || "").trim();
  const team = userTeam(user);
  if (companyFilter && company !== companyFilter) return false;
  if (departmentFilter && dept !== departmentFilter) return false;
  if (teamFilter && team !== teamFilter) return false;
  return true;
}

function filterByOrg(users) {
  return users.filter(orgMatches);
}

function departmentsForCompany(company) {
  if (company && departmentsByCompany.has(company)) {
    return [...departmentsByCompany.get(company)].sort((a, b) => a.localeCompare(b));
  }
  if (company) return [];
  const all = new Set();
  departmentsByCompany.forEach((depts) => depts.forEach((d) => all.add(d)));
  return [...all].sort((a, b) => a.localeCompare(b));
}

function teamsForFilters() {
  const all = new Set();
  teamsByCompanyDept.forEach((teams, key) => {
    const [company, dept] = key.split("\0");
    if (companyFilter && company !== companyFilter) return;
    if (departmentFilter && dept !== departmentFilter) return;
    teams.forEach((t) => all.add(t));
  });
  return [...all].sort((a, b) => a.localeCompare(b));
}

function refreshCompanyDropdown() {
  const current = els.companyFilter.value;
  const companies = [...knownCompanies].sort((a, b) => a.localeCompare(b));
  els.companyFilter.innerHTML = `<option value="">All companies</option>` +
    companies.map((c) => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join("");
  if (current && knownCompanies.has(current)) {
    els.companyFilter.value = current;
  } else {
    els.companyFilter.value = "";
    companyFilter = "";
  }
}

function refreshDeptDropdown() {
  const current = els.deptFilter.value;
  const depts = departmentsForCompany(companyFilter);
  els.deptFilter.innerHTML = `<option value="">All departments</option>` +
    depts.map((d) => `<option value="${escapeHtml(d)}">${escapeHtml(d)}</option>`).join("");
  if (current && depts.includes(current)) {
    els.deptFilter.value = current;
  } else {
    els.deptFilter.value = "";
    departmentFilter = "";
  }
}

function refreshTeamDropdown() {
  const current = els.teamFilter.value;
  const teams = teamsForFilters();
  els.teamFilter.innerHTML = `<option value="">All teams</option>` +
    teams.map((t) => `<option value="${escapeHtml(t)}">${escapeHtml(t)}</option>`).join("");
  if (current && teams.includes(current)) {
    els.teamFilter.value = current;
  } else {
    els.teamFilter.value = "";
    teamFilter = "";
  }
}

function refreshOrgFilters() {
  rebuildOrgCatalogs();
  refreshCompanyDropdown();
  refreshDeptDropdown();
  refreshTeamDropdown();
}

// ── UI helpers ──────────────────────────────────────────────────────────────
function escapeHtml(str) {
  return String(str ?? "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function initials(p) {
  return ((p.givenName?.[0] || "") + (p.surname?.[0] || "")).toUpperCase()
    || p.displayName?.[0]?.toUpperCase() || "?";
}

function setStatus(msg, isError = false) {
  els.status.textContent = msg;
  els.status.className = isError ? "status error" : "status";
  els.status.classList.toggle("hidden", !msg);
}

function seedDemoOrgFilters() {
  Object.values(DEMO_USERS).forEach((u) => cacheUser(stripInternal(u)));
  refreshOrgFilters();
}

async function nodeHtml(user, role, reportCount = 0) {
  const photo = await fetchPhotoUrl(user.id);
  const avatar = photo
    ? `<img class="avatar" src="${photo}" alt="" />`
    : `<div class="avatar">${escapeHtml(initials(user))}</div>`;
  const hintText = role === "manager" ? "↑ Go up"
    : role === "peer" ? "View branch"
    : role === "report" && reportCount > 0
      ? `${reportCount} direct report${reportCount === 1 ? "" : "s"}`
    : "";
  const hintHtml = role === "report" || role === "peer" || hintText
    ? `<span class="node-hint">${hintText ? escapeHtml(hintText) : ""}</span>`
    : "";
  const editBtn = isAdmin
    ? `<button type="button" class="node-edit" data-edit-id="${user.id}" title="Edit details" aria-label="Edit ${escapeHtml(user.displayName)}">✎</button>`
    : "";

  const cardInner = `
    ${avatar}
    <span class="node-name">${escapeHtml(user.displayName)}</span>
    <span class="node-title">${escapeHtml(user.jobTitle || "—")}</span>
    ${orgTagsHtml(user)}
    ${hintHtml}
  `;

  if (role === "current") {
    return `<div class="tree-node current" data-node-role="current" data-user-id="${user.id}">
      <div class="node-card">${cardInner}</div>
      ${editBtn}
    </div>`;
  }

  return `
    <div class="tree-node ${role}" data-node-role="${role}" data-user-id="${user.id}">
      <button type="button" class="node-card" data-nav-id="${user.id}" data-nav-role="${role}">
        ${cardInner}
      </button>
      ${editBtn}
    </div>`;
}

function curvePath(x1, y1, x2, y2) {
  const midY = (y1 + y2) / 2;
  return `M ${x1.toFixed(1)} ${y1.toFixed(1)} C ${x1.toFixed(1)} ${midY.toFixed(1)}, ${x2.toFixed(1)} ${midY.toFixed(1)}, ${x2.toFixed(1)} ${y2.toFixed(1)}`;
}

function cardAnchor(treeNode, edge, vRect) {
  const card = treeNode.querySelector(".node-card") || treeNode;
  const r = card.getBoundingClientRect();
  return {
    x: r.left - vRect.left + r.width / 2,
    y: edge === "bottom" ? r.bottom - vRect.top : r.top - vRect.top,
  };
}

function topRowNodes(nodes) {
  if (!nodes.length) return nodes;
  const tops = nodes.map((n) => n.getBoundingClientRect().top);
  const minTop = Math.min(...tops);
  return nodes.filter((n, i) => Math.abs(tops[i] - minTop) <= 10);
}

function drawBranchLines() {
  const viewport = els.treeViewport;
  const svg = els.branchSvg;
  const inner = els.treeInner;
  if (!viewport || !svg || !inner) return;

  if (inner.classList.contains("anim-up") || inner.classList.contains("anim-down")) {
    return;
  }

  const vRect = viewport.getBoundingClientRect();
  const width = Math.max(viewport.clientWidth, 1);
  const height = Math.max(viewport.scrollHeight, viewport.clientHeight, 1);
  svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
  svg.setAttribute("width", String(width));
  svg.setAttribute("height", String(height));
  svg.style.width = `${width}px`;
  svg.style.height = `${height}px`;

  const paths = [];
  const managerNode = viewport.querySelector('[data-node-role="manager"]');
  const currentNode = viewport.querySelector('[data-node-role="current"]');
  const reportNodes = topRowNodes([...viewport.querySelectorAll('[data-node-role="report"]')]);

  if (managerNode && currentNode) {
    const from = cardAnchor(managerNode, "bottom", vRect);
    const to = cardAnchor(currentNode, "top", vRect);
    paths.push(curvePath(from.x, from.y, to.x, to.y));
  }

  if (currentNode && reportNodes.length) {
    const from = cardAnchor(currentNode, "bottom", vRect);
    reportNodes.forEach((node) => {
      const to = cardAnchor(node, "top", vRect);
      paths.push(curvePath(from.x, from.y, to.x, to.y));
    });
  }

  svg.innerHTML = paths.map((d) => `<path d="${d}"></path>`).join("");
  svg.classList.remove("is-hidden");
}

function hideBranchLines() {
  els.branchSvg.classList.add("is-hidden");
  els.branchSvg.innerHTML = "";
}

function redrawLinesOnScroll() {
  requestAnimationFrame(() => drawBranchLines());
}

function scheduleDrawLines() {
  const token = ++drawFrame;
  const run = () => {
    if (token !== drawFrame) return;
    drawBranchLines();
  };

  requestAnimationFrame(() => {
    requestAnimationFrame(run);
  });
  setTimeout(run, 50);
  setTimeout(run, 420);
}

function bindReportsScroll(scrollEl) {
  scrollEl.onscroll = () => redrawLinesOnScroll();
  scrollEl.addEventListener("wheel", (e) => {
    if (scrollEl.scrollWidth <= scrollEl.clientWidth) return;
    if (Math.abs(e.deltaY) > Math.abs(e.deltaX)) {
      e.preventDefault();
      scrollEl.scrollLeft += e.deltaY;
    }
  }, { passive: false });
}

function renderPathSpine() {
  if (branchStack.length <= 1) {
    els.pathSpine.classList.add("hidden");
    return;
  }
  els.pathSpine.classList.remove("hidden");
  els.pathSpine.innerHTML = branchStack.map((u, i) => {
    const isLast = i === branchStack.length - 1;
    if (isLast) return `<span class="current-crumb">${escapeHtml(u.displayName)}</span>`;
    return `<button type="button" data-spine-idx="${i}">${escapeHtml(u.displayName)}</button><span class="sep">›</span>`;
  }).join("");

  els.pathSpine.querySelectorAll("[data-spine-idx]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const idx = Number(btn.dataset.spineIdx);
      navigateToStackIndex(idx, idx < branchStack.length - 1 ? "up" : null);
    });
  });
}

async function renderTree() {
  const current = branchStack[branchStack.length - 1];
  if (!current) return;

  const manager = branchStack.length >= 2
    ? branchStack[branchStack.length - 2]
    : await fetchManager(current.id);

  let reports = await fetchDirectReports(current.id);
  reports = filterByOrg(reports);
  reports = sortPeople(reports);

  const parts = [];
  const showManager = !!(manager?.id && orgMatches(manager));
  const filterActive = companyFilter || departmentFilter || teamFilter;
  const atOrgTop = !manager;

  let topPeers = [];
  if (atOrgTop) {
    const roots = await fetchOrgRoots();
    topPeers = filterByOrg(roots).filter((u) => u.id !== current.id);
    // Ensure current root is known even if discovery missed them
    if (!roots.some((u) => u.id === current.id)) {
      cacheUser(current);
    }
  }

  if (showManager) {
    parts.push(`<div class="tree-row manager-row">${await nodeHtml(manager, "manager")}</div>`);
  } else if (manager?.id && filterActive) {
    parts.push('<p class="empty-layer">Manager is outside the selected filters</p>');
  } else if (atOrgTop) {
    // Keep vertical rhythm consistent with the 3-level layout
    parts.push(`<div class="tree-row manager-row" aria-hidden="true">
      <div class="tree-node manager" style="visibility:hidden;pointer-events:none">
        <div class="node-card"></div>
      </div>
    </div>`);
  }

  if (atOrgTop && topPeers.length > 0) {
    const rootPeople = sortPeople([current, ...topPeers]);
    const rootCards = await Promise.all(rootPeople.map(async (u) => {
      if (u.id === current.id) return nodeHtml(u, "current");
      const count = await countDirectReports(u.id);
      return nodeHtml(u, "peer", count);
    }));
    parts.push(`<div class="tree-row current-row">${rootCards.join("")}</div>`);
  } else {
    parts.push(`<div class="tree-row current-row">${await nodeHtml(current, "current")}</div>`);
  }

  if (reports.length > 0) {
    const counts = await Promise.all(reports.map((r) => countDirectReports(r.id)));
    parts.push(`<div class="reports-slot"><div class="reports-scroll"><div class="tree-row reports-row">${(await Promise.all(
      reports.map((r, i) => nodeHtml(r, "report", counts[i]))
    )).join("")}</div></div></div>`);
  } else {
    parts.push(`<div class="reports-slot"><p class="empty-layer">No direct reports${
      filterActive ? " matching filters" : ""
    }</p></div>`);
  }

  const animating = !!animDirection;
  if (animating) hideBranchLines();

  els.treeInner.innerHTML = parts.join("");

  if (animating) {
    els.treeInner.classList.remove("anim-up", "anim-down");
    void els.treeInner.offsetWidth;
    const animClass = animDirection === "up" ? "anim-up" : "anim-down";
    els.treeInner.classList.add(animClass);
    animDirection = null;

    const onAnimEnd = () => {
      els.treeInner.classList.remove("anim-up", "anim-down");
      els.treeInner.removeEventListener("animationend", onAnimEnd);
      scheduleDrawLines();
    };
    els.treeInner.addEventListener("animationend", onAnimEnd);
  }

  els.treeInner.querySelectorAll("[data-nav-id]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const id = btn.dataset.navId;
      const role = btn.dataset.navRole;
      if (role === "manager") navigateUp(id);
      else if (role === "report") navigateDown(id);
      else if (role === "peer") navigateToPeerRoot(id);
    });
  });

  els.treeInner.querySelectorAll("[data-edit-id]").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      openEditDrawer(btn.dataset.editId);
    });
  });

  renderPathSpine();
  refreshOrgFilters();
  if (!animating) scheduleDrawLines();

  const scrollEl = els.treeViewport.querySelector(".reports-scroll");
  if (scrollEl) bindReportsScroll(scrollEl);
}

async function navigateToPeerRoot(userId) {
  if (branchStack[branchStack.length - 1]?.id === userId) return;
  animDirection = "down";
  const user = await fetchUser(userId);
  branchStack = [user];
  await renderTree();
}

async function navigateUp(managerId) {
  const idx = branchStack.findIndex((u) => u.id === managerId);
  if (idx >= 0) {
    navigateToStackIndex(idx, "up");
    return;
  }
  animDirection = "up";
  const mgr = await fetchUser(managerId);
  branchStack = branchStack.slice(0, -1);
  if (branchStack[branchStack.length - 1]?.id !== mgr.id) {
    branchStack.push(mgr);
  }
  await renderTree();
}

async function navigateDown(reportId) {
  if (branchStack[branchStack.length - 1]?.id === reportId) return;
  animDirection = "down";
  const user = await fetchUser(reportId);
  branchStack.push(user);
  await renderTree();
}

function navigateToStackIndex(idx, direction) {
  if (idx < 0 || idx >= branchStack.length) return;
  animDirection = direction || (idx < branchStack.length - 1 ? "up" : null);
  branchStack = branchStack.slice(0, idx + 1);
  renderTree();
}

async function navigateToPerson(userId) {
  setStatus("Loading…");
  try {
    branchStack = await buildBranchTo(userId);
    animDirection = null;
    els.searchInput.value = "";
    els.searchResults.classList.add("hidden");
    setStatus("");
    els.mainUi.classList.remove("hidden");
    await renderTree();
  } catch (err) {
    setStatus(err.message, true);
  }
}

function showBranchView() {
  els.adminUi.classList.add("hidden");
  els.fullTreeUi.classList.add("hidden");
  els.branchView.classList.remove("hidden");
  showExplorerChrome();
}

async function goHome() {
  if (!myUserId) return;
  showBranchView();
  animDirection = null;
  try {
    const me = await fetchUser(demoMode ? myUserId : "me");
    branchStack = [me];
    setStatus("");
    await renderTree();
  } catch (err) {
    setStatus(err.message || "Could not open your profile.", true);
  }
}

// ── Search ──────────────────────────────────────────────────────────────────
function runSearch(query) {
  clearTimeout(searchTimer);
  if (!query.trim()) {
    els.searchResults.classList.add("hidden");
    return;
  }
  searchTimer = setTimeout(async () => {
    let results = await graphSearch(query.trim());
    results = filterByOrg(results);
    results = sortPeople(results);
    if (!results.length) {
      els.searchResults.innerHTML = `<div class="empty-layer" style="padding:0.75rem;max-width:none">No matches</div>`;
    } else {
      els.searchResults.innerHTML = results.map((u) => {
        const company = u.companyName?.trim();
        const dept = u.department?.trim();
        const team = userTeam(u);
        const meta = [u.jobTitle, company, dept, team].filter(Boolean).map((s) => escapeHtml(s)).join(" · ");
        return `
        <button type="button" data-search-id="${u.id}">
          ${escapeHtml(u.displayName)}
          <span class="sub">${meta}</span>
        </button>
      `;
      }).join("");
      els.searchResults.querySelectorAll("[data-search-id]").forEach((btn) => {
        btn.addEventListener("click", () => navigateToPerson(btn.dataset.searchId));
      });
    }
    els.searchResults.classList.remove("hidden");
  }, 280);
}

// ── Full tree view ──────────────────────────────────────────────────────────
function setFullTreeLoading(on, msg = "Loading org…") {
  els.ftLoading.classList.toggle("hidden", !on);
  if (on) els.ftLoading.textContent = msg;
}

function applyFtTransform() {
  els.ftStage.style.transform =
    `translate(${ftPan.x}px, ${ftPan.y}px) scale(${ftPan.scale})`;
}

function focusPersonId() {
  return branchStack[branchStack.length - 1]?.id || myUserId;
}

async function loadFullOrg() {
  if (fullOrgPeople) return fullOrgPeople;
  if (fullOrgLoadPromise) return fullOrgLoadPromise;

  fullOrgLoadPromise = (async () => {
    const managerMap = new Map();
    const byId = new Map();

    if (demoMode) {
      Object.values(DEMO_USERS).forEach((raw) => {
        const user = stripInternal(raw);
        byId.set(user.id, user);
        managerMap.set(user.id, raw.managerId || null);
        cacheUser(user);
      });
      fullOrgPeople = [...byId.values()];
      fullOrgManagerMap = managerMap;
      return fullOrgPeople;
    }

    const roots = await fetchOrgRoots();
    const queue = roots.map((r) => ({ user: r, depth: 0 }));
    roots.forEach((r) => {
      byId.set(r.id, r);
      managerMap.set(r.id, null);
    });

    let processed = 0;
    while (queue.length) {
      const batch = queue.splice(0, 12);
      await Promise.all(batch.map(async ({ user, depth }) => {
        if (depth >= 15) return;
        processed += 1;
        if (processed % 8 === 0) {
          setFullTreeLoading(true, `Loading org… (${byId.size} people)`);
        }
        try {
          const reports = await fetchDirectReports(user.id);
          reports.forEach((rep) => {
            if (byId.has(rep.id)) return;
            byId.set(rep.id, rep);
            managerMap.set(rep.id, user.id);
            queue.push({ user: rep, depth: depth + 1 });
          });
        } catch { /* skip branch on error */ }
      }));
    }

    fullOrgPeople = [...byId.values()];
    fullOrgManagerMap = managerMap;
    return fullOrgPeople;
  })().finally(() => {
    fullOrgLoadPromise = null;
  });

  return fullOrgLoadPromise;
}

function filteredFullOrgPeople() {
  const people = fullOrgPeople || [];
  if (!companyFilter && !departmentFilter && !teamFilter) return people;
  return people.filter(orgMatches);
}

function buildEmployeeForest(people, managerMap) {
  const idSet = new Set(people.map((p) => p.id));
  const children = new Map();
  people.forEach((p) => children.set(p.id, []));

  people.forEach((p) => {
    const mid = managerMap.get(p.id);
    if (mid && idSet.has(mid) && children.has(mid)) {
      children.get(mid).push(p);
    }
  });

  children.forEach((list, id) => {
    children.set(id, sortPeople(list));
  });

  const roots = sortPeople(people.filter((p) => {
    const mid = managerMap.get(p.id);
    return !mid || !idSet.has(mid);
  }));

  function layoutNode(user) {
    const kids = (children.get(user.id) || []).map(layoutNode);
    const node = {
      id: user.id,
      kind: "person",
      user,
      children: kids,
      width: FT.PERSON_W,
      height: FT.PERSON_H,
      x: 0,
      y: 0,
    };
    if (!kids.length) {
      node.subtreeWidth = FT.PERSON_W;
    } else {
      const kidsWidth = kids.reduce((s, k) => s + k.subtreeWidth, 0)
        + FT.H_GAP * (kids.length - 1);
      node.subtreeWidth = Math.max(FT.PERSON_W, kidsWidth);
    }
    return node;
  }

  function place(node, left, depth) {
    node.y = depth * (FT.PERSON_H + FT.V_GAP);
    if (!node.children.length) {
      node.x = left + (node.subtreeWidth - node.width) / 2;
      return;
    }
    let cursor = left + Math.max(0, (node.subtreeWidth - (
      node.children.reduce((s, k) => s + k.subtreeWidth, 0)
      + FT.H_GAP * (node.children.length - 1)
    )) / 2);
    node.children.forEach((child) => {
      place(child, cursor, depth + 1);
      cursor += child.subtreeWidth + FT.H_GAP;
    });
    const first = node.children[0];
    const last = node.children[node.children.length - 1];
    const mid = (first.x + first.width / 2 + last.x + last.width / 2) / 2;
    node.x = mid - node.width / 2;
  }

  const forest = roots.map(layoutNode);
  let left = 0;
  forest.forEach((root) => {
    place(root, left, 0);
    left += root.subtreeWidth + FT.H_GAP * 2;
  });
  return forest;
}

function buildStructureForest(people) {
  const companies = new Map();

  people.forEach((p) => {
    const company = (p.companyName || "").trim() || "Unassigned";
    const dept = (p.department || "").trim() || "Unassigned";
    const team = userTeam(p) || "Unassigned";
    const title = (p.jobTitle || "").trim() || "Unassigned";
    if (!companies.has(company)) companies.set(company, new Map());
    const depts = companies.get(company);
    if (!depts.has(dept)) depts.set(dept, new Map());
    const teams = depts.get(dept);
    if (!teams.has(team)) teams.set(team, new Map());
    const titles = teams.get(team);
    titles.set(title, (titles.get(title) || 0) + 1);
  });

  const companyNames = [...companies.keys()].sort((a, b) => a.localeCompare(b));

  function layoutJob(title, count, company, dept, team) {
    return {
      id: `job:${company}|${dept}|${team}|${title}`,
      kind: "job",
      label: title,
      count,
      children: [],
      width: FT.JOB_W,
      height: FT.JOB_H,
      subtreeWidth: FT.JOB_W,
      x: 0,
      y: 0,
    };
  }

  function layoutTeam(team, titlesMap, company, dept) {
    const titles = [...titlesMap.entries()]
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([title, count]) => layoutJob(title, count, company, dept, team));
    const count = titles.reduce((s, t) => s + t.count, 0);
    const kidsWidth = titles.reduce((s, t) => s + t.subtreeWidth, 0)
      + FT.H_GAP * Math.max(0, titles.length - 1);
    return {
      id: `team:${company}|${dept}|${team}`,
      kind: "team",
      label: team,
      count,
      children: titles,
      width: FT.TEAM_W,
      height: FT.TEAM_H,
      subtreeWidth: Math.max(FT.TEAM_W, kidsWidth),
      x: 0,
      y: 0,
    };
  }

  function layoutDept(dept, teamsMap, company) {
    const teamNames = [...teamsMap.keys()].sort((a, b) => a.localeCompare(b));
    const teams = teamNames.map((t) => layoutTeam(t, teamsMap.get(t), company, dept));
    const count = teams.reduce((s, d) => s + d.count, 0);
    const kidsWidth = teams.reduce((s, t) => s + t.subtreeWidth, 0)
      + FT.H_GAP * Math.max(0, teams.length - 1);
    return {
      id: `dept:${company}|${dept}`,
      kind: "dept",
      label: dept,
      count,
      children: teams,
      width: FT.DEPT_W,
      height: FT.DEPT_H,
      subtreeWidth: Math.max(FT.DEPT_W, kidsWidth),
      x: 0,
      y: 0,
    };
  }

  function layoutCompany(company, deptsMap) {
    const deptNames = [...deptsMap.keys()].sort((a, b) => a.localeCompare(b));
    const depts = deptNames.map((d) => layoutDept(d, deptsMap.get(d), company));
    const count = depts.reduce((s, d) => s + d.count, 0);
    const kidsWidth = depts.reduce((s, d) => s + d.subtreeWidth, 0)
      + FT.H_GAP * Math.max(0, depts.length - 1);
    return {
      id: `company:${company}`,
      kind: "company",
      label: company,
      count,
      company,
      children: depts,
      width: FT.COMPANY_W,
      height: FT.COMPANY_H,
      subtreeWidth: Math.max(FT.COMPANY_W, kidsWidth),
      x: 0,
      y: 0,
    };
  }

  function place(node, left, depth) {
    const rowH = Math.max(FT.COMPANY_H, FT.DEPT_H, FT.TEAM_H, FT.JOB_H);
    node.y = depth * (rowH + FT.V_GAP);
    if (!node.children.length) {
      node.x = left + (node.subtreeWidth - node.width) / 2;
      return;
    }
    const kidsSpan = node.children.reduce((s, k) => s + k.subtreeWidth, 0)
      + FT.H_GAP * (node.children.length - 1);
    let cursor = left + Math.max(0, (node.subtreeWidth - kidsSpan) / 2);
    node.children.forEach((child) => {
      place(child, cursor, depth + 1);
      cursor += child.subtreeWidth + FT.H_GAP;
    });
    const first = node.children[0];
    const last = node.children[node.children.length - 1];
    const mid = (first.x + first.width / 2 + last.x + last.width / 2) / 2;
    node.x = mid - node.width / 2;
  }

  const forest = companyNames.map((c) => layoutCompany(c, companies.get(c)));
  let left = 0;
  forest.forEach((root) => {
    place(root, left, 0);
    left += root.subtreeWidth + FT.H_GAP * 2;
  });
  return forest;
}

function flattenForest(forest) {
  const nodes = [];
  const edges = [];
  function walk(node) {
    nodes.push(node);
    node.children.forEach((child) => {
      edges.push({ from: node, to: child });
      walk(child);
    });
  }
  forest.forEach(walk);
  return { nodes, edges };
}

function measureForest(nodes) {
  let maxX = 0;
  let maxY = 0;
  nodes.forEach((n) => {
    maxX = Math.max(maxX, n.x + n.width);
    maxY = Math.max(maxY, n.y + n.height);
  });
  return { width: Math.max(maxX + 40, 400), height: Math.max(maxY + 40, 300) };
}

function ftPersonHtml(node, meId) {
  const u = node.user;
  const isMe = u.id === meId;
  const photo = photoCache.get(u.id);
  const avatar = photo
    ? `<img class="ft-avatar" src="${photo}" alt="" />`
    : `<div class="ft-avatar">${escapeHtml(initials(u))}</div>`;
  return `<div class="ft-node ft-person${isMe ? " is-me" : ""}" data-ft-id="${escapeHtml(u.id)}" style="left:${node.x}px;top:${node.y}px;width:${node.width}px;height:${node.height}px">
    ${avatar}
    <span class="ft-name">${escapeHtml(u.displayName)}</span>
    <span class="ft-title">${escapeHtml(u.jobTitle || "—")}</span>
  </div>`;
}

function ftStructureHtml(node) {
  const tone = node.kind === "company" ? companyBadgeClass(node.company || node.label) : "";
  const focus = node.id === ftFocusNodeId ? " ft-focus" : "";
  const countLabel = node.kind === "job"
    ? `${escapeHtml(node.label)} · ${node.count}`
    : escapeHtml(node.label);
  const sub = node.kind === "job"
    ? ""
    : `<span class="ft-count">${node.count} people</span>`;
  return `<div class="ft-node ft-${node.kind}${focus} ${tone}" data-ft-id="${escapeHtml(node.id)}" style="left:${node.x}px;top:${node.y}px;width:${node.width}px;min-height:${node.height}px">
    <span class="ft-label">${countLabel}</span>
    ${sub}
  </div>`;
}

function drawFtEdges(edges, bounds) {
  const svg = els.ftSvg;
  svg.setAttribute("width", String(bounds.width));
  svg.setAttribute("height", String(bounds.height));
  svg.style.width = `${bounds.width}px`;
  svg.style.height = `${bounds.height}px`;
  const paths = edges.map(({ from, to }) => {
    const x1 = from.x + from.width / 2;
    const y1 = from.y + from.height;
    const x2 = to.x + to.width / 2;
    const y2 = to.y;
    return `<path d="${curvePath(x1, y1, x2, y2)}"></path>`;
  });
  svg.innerHTML = paths.join("");
}

function centerOnNodeId(nodeId, nodes) {
  const node = nodes.find((n) => n.id === nodeId)
    || nodes.find((n) => n.user?.id === nodeId);
  const vp = els.ftViewport.getBoundingClientRect();
  if (!node) {
    ftPan.scale = Math.min(1, Math.min(vp.width / (ftLayoutBounds.width || 1), vp.height / (ftLayoutBounds.height || 1)) * 0.9);
    ftPan.x = (vp.width - ftLayoutBounds.width * ftPan.scale) / 2;
    ftPan.y = 40;
    applyFtTransform();
    return;
  }
  const cx = node.x + node.width / 2;
  const cy = node.y + node.height / 2;
  const scale = Math.min(1.15, Math.max(0.45, Math.min(vp.width / 700, vp.height / 500)));
  ftPan.scale = scale;
  ftPan.x = vp.width / 2 - cx * scale;
  ftPan.y = vp.height / 2 - cy * scale;
  applyFtTransform();
}

function structureFocusIdForUser(user) {
  if (!user) return null;
  const company = (user.companyName || "").trim() || "Unassigned";
  const dept = (user.department || "").trim() || "Unassigned";
  const team = userTeam(user) || "Unassigned";
  const title = (user.jobTitle || "").trim() || "Unassigned";
  return `job:${company}|${dept}|${team}|${title}`;
}

async function renderFullTree() {
  const people = filteredFullOrgPeople();
  const meId = focusPersonId();
  const me = people.find((p) => p.id === meId) || userCache.get(meId) || branchStack[branchStack.length - 1];

  if (!people.length) {
    els.ftSvg.innerHTML = "";
    els.ftNodes.innerHTML = `<div class="empty-layer" style="position:absolute;left:50%;top:40%;transform:translate(-50%,-50%);max-width:320px">No people match the current filters</div>`;
    ftLayoutBounds = { width: 400, height: 300 };
    ftPan = { x: 0, y: 0, scale: 1 };
    applyFtTransform();
    return;
  }

  let forest;
  if (ftMode === "employees") {
    forest = buildEmployeeForest(people, fullOrgManagerMap || new Map());
    ftFocusNodeId = meId;
  } else {
    forest = buildStructureForest(people);
    ftFocusNodeId = structureFocusIdForUser(me);
  }

  const { nodes, edges } = flattenForest(forest);
  ftLayoutBounds = measureForest(nodes);
  els.ftStage.style.width = `${ftLayoutBounds.width}px`;
  els.ftStage.style.height = `${ftLayoutBounds.height}px`;
  els.ftNodes.style.width = `${ftLayoutBounds.width}px`;
  els.ftNodes.style.height = `${ftLayoutBounds.height}px`;

  if (ftMode === "employees") {
    els.ftNodes.innerHTML = nodes.map((n) => ftPersonHtml(n, meId)).join("");
    hydrateFtPhotos(nodes);
  } else {
    els.ftNodes.innerHTML = nodes.map((n) => ftStructureHtml(n)).join("");
  }

  drawFtEdges(edges, ftLayoutBounds);
  centerOnNodeId(ftFocusNodeId, nodes);
}

async function hydrateFtPhotos(nodes) {
  const nearby = nodes.slice(0, 80);
  await Promise.all(nearby.map(async (n) => {
    try {
      const url = await fetchPhotoUrl(n.user.id);
      if (!url) return;
      const host = els.ftNodes.querySelector(`[data-ft-id="${CSS.escape(n.user.id)}"]`);
      if (!host) return;
      const existing = host.querySelector(".ft-avatar");
      if (!existing || existing.tagName === "IMG") return;
      const img = document.createElement("img");
      img.className = "ft-avatar";
      img.src = url;
      img.alt = "";
      existing.replaceWith(img);
    } catch { /* ignore */ }
  }));
}

async function openFullTree() {
  els.adminUi.classList.add("hidden");
  els.branchView.classList.add("hidden");
  els.fullTreeUi.classList.remove("hidden");
  els.btnFullTree.classList.add("hidden");
  setFullTreeLoading(true);
  try {
    await loadFullOrg();
    await renderFullTree();
  } catch (err) {
    setStatus(err.message || "Could not load full tree.", true);
  } finally {
    setFullTreeLoading(false);
  }
  bindFtPanZoom();
}

function closeFullTree() {
  els.fullTreeUi.classList.add("hidden");
  els.branchView.classList.remove("hidden");
  els.btnFullTree.classList.remove("hidden");
  scheduleDrawLines();
}

async function setFtMode(mode) {
  if (ftMode === mode) return;
  ftMode = mode;
  els.btnFtEmployees.classList.toggle("active", mode === "employees");
  els.btnFtStructure.classList.toggle("active", mode === "structure");
  setFullTreeLoading(true, "Updating view…");
  try {
    await renderFullTree();
  } finally {
    setFullTreeLoading(false);
  }
}

async function focusMeInFullTree() {
  if (!fullOrgPeople) return;
  await renderFullTree();
}

function bindFtPanZoom() {
  if (ftPanZoomBound) return;
  ftPanZoomBound = true;
  const vp = els.ftViewport;

  vp.addEventListener("pointerdown", (e) => {
    if (e.button !== 0) return;
    ftDragging = true;
    ftDragStart = { x: e.clientX, y: e.clientY, panX: ftPan.x, panY: ftPan.y };
    vp.classList.add("is-dragging");
    vp.setPointerCapture(e.pointerId);
  });

  vp.addEventListener("pointermove", (e) => {
    if (!ftDragging || !ftDragStart) return;
    ftPan.x = ftDragStart.panX + (e.clientX - ftDragStart.x);
    ftPan.y = ftDragStart.panY + (e.clientY - ftDragStart.y);
    applyFtTransform();
  });

  const endDrag = (e) => {
    if (!ftDragging) return;
    ftDragging = false;
    ftDragStart = null;
    vp.classList.remove("is-dragging");
    try { vp.releasePointerCapture(e.pointerId); } catch { /* ignore */ }
  };
  vp.addEventListener("pointerup", endDrag);
  vp.addEventListener("pointercancel", endDrag);

  vp.addEventListener("wheel", (e) => {
    e.preventDefault();
    const rect = vp.getBoundingClientRect();
    const mx = e.clientX - rect.left;
    const my = e.clientY - rect.top;
    const prev = ftPan.scale;
    const delta = e.deltaY > 0 ? 0.9 : 1.1;
    const next = Math.min(FT.MAX_SCALE, Math.max(FT.MIN_SCALE, prev * delta));
    if (next === prev) return;
    // Zoom toward cursor
    const worldX = (mx - ftPan.x) / prev;
    const worldY = (my - ftPan.y) / prev;
    ftPan.scale = next;
    ftPan.x = mx - worldX * next;
    ftPan.y = my - worldY * next;
    applyFtTransform();
  }, { passive: false });
}

function showExplorerChrome() {
  els.btnFullTree.classList.remove("hidden");
  els.btnAdmin.classList.toggle("hidden", !isAdmin);
}

// ── Admin ───────────────────────────────────────────────────────────────────
function personMeta(user) {
  return [user.jobTitle, user.companyName, user.department, userTeam(user)]
    .filter(Boolean)
    .map((s) => escapeHtml(s))
    .join(" · ");
}

const extraCatalog = {
  company: new Set(),
  department: new Set(),
  team: new Set(),
  role: new Set(),
};

function clearExtraCatalog() {
  Object.keys(extraCatalog).forEach((kind) => extraCatalog[kind].clear());
}

const ADD_NEW = "__new__";

const EDIT_FIELDS = {
  company: {
    selectId: "edit-company",
    wrapId: "new-company-wrap",
    inputId: "new-company",
    addId: "btn-add-company",
    catalog: "company",
    payload: "companyName",
    fromUser: (u) => (u.companyName || "").trim(),
  },
  department: {
    selectId: "edit-department",
    wrapId: "new-department-wrap",
    inputId: "new-department",
    addId: "btn-add-department",
    catalog: "department",
    payload: "department",
    fromUser: (u) => (u.department || "").trim(),
  },
  team: {
    selectId: "edit-team",
    wrapId: "new-team-wrap",
    inputId: "new-team",
    addId: "btn-add-team",
    catalog: "team",
    payload: "team",
    fromUser: (u) => userTeam(u),
  },
  role: {
    selectId: "edit-role",
    wrapId: "new-role-wrap",
    inputId: "new-role",
    addId: "btn-add-role",
    catalog: "role",
    payload: "jobTitle",
    fromUser: (u) => (u.jobTitle || "").trim(),
  },
};

function catalogValues(kind) {
  rebuildOrgCatalogs();
  const values = new Set(extraCatalog[kind]);
  if (kind === "company") knownCompanies.forEach((v) => values.add(v));
  if (kind === "role") knownRoles.forEach((v) => values.add(v));
  if (kind === "department") {
    departmentsByCompany.forEach((set) => set.forEach((v) => values.add(v)));
  }
  if (kind === "team") {
    teamsByCompanyDept.forEach((set) => set.forEach((v) => values.add(v)));
  }
  userCache.forEach((u) => {
    const value = EDIT_FIELDS[kind].fromUser(u);
    if (value) values.add(value);
  });
  return [...values].sort((a, b) => a.localeCompare(b));
}

function fillEditSelect(kind, selected) {
  const field = EDIT_FIELDS[kind];
  const select = document.getElementById(field.selectId);
  const values = catalogValues(kind);
  if (selected && !values.includes(selected)) {
    values.push(selected);
    values.sort((a, b) => a.localeCompare(b));
  }
  select.innerHTML =
    `<option value="">— None —</option>` +
    values.map((v) => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join("") +
    `<option value="${ADD_NEW}">Add new value…</option>`;
  select.value = selected || "";
  toggleNewValueRow(kind, false);
}

function toggleNewValueRow(kind, on) {
  const wrap = document.getElementById(EDIT_FIELDS[kind].wrapId);
  const input = document.getElementById(EDIT_FIELDS[kind].inputId);
  wrap.classList.toggle("hidden", !on);
  if (on) {
    input.value = "";
    input.focus();
  }
}

function bindEditField(kind) {
  const field = EDIT_FIELDS[kind];
  const select = document.getElementById(field.selectId);
  select.addEventListener("change", () => {
    toggleNewValueRow(kind, select.value === ADD_NEW);
  });
  document.getElementById(field.addId).addEventListener("click", () => addCatalogValue(kind));
  document.getElementById(field.inputId).addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      addCatalogValue(kind);
    }
  });
}

function addCatalogValue(kind) {
  const field = EDIT_FIELDS[kind];
  const input = document.getElementById(field.inputId);
  const value = input.value.trim();
  if (!value) {
    input.focus();
    return;
  }
  const existing = catalogValues(kind).find((v) => v.toLowerCase() === value.toLowerCase());
  const chosen = existing || value;
  extraCatalog[kind].add(chosen);
  fillEditSelect(kind, chosen);
}

function selectedEditValue(kind) {
  const select = document.getElementById(EDIT_FIELDS[kind].selectId);
  if (select.value === ADD_NEW) return null;
  return select.value.trim();
}

function fillEditSelects() {
  const user = userCache.get(editingUserId);
  fillEditSelect("company", user ? EDIT_FIELDS.company.fromUser(user) : "");
  fillEditSelect("department", user ? EDIT_FIELDS.department.fromUser(user) : "");
  fillEditSelect("team", user ? EDIT_FIELDS.team.fromUser(user) : "");
  fillEditSelect("role", user ? EDIT_FIELDS.role.fromUser(user) : "");
}

function adminPeople() {
  if (directoryPeople?.length) return sortPeople(directoryPeople);
  if (fullOrgPeople?.length) return sortPeople(fullOrgPeople);
  return sortPeople([...userCache.values()]);
}

function renderAdminList(query) {
  const q = (query || "").trim().toLowerCase();
  let people = adminPeople();
  if (q) {
    people = people.filter((u) =>
      (u.displayName || "").toLowerCase().includes(q)
      || (u.jobTitle || "").toLowerCase().includes(q)
      || (u.companyName || "").toLowerCase().includes(q)
      || (u.department || "").toLowerCase().includes(q)
      || userTeam(u).toLowerCase().includes(q)
    );
  }
  if (!people.length) {
    els.adminResults.innerHTML = `<div class="empty-layer" style="max-width:none">No people found</div>`;
    return;
  }
  els.adminResults.innerHTML = people.map((u) => `
    <div class="admin-row">
      <div class="admin-row-main">
        <div class="admin-row-name">${escapeHtml(u.displayName)}</div>
        <span class="admin-row-meta">${personMeta(u) || "No org details yet"}</span>
      </div>
      <button type="button" data-edit-id="${u.id}">Edit</button>
    </div>
  `).join("");
  els.adminResults.querySelectorAll("[data-edit-id]").forEach((btn) => {
    btn.addEventListener("click", () => openEditDrawer(btn.dataset.editId));
  });
}

async function openAdmin() {
  if (!isAdmin) return;
  els.fullTreeUi.classList.add("hidden");
  els.branchView.classList.add("hidden");
  els.adminUi.classList.remove("hidden");
  els.btnFullTree.classList.add("hidden");
  els.btnAdmin.classList.add("hidden");
  const hint = els.adminUi.querySelector(".admin-hint");
  if (hint) {
    hint.textContent = demoMode
      ? "Preview mode — edits stay in this browser session."
      : "Pick company, department, team, and role from the lists. Use Add new value to create one.";
  }
  try {
    await fetchDirectoryPeople();
  } catch { /* list from cache */ }
  renderAdminList(els.adminSearchInput.value);
}

function csvEscape(value) {
  const text = String(value ?? "");
  if (/[",\n\r]/.test(text)) return `"${text.replace(/"/g, '""')}"`;
  return text;
}

function orgUsersCsv(people) {
  const header = ["id", "displayName", "userPrincipalName", "mail", "companyName", "department", "team", "jobTitle"];
  const rows = people.map((u) => [
    u.id,
    u.displayName,
    u.userPrincipalName,
    u.mail,
    u.companyName,
    u.department,
    userTeam(u),
    u.jobTitle,
  ].map(csvEscape).join(","));
  return [header.join(","), ...rows].join("\r\n");
}

async function downloadOrgCsv() {
  els.btnAdminCsv.disabled = true;
  const previous = els.btnAdminCsv.textContent;
  els.btnAdminCsv.textContent = "Preparing CSV…";
  try {
    await fetchDirectoryPeople();
    const people = adminPeople();
    const csv = orgUsersCsv(people);
    const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "org-users.csv";
    a.click();
    URL.revokeObjectURL(url);
  } catch (err) {
    setStatus(err.message || "Could not export CSV.", true);
  } finally {
    els.btnAdminCsv.disabled = false;
    els.btnAdminCsv.textContent = previous;
  }
}

function parseCsvText(text) {
  const rows = [];
  let row = [];
  let field = "";
  let quoted = false;
  const input = String(text || "").replace(/^\uFEFF/, "");
  for (let i = 0; i < input.length; i++) {
    const char = input[i];
    if (quoted) {
      if (char === '"') {
        if (input[i + 1] === '"') {
          field += '"';
          i += 1;
        } else {
          quoted = false;
        }
      } else {
        field += char;
      }
    } else if (char === '"') {
      quoted = true;
    } else if (char === ",") {
      row.push(field);
      field = "";
    } else if (char === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else if (char !== "\r") {
      field += char;
    }
  }
  if (field.length || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows.filter((item) => item.some((value) => String(value).trim()));
}

function updatesFromCsv(text) {
  const rows = parseCsvText(text);
  if (!rows.length) throw new Error("CSV is empty.");
  const header = rows[0].map((value) => value.trim().toLowerCase());
  const col = (name) => header.indexOf(name.toLowerCase());
  const idCol = col("id");
  if (idCol < 0) throw new Error("CSV must include an id column.");
  const companyCol = col("companyName");
  const deptCol = col("department");
  const teamCol = col("team");
  const titleCol = col("jobTitle");
  if (companyCol < 0 || deptCol < 0 || teamCol < 0 || titleCol < 0) {
    throw new Error("CSV must include companyName, department, team, and jobTitle columns.");
  }
  return rows.slice(1).map((row) => ({
    userId: (row[idCol] || "").trim(),
    companyName: companyCol >= 0 ? (row[companyCol] || "").trim() : "",
    department: deptCol >= 0 ? (row[deptCol] || "").trim() : "",
    team: teamCol >= 0 ? (row[teamCol] || "").trim() : "",
    jobTitle: titleCol >= 0 ? (row[titleCol] || "").trim() : "",
  })).filter((item) => item.userId);
}

async function postAdminUpdates(updates) {
  if (demoMode) {
    return updates.map((item) => {
      const raw = DEMO_USERS[item.userId];
      if (raw) {
        raw.companyName = item.companyName;
        raw.department = item.department;
        raw.team = item.team;
        raw.jobTitle = item.jobTitle;
      }
      applyUserPatch(item.userId, item);
      return { userId: item.userId, ok: true, ...item };
    });
  }

  const token = await getToken();
  if (!token) throw new Error("Sign in required.");
  const res = await fetch("/api/admin", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ updates }),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || "Bulk update failed.");
  return data.results || [];
}

async function applyCsvUpdates(updates) {
  const results = [];
  for (let i = 0; i < updates.length; i += 10) {
    const chunk = updates.slice(i, i + 10);
    els.btnAdminCsvApply.textContent = `Applying ${Math.min(i + chunk.length, updates.length)}/${updates.length}…`;
    const chunkResults = await postAdminUpdates(chunk);
    chunkResults.forEach((item) => {
      if (item.ok) applyUserPatch(item.userId, item);
    });
    results.push(...chunkResults);
  }
  directoryPeople = null;
  try {
    await fetchDirectoryPeople();
  } catch { /* keep patched cache */ }
  await refreshAfterEdit();
  return results;
}

function chooseAdminCsv() {
  els.adminCsvFile.value = "";
  els.adminCsvFile.click();
}

async function onAdminCsvChosen(e) {
  const file = e.target.files?.[0];
  if (!file) return;
  const previous = els.btnAdminCsvApply.textContent;
  els.btnAdminCsvApply.disabled = true;
  try {
    const updates = updatesFromCsv(await file.text());
    if (!updates.length) throw new Error("No people found in that CSV.");
    const ok = window.confirm(demoMode
      ? `Apply company, department, team, and role for ${updates.length} people in this preview session?`
      : `Write company, department, team, and role for ${updates.length} people to Microsoft 365?`);
    if (!ok) return;
    const results = await applyCsvUpdates(updates);
    const failed = results.filter((item) => !item.ok);
    const teamIssues = results.filter((item) => item.ok && item.teamError);
    if (failed.length) {
      const reasons = [...new Set(failed.map((item) => item.error).filter(Boolean))].slice(0, 2);
      setStatus(
        `Updated ${results.length - failed.length} people. ${failed.length} failed.${reasons.length ? ` ${reasons.join(" | ")}` : ""}`,
        true
      );
      console.warn("CSV apply failures", failed.slice(0, 10));
    } else if (teamIssues.length) {
      setStatus(`Updated ${results.length} people. Team (CustomAttribute1) failed for ${teamIssues.length} hybrid/Exchange-mastered mailboxes.`);
    } else {
      setStatus(`Updated ${results.length} people in Microsoft 365.`);
    }
  } catch (err) {
    setStatus(err.message || "Could not apply CSV.", true);
  } finally {
    els.btnAdminCsvApply.disabled = false;
    els.btnAdminCsvApply.textContent = previous;
    els.adminCsvFile.value = "";
  }
}

function closeAdmin() {
  els.adminUi.classList.add("hidden");
  els.branchView.classList.remove("hidden");
  showExplorerChrome();
  scheduleDrawLines();
}

function openEditDrawer(userId) {
  const user = userCache.get(userId);
  if (!user) return;
  editingUserId = userId;
  fillEditSelects();
  els.editDrawerPerson.textContent = [user.displayName, user.userPrincipalName || user.mail]
    .filter(Boolean)
    .join(" · ");
  els.editMessage.classList.add("hidden");
  els.editMessage.classList.remove("is-error");
  els.editMessage.textContent = "";
  els.editDrawer.classList.remove("hidden");
}

function closeEditDrawer() {
  els.editDrawer.classList.add("hidden");
  editingUserId = null;
}

function applyUserPatch(userId, payload) {
  const apply = (user) => {
    if (!user) return;
    user.companyName = payload.companyName;
    user.department = payload.department;
    user.team = payload.team;
    user.jobTitle = payload.jobTitle;
    if (!user.onPremisesExtensionAttributes) user.onPremisesExtensionAttributes = {};
    user.onPremisesExtensionAttributes.extensionAttribute1 = payload.team || null;
    cacheUser(user);
  };

  apply(userCache.get(userId));
  (directoryPeople || []).forEach((u) => { if (u.id === userId) apply(u); });
  branchStack.forEach((u) => { if (u.id === userId) apply(u); });
  if (fullOrgPeople) {
    const person = fullOrgPeople.find((u) => u.id === userId);
    if (person) apply(person);
  }
}

async function refreshAfterEdit() {
  clearExtraCatalog();
  refreshOrgFilters();
  if (!els.adminUi.classList.contains("hidden")) renderAdminList(els.adminSearchInput.value);
  if (!els.branchView.classList.contains("hidden")) await renderTree();
  if (!els.fullTreeUi.classList.contains("hidden") && fullOrgPeople) await renderFullTree();
}

async function saveEdit(e) {
  e.preventDefault();
  if (!editingUserId) return;

  const companyName = selectedEditValue("company");
  const department = selectedEditValue("department");
  const team = selectedEditValue("team");
  const jobTitle = selectedEditValue("role");
  if (companyName == null || department == null || team == null || jobTitle == null) {
    els.editMessage.classList.remove("hidden");
    els.editMessage.classList.add("is-error");
    els.editMessage.textContent = "Finish adding the new value, or pick an existing one.";
    return;
  }

  const payload = {
    userId: editingUserId,
    companyName,
    department,
    team,
    jobTitle,
  };

  els.btnEditSave.disabled = true;
  els.editMessage.classList.add("hidden");
  try {
    if (demoMode) {
      const raw = DEMO_USERS[editingUserId];
      if (raw) {
        raw.companyName = payload.companyName;
        raw.department = payload.department;
        raw.team = payload.team;
        raw.jobTitle = payload.jobTitle;
      }
      applyUserPatch(editingUserId, payload);
      els.editMessage.classList.remove("hidden", "is-error");
      els.editMessage.textContent = "Saved in preview (not written to Microsoft 365).";
    } else {
      const token = await getToken();
      if (!token) throw new Error("Sign in required.");
      const res = await fetch("/api/admin", {
        method: "PATCH",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(data.error || "Save failed.");
      applyUserPatch(editingUserId, {
        companyName: data.companyName || "",
        department: data.department || "",
        team: data.team || "",
        jobTitle: data.jobTitle || "",
      });
      els.editMessage.classList.remove("hidden", "is-error");
      els.editMessage.textContent = "Saved to Microsoft 365.";
    }
    await refreshAfterEdit();
  } catch (err) {
    els.editMessage.classList.remove("hidden");
    els.editMessage.classList.add("is-error");
    els.editMessage.textContent = err.message;
  } finally {
    els.btnEditSave.disabled = false;
  }
}

async function resolveAdminAccess() {
  if (demoMode) {
    isAdmin = true;
    els.btnAdmin.classList.remove("hidden");
    return;
  }
  try {
    const token = await getToken();
    if (!token) return;
    const res = await fetch("/api/admin", {
      headers: { Authorization: `Bearer ${token}` },
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) {
      isAdmin = false;
      els.btnAdmin.classList.add("hidden");
      const detail = data.error || res.statusText;
      console.warn("Admin check failed:", data);
      if (res.status >= 500) {
        const seen = data.envPresent
          ? Object.entries(data.envPresent)
            .filter(([, on]) => on)
            .map(([key]) => key)
            .join(", ")
          : "";
        setStatus(
          `Admin is unavailable: ${detail}${seen ? ` Visible env: ${seen}.` : " No matching env vars on this deploy."}`,
          true
        );
      }
      return;
    }
    isAdmin = !!data.isAdmin;
    els.btnAdmin.classList.toggle("hidden", !isAdmin);
    if (!isAdmin) {
      console.warn("Admin check: not a member of ADMIN_GROUP_ID", data.group, data.checks);
    }
  } catch (err) {
    isAdmin = false;
    els.btnAdmin.classList.add("hidden");
    console.warn("Admin check failed:", err);
  }
}

// ── Boot ────────────────────────────────────────────────────────────────────
async function startDemo() {
  demoMode = true;
  myUserId = "05295543-6a39-479c-a5d9-6fc9db95e1ed";
  branchStack = [stripInternal(DEMO_USERS[myUserId])];
  seedDemoOrgFilters();
  els.demoBanner.classList.remove("hidden");
  els.btnSignIn.classList.add("hidden");
  els.btnHome.classList.remove("hidden");
  showExplorerChrome();
  await resolveAdminAccess();
  setStatus("");
  els.mainUi.classList.remove("hidden");
  await renderTree();
}

els.btnSignIn.addEventListener("click", signIn);
els.btnSignOut.addEventListener("click", signOut);
els.btnHome.addEventListener("click", goHome);
els.btnFullTree.addEventListener("click", () => openFullTree());
els.btnAdmin.addEventListener("click", () => openAdmin());
els.btnAdminBack.addEventListener("click", () => closeAdmin());
els.btnAdminCsv.addEventListener("click", () => downloadOrgCsv());
els.btnAdminCsvApply.addEventListener("click", () => chooseAdminCsv());
els.adminCsvFile.addEventListener("change", (e) => onAdminCsvChosen(e));
Object.keys(EDIT_FIELDS).forEach((kind) => bindEditField(kind));
els.btnFtBack.addEventListener("click", () => closeFullTree());
els.btnFtEmployees.addEventListener("click", () => setFtMode("employees"));
els.btnFtStructure.addEventListener("click", () => setFtMode("structure"));
els.btnFtFocus.addEventListener("click", () => focusMeInFullTree());
els.btnLive.addEventListener("click", async () => {
  try {
    await initAuth();
    await signIn();
  } catch (err) {
    setStatus(err.message, true);
  }
});

els.searchInput.addEventListener("input", (e) => runSearch(e.target.value));
els.searchInput.addEventListener("focus", (e) => { if (e.target.value.trim()) runSearch(e.target.value); });
els.adminSearchInput.addEventListener("input", (e) => renderAdminList(e.target.value));
document.addEventListener("click", (e) => {
  if (!els.searchInput.contains(e.target) && !els.searchResults.contains(e.target)) {
    els.searchResults.classList.add("hidden");
  }
});

els.companyFilter.addEventListener("change", (e) => {
  companyFilter = e.target.value;
  refreshDeptDropdown();
  refreshTeamDropdown();
  renderTree();
  if (!els.fullTreeUi.classList.contains("hidden") && fullOrgPeople) renderFullTree();
  if (els.searchInput.value.trim()) runSearch(els.searchInput.value);
});

els.deptFilter.addEventListener("change", (e) => {
  departmentFilter = e.target.value;
  refreshTeamDropdown();
  renderTree();
  if (!els.fullTreeUi.classList.contains("hidden") && fullOrgPeople) renderFullTree();
  if (els.searchInput.value.trim()) runSearch(els.searchInput.value);
});

els.teamFilter.addEventListener("change", (e) => {
  teamFilter = e.target.value;
  renderTree();
  if (!els.fullTreeUi.classList.contains("hidden") && fullOrgPeople) renderFullTree();
  if (els.searchInput.value.trim()) runSearch(els.searchInput.value);
});

els.editForm.addEventListener("submit", saveEdit);
els.editDrawer.querySelectorAll("[data-drawer-close]").forEach((el) => {
  el.addEventListener("click", () => closeEditDrawer());
});
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && !els.editDrawer.classList.contains("hidden")) closeEditDrawer();
});

window.addEventListener("resize", () => {
  scheduleDrawLines();
  if (!els.fullTreeUi.classList.contains("hidden")) applyFtTransform();
});

(async function boot() {
  try {
    if (isDemoMode()) {
      setStatus("Loading preview…");
      showSetupNote();
      await startDemo();
      return;
    }

    await ensureMsal();
    const authed = await initAuth();
    if (!authed) {
      setStatus("Sign in with your Microsoft 365 account.");
      return;
    }

    demoMode = false;
    els.demoBanner.classList.add("hidden");
    setStatus("Loading…");
    const me = await fetchUser("me");
    myUserId = me.id;
    branchStack = [me];
    setStatus("");
    els.mainUi.classList.remove("hidden");
    showExplorerChrome();
    await resolveAdminAccess();
    await renderTree();
  } catch (err) {
    setStatus(err.message, true);
  }
})();
