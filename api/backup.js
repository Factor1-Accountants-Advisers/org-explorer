import { createHash, timingSafeEqual } from "node:crypto";
import { env, getAppToken } from "./_lib/graph.js";
import { buildTree, loadPeople, snapshot } from "./_lib/org.js";

// Nightly Vercel cron (see vercel.json). Saves the org chart to Azure Blob Storage as
// org-chart/YYYY/MM/DD/org-chart.json and .csv, dated in Melbourne time. A rerun on the
// same day overwrites that day's files. Writes as the Org Explorer app, which needs
// Storage Blob Data Contributor on the container.

const STORAGE_API_VERSION = "2023-11-03";

const CSV_COLUMNS = [
  "snapshotDate", "id", "name", "givenName", "surname", "email", "userPrincipalName",
  "company", "department", "team", "role", "location", "officeLocation", "employeeType",
  "managerId", "managerName", "directReportCount",
];

function sameSecret(a, b) {
  const ha = createHash("sha256").update(String(a)).digest();
  const hb = createHash("sha256").update(String(b)).digest();
  return timingSafeEqual(ha, hb);
}

function melbourneDate(date = new Date()) {
  // en-CA formats as YYYY-MM-DD.
  return new Intl.DateTimeFormat("en-CA", {
    timeZone: "Australia/Melbourne",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(date);
}

function csvCell(value) {
  const text = value == null ? "" : String(value);
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

function toCsv(snap, snapshotDate) {
  const rows = snap.people.map((p) => CSV_COLUMNS.map((col) => {
    if (col === "snapshotDate") return snapshotDate;
    if (col === "directReportCount") return p.directReportIds.length;
    return p[col];
  }).map(csvCell).join(","));
  // BOM so Excel opens names with accents correctly.
  return `﻿${CSV_COLUMNS.join(",")}\r\n${rows.join("\r\n")}\r\n`;
}

async function putBlob(token, account, container, name, body, contentType) {
  const path = name.split("/").map(encodeURIComponent).join("/");
  const res = await fetch(`https://${account}.blob.core.windows.net/${container}/${path}`, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`,
      "x-ms-version": STORAGE_API_VERSION,
      "x-ms-date": new Date().toUTCString(),
      "x-ms-blob-type": "BlockBlob",
      "x-ms-blob-content-type": contentType,
      "Content-Type": contentType,
    },
    body,
  });
  if (!res.ok) {
    const detail = (await res.text().catch(() => "")).match(/<Message>([^<]*)/)?.[1] || res.statusText;
    throw new Error(`Upload of ${name} failed (${res.status}): ${detail.trim()}`);
  }
}

export default async function handler(req, res) {
  res.setHeader("Cache-Control", "no-store");

  // Vercel cron sends Authorization: Bearer <CRON_SECRET>.
  const cronSecret = env("CRON_SECRET");
  if (!cronSecret) {
    res.status(503).json({ error: "CRON_SECRET is not set on this deployment." });
    return;
  }
  const auth = req.headers.authorization || "";
  if (!auth.startsWith("Bearer ") || !sameSecret(auth.slice(7).trim(), cronSecret)) {
    res.status(401).json({ error: "Unauthorized." });
    return;
  }

  const account = env("BACKUP_STORAGE_ACCOUNT") || "stf1cwm";
  const container = env("BACKUP_CONTAINER") || "hr-user-data";

  try {
    const [graphToken, storageToken] = await Promise.all([
      getAppToken(),
      getAppToken("https://storage.azure.com/.default"),
    ]);
    const { people, byId } = await loadPeople(graphToken);
    const snap = snapshot(people, buildTree(people, byId));
    const snapshotDate = melbourneDate();
    const prefix = `org-chart/${snapshotDate.replace(/-/g, "/")}/org-chart`;

    await putBlob(storageToken, account, container, `${prefix}.json`,
      JSON.stringify({ snapshotDate, ...snap }, null, 2), "application/json; charset=utf-8");
    await putBlob(storageToken, account, container, `${prefix}.csv`,
      toCsv(snap, snapshotDate), "text/csv; charset=utf-8");

    res.status(200).json({
      ok: true,
      snapshotDate,
      count: snap.count,
      blobs: [`${container}/${prefix}.json`, `${container}/${prefix}.csv`],
    });
  } catch (err) {
    console.error("Org chart backup failed:", err);
    res.status(500).json({ error: err.message || "Backup failed." });
  }
}
