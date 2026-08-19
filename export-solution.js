/**
 * export-solution.js — Export the "MakeDatesPerfect" solution from a Dataverse
 * environment to managed + unmanaged .zip files (into ./solution).
 *
 * Read-only against the environment (ExportSolution generates a package; it does
 * not modify anything). Auth is MSAL device-code (same public client as deploy.js).
 *
 * Usage:  node export-solution.js https://yourorg.crm.dynamics.com
 */

const msal = require("@azure/msal-node");
const fetch = require("node-fetch");
const fs = require("fs");
const path = require("path");

const CRM_BASE_URL = (process.argv[2] || process.env.CRM_URL || "").replace(/\/+$/, "");
if (!CRM_BASE_URL) {
    console.error("\n\u2716 Missing environment URL.");
    console.error("  Usage: node export-solution.js https://yourorg.crm.dynamics.com\n");
    process.exit(1);
}
const API_URL = `${CRM_BASE_URL}/api/data/v9.2`;
const PUBLIC_CLIENT_ID = "51f81489-12ee-4a9e-aaae-a2591f45987d";
const SOLUTION_UNIQUE_NAME = "MakeDatesPerfect";
const OUT_DIR = path.join(__dirname, "solution");
const CACHE_FILE = path.join(__dirname, ".msal-cache.json");

let accessToken = null;

async function authenticate() {
    const pca = new msal.PublicClientApplication({
        auth: { clientId: PUBLIC_CLIENT_ID, authority: "https://login.microsoftonline.com/organizations" },
    });
    if (fs.existsSync(CACHE_FILE)) pca.getTokenCache().deserialize(fs.readFileSync(CACHE_FILE, "utf-8"));

    const accounts = await pca.getTokenCache().getAllAccounts();
    if (accounts.length > 0) {
        try {
            const r = await pca.acquireTokenSilent({ scopes: [`${CRM_BASE_URL}/.default`], account: accounts[0] });
            accessToken = r.accessToken;
            fs.writeFileSync(CACHE_FILE, pca.getTokenCache().serialize(), "utf-8");
            console.log(`\u2714 Authenticated silently as: ${accounts[0].username}\n`);
            return;
        } catch (e) { /* fall through to device code */ }
    }
    const r = await pca.acquireTokenByDeviceCode({
        scopes: [`${CRM_BASE_URL}/.default`],
        deviceCodeCallback: (resp) => console.log(`\n${resp.message}\n`),
    });
    accessToken = r.accessToken;
    fs.writeFileSync(CACHE_FILE, pca.getTokenCache().serialize(), "utf-8");
    console.log(`\u2714 Authenticated as: ${r.account.username}\n`);
}

async function exportSolution(managed) {
    const res = await fetch(`${API_URL}/ExportSolution`, {
        method: "POST",
        headers: {
            Authorization: `Bearer ${accessToken}`,
            Accept: "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            "Content-Type": "application/json",
        },
        body: JSON.stringify({ SolutionName: SOLUTION_UNIQUE_NAME, Managed: managed }),
    });
    if (!res.ok) throw new Error(`ExportSolution (managed=${managed}) \u2192 ${res.status}: ${await res.text()}`);
    const data = await res.json();
    const buffer = Buffer.from(data.ExportSolutionFile, "base64");
    const file = path.join(OUT_DIR, `MakeDatesPerfect_${managed ? "managed" : "unmanaged"}.zip`);
    fs.writeFileSync(file, buffer);
    console.log(`  \u2714 Wrote ${path.relative(__dirname, file)} (${(buffer.length / 1024).toFixed(1)} KB)`);
}

async function main() {
    console.log("Exporting solution \"MakeDatesPerfect\" from " + CRM_BASE_URL + "\n");
    await authenticate();
    fs.mkdirSync(OUT_DIR, { recursive: true });
    await exportSolution(false); // unmanaged
    await exportSolution(true);  // managed
    console.log("\nDone.\n");
}

main().catch((err) => { console.error("\n\u2716 Export failed:", err.message); process.exit(1); });
