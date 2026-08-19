/**
 * deploy.js — Authenticate via MSAL Device Code Flow, then deploy
 * the "Make Dates Perfect" wizard as a Web Resource into Dynamics 365.
 *
 * Usage:  node deploy.js
 */

const msal = require("@azure/msal-node");
const fetch = require("node-fetch");
const fs = require("fs");
const path = require("path");

// ─── Configuration ──────────────────────────────────────────────
// Target environment URL — pass it as an argument or via the CRM_URL env var:
//   node deploy.js https://yourorg.crm.dynamics.com
const CRM_BASE_URL = (process.argv[2] || process.env.CRM_URL || "").replace(/\/+$/, "");
if (!CRM_BASE_URL) {
    console.error("\n\u2716 Missing environment URL.");
    console.error("  Usage: node deploy.js https://yourorg.crm.dynamics.com");
    console.error("         (or set the CRM_URL environment variable)\n");
    process.exit(1);
}
const API_URL = `${CRM_BASE_URL}/api/data/v9.2`;
const PUBLIC_CLIENT_ID = "51f81489-12ee-4a9e-aaae-a2591f45987d";

const SOLUTION_UNIQUE_NAME = "MakeDatesPerfect";
const SOLUTION_DISPLAY_NAME = "Make Dates Perfect";
const SOLUTION_DESCRIPTION = "Demo-prep wizard that shifts opportunity and related touchpoint dates forward while preserving spacing. Update-only, fully reversible.";

const WEB_RESOURCE_FILES = [
    { file: "make-dates-perfect.html", displayName: "Make Dates Perfect", type: 1 },
];

const APP_DISPLAY_NAME = "Make Dates Perfect";
const APP_DESCRIPTION = "Demo-prep wizard that shifts opportunity and related touchpoint dates forward while preserving spacing. Update-only and fully reversible.";

const APP_ICON_SVG =
    '<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 32 32">' +
    '<rect width="32" height="32" rx="6" fill="#5B3FD6"/>' +
    '<rect x="7" y="9" width="18" height="15" rx="2" fill="#ffffff"/>' +
    '<rect x="7" y="9" width="18" height="5" rx="2" fill="#EFEAFF"/>' +
    '<rect x="11" y="6" width="2" height="5" rx="1" fill="#ffffff"/>' +
    '<rect x="19" y="6" width="2" height="5" rx="1" fill="#ffffff"/>' +
    '<circle cx="12" cy="18" r="1.6" fill="#5B3FD6"/>' +
    '<circle cx="16" cy="18" r="1.6" fill="#5B3FD6"/>' +
    '<circle cx="20" cy="18" r="1.6" fill="#7C5CFF"/>' +
    '</svg>';

const CACHE_FILE = path.join(__dirname, ".msal-cache.json");

let accessToken = null;

// ─── MSAL Authentication (Device Code Flow) ────────────────────

async function authenticate() {
    const config = {
        auth: {
            clientId: PUBLIC_CLIENT_ID,
            authority: "https://login.microsoftonline.com/organizations",
        },
    };

    const pca = new msal.PublicClientApplication(config);

    // Restore cache from local location only (per-environment login)
    if (fs.existsSync(CACHE_FILE)) {
        pca.getTokenCache().deserialize(fs.readFileSync(CACHE_FILE, "utf-8"));
    }

    // Try silent auth first
    const accounts = await pca.getTokenCache().getAllAccounts();
    if (accounts.length > 0) {
        try {
            const silentResponse = await pca.acquireTokenSilent({
                scopes: [`${CRM_BASE_URL}/.default`],
                account: accounts[0],
            });
            accessToken = silentResponse.accessToken;
            fs.writeFileSync(CACHE_FILE, pca.getTokenCache().serialize(), "utf-8");
            console.log(`✔ Authenticated silently as: ${accounts[0].username}\n`);
            return silentResponse;
        } catch (e) {
            console.log("  Silent auth failed, falling back to device code...\n");
        }
    }

    const deviceCodeRequest = {
        scopes: [`${CRM_BASE_URL}/.default`],
        deviceCodeCallback: (response) => {
            console.log("\n╔══════════════════════════════════════════════════════════╗");
            console.log("║  SIGN IN REQUIRED                                        ║");
            console.log("╚══════════════════════════════════════════════════════════╝");
            console.log(`\n${response.message}\n`);
        },
    };

    const response = await pca.acquireTokenByDeviceCode(deviceCodeRequest);
    accessToken = response.accessToken;
    fs.writeFileSync(CACHE_FILE, pca.getTokenCache().serialize(), "utf-8");
    console.log(`✔ Authenticated as: ${response.account.username}\n`);
    return response;
}

// ─── API Helpers ────────────────────────────────────────────────

async function crmRequest(method, relativeUrl, body) {
    const url = `${API_URL}${relativeUrl}`;
    const headers = {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        "Content-Type": "application/json",
    };

    const options = { method, headers };
    if (body) options.body = JSON.stringify(body);

    const res = await fetch(url, options);
    if (res.status === 204) return null;

    if (!res.ok) {
        const errText = await res.text();
        throw new Error(`API ${method} ${relativeUrl} → ${res.status}: ${errText}`);
    }

    if (method === "POST") {
        const entityIdHeader = res.headers.get("OData-EntityId");
        if (entityIdHeader) {
            const match = entityIdHeader.match(/\(([0-9a-f-]+)\)/i);
            if (match) return match[1];
        }
    }

    const text = await res.text();
    if (!text) return null;
    try {
        return JSON.parse(text);
    } catch (e) {
        return null;
    }
}

async function crmGet(url) { return crmRequest("GET", url); }
async function crmPost(url, body) { return crmRequest("POST", url, body); }
async function crmPatch(url, body) { return crmRequest("PATCH", url, body); }

// ─── Get Default Publisher ──────────────────────────────────────

async function getDefaultPublisher() {
    const defaultPub = await crmGet("/publishers?$filter=startswith(uniquename,'DefaultPublisherFor')&$select=publisherid,uniquename,customizationprefix");
    if (defaultPub.value && defaultPub.value.length > 0) return defaultPub.value[0];

    const allPubs = await crmGet("/publishers?$select=publisherid,uniquename,customizationprefix&$top=20&$orderby=createdon asc");
    if (allPubs.value) {
        const custom = allPubs.value.find(p =>
            p.uniquename !== "MicrosoftCorporation" &&
            p.uniquename !== "MicrosoftDynamics" &&
            !p.uniquename.startsWith("Microsoft") &&
            p.customizationprefix !== "none"
        );
        if (custom) return custom;
    }

    console.log("  → No custom publisher found. Creating one...");
    const newPubId = await crmPost("/publishers", {
        uniquename: "MakeDatesPerfectPublisher",
        friendlyname: "Make Dates Perfect Publisher",
        customizationprefix: "mdp",
        customizationoptionvalueprefix: 78710,
    });
    return { publisherid: newPubId, uniquename: "MakeDatesPerfectPublisher", customizationprefix: "mdp" };
}

// ─── Solution Management ────────────────────────────────────────

async function findOrCreateSolution(publisherId) {
    console.log("→ Checking if solution already exists...");
    const existing = await crmGet(`/solutions?$filter=uniquename eq '${SOLUTION_UNIQUE_NAME}'&$select=solutionid,uniquename`);

    if (existing.value && existing.value.length > 0) {
        console.log(`  ✔ Solution "${SOLUTION_UNIQUE_NAME}" already exists.`);
        return existing.value[0].solutionid;
    }

    console.log(`→ Creating solution "${SOLUTION_DISPLAY_NAME}"...`);
    let solutionId = await crmPost("/solutions", {
        uniquename: SOLUTION_UNIQUE_NAME,
        friendlyname: SOLUTION_DISPLAY_NAME,
        description: SOLUTION_DESCRIPTION,
        version: "1.0.0.0",
        "publisherid@odata.bind": `/publishers(${publisherId})`,
    });

    // If ID wasn't in header, query for it
    if (!solutionId) {
        const check = await crmGet(`/solutions?$filter=uniquename eq '${SOLUTION_UNIQUE_NAME}'&$select=solutionid`);
        if (check.value && check.value.length > 0) solutionId = check.value[0].solutionid;
    }

    console.log(`  ✔ Solution created (ID: ${solutionId})`);
    return solutionId;
}

// ─── Web Resource Management ────────────────────────────────────

function getWebResourceName(prefix, fileName) {
    return `${prefix}_/makedatesperfect/${fileName}`;
}

async function findWebResource(name) {
    const encoded = encodeURIComponent(name);
    const data = await crmGet(`/webresourceset?$filter=name eq '${encoded}'&$select=webresourceid,name`);
    if (data.value && data.value.length > 0) return data.value[0];
    return null;
}

async function createOrUpdateWebResource(prefix, fileDef) {
    const wrName = getWebResourceName(prefix, fileDef.file);
    const filePath = path.join(__dirname, fileDef.file);
    const content = fs.readFileSync(filePath, "utf-8");
    const contentBase64 = Buffer.from(content, "utf-8").toString("base64");

    const existing = await findWebResource(wrName);

    const wrBody = {
        name: wrName,
        displayname: fileDef.displayName,
        webresourcetype: fileDef.type,
        content: contentBase64,
    };

    if (existing) {
        console.log(`  ↻ Updating: ${wrName}`);
        await crmPatch(`/webresourceset(${existing.webresourceid})`, wrBody);
        return existing.webresourceid;
    } else {
        console.log(`  + Creating: ${wrName}`);
        let id = await crmPost("/webresourceset", wrBody);
        // If ID wasn't in header, query for it
        if (!id) {
            const check = await findWebResource(wrName);
            if (check) id = check.webresourceid;
        }
        return id;
    }
}

// ─── Add Component to Solution ──────────────────────────────────

async function addWebResourceToSolution(webResourceId) {
    try {
        await crmPost("/AddSolutionComponent", {
            ComponentId: webResourceId,
            ComponentType: 61,
            SolutionUniqueName: SOLUTION_UNIQUE_NAME,
            AddRequiredComponents: false,
            DoNotIncludeSubcomponents: false,
        });
    } catch (err) {
        if (err.message && err.message.includes("already exists")) return;
        console.log(`  ⚠ Warning adding to solution: ${err.message.substring(0, 120)}`);
    }
}

// ─── Model-Driven App (App Module + Site Map) ───────────────────

function getAppUniqueName(prefix) {
    return `${prefix}_makedatesperfect`;
}

async function createOrUpdateIcon(prefix) {
    const iconName = `${prefix}_/makedatesperfect/icon.svg`;
    const contentBase64 = Buffer.from(APP_ICON_SVG, "utf-8").toString("base64");
    const existing = await findWebResource(iconName);
    const body = { name: iconName, displayname: "Make Dates Perfect Icon", webresourcetype: 11, content: contentBase64 };
    if (existing) {
        console.log(`  ↻ Updating app icon: ${iconName}`);
        await crmPatch(`/webresourceset(${existing.webresourceid})`, body);
        return existing.webresourceid;
    }
    console.log(`  + Creating app icon: ${iconName}`);
    let id = await crmPost("/webresourceset", body);
    if (!id) {
        const check = await findWebResource(iconName);
        if (check) id = check.webresourceid;
    }
    return id;
}

function buildSiteMapXml(webResourceName) {
    return "<SiteMap>" +
        "<Area Id=\"mdp_area\" Title=\"Make Dates Perfect\" ShowGroups=\"true\">" +
        "<Group Id=\"mdp_group\" Title=\"Demo Tools\">" +
        "<SubArea Id=\"mdp_subarea\" Title=\"Make Dates Perfect\" " +
        "Url=\"$webresource:" + webResourceName + "\" />" +
        "</Group>" +
        "</Area>" +
        "</SiteMap>";
}

async function findAppModule(uniqueName) {
    // Newly created apps are unpublished until PublishAllXml, so a normal query won't see them.
    try {
        const un = await crmGet("/appmodules/Microsoft.Dynamics.CRM.RetrieveUnpublishedMultiple()?$select=appmoduleid,uniquename");
        const match = (un.value || []).find((a) => a.uniquename === uniqueName);
        if (match) return match;
    } catch (e) { /* fall back to published query */ }
    const data = await crmGet(`/appmodules?$filter=uniquename eq '${uniqueName}'&$select=appmoduleid,uniquename`);
    if (data.value && data.value.length > 0) return data.value[0];
    return null;
}

async function createOrUpdateAppModule(uniqueName, iconId) {
    const existing = await findAppModule(uniqueName);
    if (existing) {
        console.log(`  ↻ App module already exists: ${uniqueName} (reusing)`);
        return existing.appmoduleid;
    }
    console.log(`  + Creating app module: ${uniqueName}`);
    await crmPost("/appmodules", {
        name: APP_DISPLAY_NAME,
        uniquename: uniqueName,
        description: APP_DESCRIPTION,
        webresourceid: iconId,
        clienttype: 4,       // Unified Interface
        navigationtype: 0,   // Single session
    });
    const created = await findAppModule(uniqueName);
    return created ? created.appmoduleid : null;
}

async function findSiteMap(uniqueName) {
    try {
        const un = await crmGet("/sitemaps/Microsoft.Dynamics.CRM.RetrieveUnpublishedMultiple()?$select=sitemapid,sitemapnameunique");
        const match = (un.value || []).find((s) => s.sitemapnameunique === uniqueName);
        if (match) return match;
    } catch (e) { /* fall back to published query */ }
    const data = await crmGet(`/sitemaps?$filter=sitemapnameunique eq '${uniqueName}'&$select=sitemapid,sitemapnameunique`);
    if (data.value && data.value.length > 0) return data.value[0];
    return null;
}

async function createOrUpdateSiteMap(uniqueName, xml) {
    const existing = await findSiteMap(uniqueName);
    if (existing) {
        console.log(`  ↻ Site map already exists: ${uniqueName} (reusing)`);
        return existing.sitemapid;
    }
    console.log(`  + Creating site map: ${uniqueName}`);
    await crmPost("/sitemaps", { sitemapname: uniqueName, sitemapnameunique: uniqueName, sitemapxml: xml });
    const created = await findSiteMap(uniqueName);
    return created ? created.sitemapid : null;
}

async function addAppComponents(appId, components) {
    await crmPost("/AddAppComponents", { AppId: appId, Components: components });
}

async function addComponentToSolution(componentId, componentType) {
    try {
        await crmPost("/AddSolutionComponent", {
            ComponentId: componentId,
            ComponentType: componentType,
            SolutionUniqueName: SOLUTION_UNIQUE_NAME,
            AddRequiredComponents: false,
            DoNotIncludeSubcomponents: false,
        });
    } catch (err) {
        if (err.message && err.message.includes("already exists")) return;
        console.log(`  ⚠ Warning adding component (type ${componentType}) to solution: ${err.message.substring(0, 140)}`);
    }
}

async function validateApp(appId) {
    try {
        const res = await crmGet(`/ValidateApp(AppModuleId=${appId})`);
        const results = (res && res.AppValidationResponse && res.AppValidationResponse.ValidationIssueList) || res.ValidationIssueList || [];
        const errors = results.filter((r) => r.ValidationIssueType === 1 || r.errorType === "Error");
        if (errors.length) {
            console.log(`  ⚠ App validation reported ${errors.length} issue(s).`);
        } else {
            console.log("  ✔ App validation passed.");
        }
    } catch (err) {
        console.log(`  (Skipped app validation: ${err.message.substring(0, 100)})`);
    }
}

// ─── Publish ────────────────────────────────────────────────────

async function publishAll() {
    console.log("\n→ Publishing all customizations...");
    await crmPost("/PublishAllXml", {});
    console.log("  ✔ Published successfully!\n");
}

// ─── Main ───────────────────────────────────────────────────────

async function main() {
    console.log("╔══════════════════════════════════════════════════════════╗");
    console.log("║  Make Dates Perfect — Deploy to Dynamics 365            ║");
    console.log("╚══════════════════════════════════════════════════════════╝\n");

    // Step 1: Authenticate
    console.log("Step 1/5: Authenticating...");
    await authenticate();

    // Step 2: Get publisher
    console.log("Step 2/5: Fetching publisher...");
    const publisher = await getDefaultPublisher();
    const prefix = publisher.customizationprefix;
    console.log(`  ✔ Using publisher "${publisher.uniquename}" (prefix: ${prefix})\n`);

    // Step 3: Create/find solution
    console.log("Step 3/5: Setting up solution...");
    await findOrCreateSolution(publisher.publisherid);
    console.log();

    // Step 4: Upload web resources
    console.log("Step 4/6: Uploading web resources...");
    const webResourceIds = [];
    for (const fileDef of WEB_RESOURCE_FILES) {
        const wrId = await createOrUpdateWebResource(prefix, fileDef);
        webResourceIds.push(wrId);
    }
    console.log();

    // Adding to solution
    console.log("Adding web resources to solution...");
    for (const wrId of webResourceIds) {
        await addWebResourceToSolution(wrId);
    }
    console.log("  ✔ Web resources added to solution.\n");

    // Step 5: Build the model-driven app (app module + site map)
    console.log("Step 5/6: Building model-driven app...");
    const webResourceName = getWebResourceName(prefix, "make-dates-perfect.html");
    const appUniqueName = getAppUniqueName(prefix);
    const iconId = await createOrUpdateIcon(prefix);
    await addWebResourceToSolution(iconId);
    const appId = await createOrUpdateAppModule(appUniqueName, iconId);
    const siteMapId = await createOrUpdateSiteMap(appUniqueName, buildSiteMapXml(webResourceName));
    if (!appId || !siteMapId) throw new Error(`Could not resolve app/sitemap ids (app=${appId}, sitemap=${siteMapId}).`);

    console.log("  → Wiring site map into the app...");
    await addAppComponents(appId, [
        { "@odata.type": "Microsoft.Dynamics.CRM.sitemap", sitemapid: siteMapId },
    ]);

    console.log("  → Packaging app + site map into the solution...");
    await addComponentToSolution(appId, 80);     // 80 = Model-driven App
    await addComponentToSolution(siteMapId, 62); // 62 = Site Map
    await validateApp(appId);
    console.log();

    // Step 6: Publish
    console.log("Step 6/6: Publishing...");
    await publishAll();

    console.log("══════════════════════════════════════════════════════════");
    console.log("  DEPLOYMENT COMPLETE!");
    console.log(`  Solution: ${SOLUTION_DISPLAY_NAME}`);
    console.log(`  Environment: ${CRM_BASE_URL}`);
    console.log(`  Model-driven App: ${APP_DISPLAY_NAME}`);
    console.log("  Web Resource: " + webResourceName);
    console.log("  Open the app: " + CRM_BASE_URL + "/main.aspx?appid=" + appId);
    console.log("══════════════════════════════════════════════════════════\n");
}

main().catch((err) => {
    console.error("\n✖ Deployment failed:", err.message);
    process.exit(1);
});
