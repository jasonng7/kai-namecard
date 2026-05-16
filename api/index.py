import os
import sys

from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, JSONResponse, Response

sys.path.append(os.path.dirname(__file__))

from _shared import (
    build_hubspot_workbook,
    download_drive_file,
    exchange_oauth_code,
    extract_card_from_bytes,
    fetch_google_email,
    find_existing_scan,
    get_drive_connection,
    get_drive_file,
    google_auth_url,
    insert_scan,
    list_drive_images,
    list_scans,
    parse_drive_folder_id,
    refresh_google_access_token,
    unix_filename_stamp,
    upsert_drive_connection,
    verify_state,
)


app = FastAPI()

INDEX_HTML = """
<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Name Card Drive Scanner</title>
    <style>
      * { box-sizing: border-box; }
      body { margin: 0; font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; color: #17202a; background: #fff; }
      main { width: min(1080px, calc(100% - 32px)); margin: 0 auto; padding: 32px 0 56px; }
      header { display: flex; align-items: flex-end; justify-content: space-between; gap: 24px; padding-bottom: 24px; border-bottom: 1px solid #d9e0ea; }
      h1 { margin: 0; font-size: clamp(28px, 4vw, 44px); line-height: 1.05; letter-spacing: 0; }
      h2 { margin: 0 0 12px; font-size: 18px; letter-spacing: 0; }
      p { margin: 8px 0 0; color: #657184; line-height: 1.55; }
      button, input { font: inherit; }
      button { min-height: 42px; border: 1px solid #0f766e; background: #0f766e; color: #fff; padding: 0 16px; border-radius: 8px; cursor: pointer; white-space: nowrap; }
      button.secondary { background: #fff; color: #115e59; }
      button:disabled { cursor: not-allowed; opacity: .55; }
      input[type="text"] { width: 100%; min-height: 42px; border: 1px solid #d9e0ea; border-radius: 8px; background: #f7f9fc; color: #17202a; padding: 0 12px; }
      label { display: block; margin: 12px 0 6px; color: #354052; font-size: 14px; font-weight: 650; }
      .status { display: inline-flex; align-items: center; min-height: 34px; padding: 0 12px; border: 1px solid #d9e0ea; border-radius: 999px; color: #657184; background: #fff; font-size: 14px; }
      .status.connected { color: #115e59; border-color: #9ad6cf; background: #eef8f7; }
      .workflow { display: grid; grid-template-columns: 320px 1fr; gap: 28px; margin-top: 28px; }
      .panel, .metric { border: 1px solid #d9e0ea; border-radius: 8px; background: #fff; }
      .panel { padding: 18px; }
      .panel + .panel { margin-top: 14px; }
      .actions { display: flex; flex-wrap: wrap; gap: 10px; margin-top: 14px; }
      .check-row { display: flex; align-items: center; gap: 8px; margin-top: 12px; color: #657184; font-size: 14px; }
      .summary { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px; margin-bottom: 16px; }
      .metric { padding: 14px; }
      .metric strong { display: block; font-size: 26px; }
      .metric span { color: #657184; font-size: 13px; }
      .message { min-height: 44px; padding: 12px 0; color: #657184; }
      .message.error { color: #b42318; }
      table { width: 100%; border-collapse: collapse; table-layout: fixed; }
      th, td { border-bottom: 1px solid #d9e0ea; padding: 11px 8px; text-align: left; vertical-align: top; overflow-wrap: anywhere; font-size: 14px; }
      th { color: #354052; background: #f7f9fc; font-weight: 700; }
      .empty { border: 1px dashed #d9e0ea; border-radius: 8px; padding: 28px; color: #657184; text-align: center; }
      .tag { display: inline-flex; min-height: 24px; align-items: center; padding: 0 8px; border-radius: 999px; background: #f3f5f8; color: #526071; font-size: 12px; }
      @media (max-width: 820px) { header, .workflow { display: block; } header .status { margin-top: 18px; } aside { margin-bottom: 18px; } .summary { grid-template-columns: repeat(2, minmax(0, 1fr)); } table { min-width: 760px; } .table-wrap { overflow-x: auto; } }
    </style>
  </head>
  <body>
    <main>
      <header>
        <div>
          <h1>Name Card Drive Scanner</h1>
          <p>Connect Google Drive, sync an image folder, skip cards already scanned, and export a HubSpot-ready workbook.</p>
        </div>
        <div id="connectionStatus" class="status">Checking Drive connection</div>
      </header>
      <div class="workflow">
        <aside>
          <div class="panel">
            <h2>Google Drive</h2>
            <p id="connectionText">Connect the Google account that owns or can view the folder.</p>
            <div class="actions">
              <button id="connectButton">Connect Drive</button>
              <button id="refreshButton" class="secondary">Refresh</button>
            </div>
          </div>
          <div class="panel">
            <h2>Folder Sync</h2>
            <label for="folderInput">Drive folder link or ID</label>
            <input id="folderInput" type="text" placeholder="https://drive.google.com/drive/folders/..." />
            <label class="check-row"><input id="forceInput" type="checkbox" /> Rescan even if cached</label>
            <div class="actions">
              <button id="syncButton">Sync Folder</button>
              <button id="exportButton" class="secondary">Export Excel</button>
            </div>
          </div>
        </aside>
        <section>
          <div class="summary">
            <div class="metric"><strong id="foundMetric">0</strong><span>Images found</span></div>
            <div class="metric"><strong id="processedMetric">0</strong><span>New scans</span></div>
            <div class="metric"><strong id="skippedMetric">0</strong><span>Skipped</span></div>
            <div class="metric"><strong id="failedMetric">0</strong><span>Failed</span></div>
          </div>
          <div id="message" class="message">Ready.</div>
          <div id="results" class="empty">No scans loaded yet.</div>
        </section>
      </div>
    </main>
    <script>
      const els = {
        connectionStatus: document.getElementById("connectionStatus"),
        connectionText: document.getElementById("connectionText"),
        connectButton: document.getElementById("connectButton"),
        refreshButton: document.getElementById("refreshButton"),
        syncButton: document.getElementById("syncButton"),
        exportButton: document.getElementById("exportButton"),
        folderInput: document.getElementById("folderInput"),
        forceInput: document.getElementById("forceInput"),
        foundMetric: document.getElementById("foundMetric"),
        processedMetric: document.getElementById("processedMetric"),
        skippedMetric: document.getElementById("skippedMetric"),
        failedMetric: document.getElementById("failedMetric"),
        message: document.getElementById("message"),
        results: document.getElementById("results"),
      };
      const workspaceKey = "kai-namecard-workspace-id";
      let workspaceId = localStorage.getItem(workspaceKey);
      if (!workspaceId) { workspaceId = crypto.randomUUID(); localStorage.setItem(workspaceKey, workspaceId); }
      function setMessage(text, isError = false) { els.message.textContent = text; els.message.classList.toggle("error", isError); }
      function setBusy(isBusy) { [els.connectButton, els.refreshButton, els.syncButton, els.exportButton].forEach((b) => b.disabled = isBusy); }
      function escapeHtml(value) { return String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&#039;"); }
      async function fetchJson(url, options) { const response = await fetch(url, options); const data = await response.json(); if (!response.ok) throw new Error(data.error || "Request failed."); return data; }
      function renderRows(rows) {
        if (!rows.length) { els.results.className = "empty"; els.results.textContent = "No scanned cards yet."; return; }
        els.results.className = "table-wrap";
        els.results.innerHTML = `<table><thead><tr><th>File</th><th>Name</th><th>Email</th><th>Phone</th><th>Company</th><th>Status</th></tr></thead><tbody>${rows.map((row) => {
          const ex = row.extraction || {};
          const fullName = [ex.first_name, ex.last_name].filter(Boolean).join(" ");
          const unclear = ex.unclear_fields?.length ? `Unclear: ${ex.unclear_fields.join(", ")}` : "OK";
          return `<tr><td>${escapeHtml(row.file_name || row.fileName || "")}</td><td>${escapeHtml(fullName)}</td><td>${escapeHtml(ex.email || "")}</td><td>${escapeHtml(ex.phone_number || "")}</td><td>${escapeHtml(ex.company_name || "")}</td><td><span class="tag">${escapeHtml(unclear)}</span></td></tr>`;
        }).join("")}</tbody></table>`;
      }
      async function checkStatus() {
        const data = await fetchJson(`/api/status?workspace_id=${encodeURIComponent(workspaceId)}`);
        els.connectionStatus.classList.toggle("connected", data.connected);
        els.connectionStatus.textContent = data.connected ? "Drive connected" : "Drive not connected";
        els.connectionText.textContent = data.connected ? `Connected as ${data.google_email || "Google Drive user"}.` : "Connect the Google account that owns or can view the folder.";
      }
      async function loadScans() {
        const data = await fetchJson(`/api/scans?workspace_id=${encodeURIComponent(workspaceId)}`);
        renderRows(data.records.map((record) => ({ file_name: record.file_name, extraction: record.extraction })));
      }
      els.connectButton.addEventListener("click", async () => {
        try { setBusy(true); setMessage("Opening Google sign-in..."); const data = await fetchJson(`/api/auth_url?workspace_id=${encodeURIComponent(workspaceId)}`); window.location.href = data.auth_url; }
        catch (error) { setMessage(error.message, true); setBusy(false); }
      });
      els.refreshButton.addEventListener("click", async () => {
        try { setBusy(true); await checkStatus(); await loadScans(); setMessage("Status refreshed."); }
        catch (error) { setMessage(error.message, true); }
        finally { setBusy(false); }
      });
      els.syncButton.addEventListener("click", async () => {
        try {
          setBusy(true); setMessage("Scanning Drive folder. This can take a little while for large folders.");
          const data = await fetchJson("/api/sync", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ workspace_id: workspaceId, folder: els.folderInput.value, force: els.forceInput.checked }) });
          els.foundMetric.textContent = data.found; els.processedMetric.textContent = data.processed.length; els.skippedMetric.textContent = data.skipped.length; els.failedMetric.textContent = data.failed.length;
          renderRows([...data.processed, ...data.skipped]); setMessage(`Synced ${data.folder_name || "Drive folder"}.`);
        } catch (error) { setMessage(error.message, true); }
        finally { setBusy(false); }
      });
      els.exportButton.addEventListener("click", () => { window.location.href = `/api/export?workspace_id=${encodeURIComponent(workspaceId)}`; });
      (async function boot() { try { await checkStatus(); await loadScans(); setMessage("Ready."); } catch (error) { setMessage(error.message, true); } })();
    </script>
  </body>
</html>
"""


def error_response(exc, status_code=400):
    return JSONResponse({"error": str(exc)}, status_code=status_code)


@app.get("/")
def home():
    return HTMLResponse(INDEX_HTML)


@app.get("/api/auth_url")
def auth_url(request: Request, workspace_id: str):
    try:
        return {"auth_url": google_auth_url(request, workspace_id)}
    except Exception as exc:
        return error_response(exc)


@app.get("/api/oauth_callback")
def oauth_callback(request: Request, code: str = "", state: str = "", error: str = ""):
    try:
        if error:
            raise RuntimeError(error)
        if not code:
            raise RuntimeError("Missing OAuth code.")

        workspace_id = verify_state(state)
        token = exchange_oauth_code(request, code)
        refresh_token = token.get("refresh_token")
        if not refresh_token:
            raise RuntimeError(
                "Google did not return a refresh token. Reconnect and approve offline access."
            )

        email = fetch_google_email(token["access_token"])
        scopes = token.get("scope", "").split()
        upsert_drive_connection(workspace_id, email, refresh_token, scopes)

        return HTMLResponse(
            """
<!doctype html>
<html>
  <head>
    <meta charset="utf-8">
    <title>Google Drive connected</title>
    <style>
      body { font-family: system-ui, sans-serif; margin: 40px; color: #17202a; }
      a { color: #146c94; }
    </style>
  </head>
  <body>
    <h1>Google Drive connected</h1>
    <p>You can return to the name card scanner and sync a folder now.</p>
    <p><a href="/">Back to scanner</a></p>
  </body>
</html>
            """.strip()
        )
    except Exception as exc:
        return HTMLResponse(
            f"<h1>Connection failed</h1><p>{str(exc)}</p><p><a href='/'>Back</a></p>",
            status_code=400,
        )


@app.get("/api/status")
def status(workspace_id: str):
    try:
        connection = get_drive_connection(workspace_id)
        return {
            "connected": bool(connection),
            "google_email": connection.get("google_email") if connection else None,
            "updated_at": connection.get("updated_at") if connection else None,
        }
    except Exception as exc:
        return error_response(exc)


@app.post("/api/sync")
async def sync(request: Request):
    try:
        body = await request.json()
        workspace_id = body.get("workspace_id")
        folder_input = body.get("folder")
        force = bool(body.get("force"))
        if not workspace_id:
            raise RuntimeError("workspace_id is required.")

        connection = get_drive_connection(workspace_id)
        if not connection:
            raise RuntimeError("Connect Google Drive before syncing.")

        folder_id = parse_drive_folder_id(folder_input)
        access_token = refresh_google_access_token(connection["refresh_token"])
        folder = get_drive_file(access_token, folder_id, "id,name,mimeType,webViewLink")
        if folder.get("mimeType") != "application/vnd.google-apps.folder":
            raise RuntimeError("The Drive link must point to a folder.")

        files = list_drive_images(access_token, folder_id)
        processed = []
        skipped = []
        failed = []

        for file_info in files:
            from _shared import source_fingerprint

            fingerprint = source_fingerprint(file_info)
            existing = None if force else find_existing_scan(workspace_id, fingerprint)
            if existing:
                skipped.append(
                    {
                        "file_id": file_info["id"],
                        "file_name": file_info["name"],
                        "reason": "already_scanned",
                        "extraction": existing["extraction"],
                    }
                )
                continue

            try:
                image_bytes = download_drive_file(access_token, file_info["id"])
                extraction = extract_card_from_bytes(image_bytes, file_info["mimeType"])
                row = insert_scan(
                    workspace_id,
                    folder_id,
                    folder.get("name"),
                    file_info,
                    fingerprint,
                    extraction,
                )
                processed.append(
                    {
                        "file_id": file_info["id"],
                        "file_name": file_info["name"],
                        "extraction": row["extraction"],
                    }
                )
            except Exception as exc:
                failed.append(
                    {
                        "file_id": file_info.get("id"),
                        "file_name": file_info.get("name"),
                        "error": str(exc),
                    }
                )

        return {
            "folder_id": folder_id,
            "folder_name": folder.get("name"),
            "found": len(files),
            "processed": processed,
            "skipped": skipped,
            "failed": failed,
        }
    except Exception as exc:
        return error_response(exc)


@app.get("/api/export")
def export(workspace_id: str):
    try:
        records = list_scans(workspace_id)
        workbook = build_hubspot_workbook(records)
        filename = f"namecards-hubspot-{unix_filename_stamp()}.xlsx"
        return Response(
            workbook,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
    except Exception as exc:
        return error_response(exc)


@app.get("/api/scans")
def scans(workspace_id: str):
    try:
        rows = list_scans(workspace_id)
        return {
            "records": [
                {
                    "file_name": row["drive_file_name"],
                    "folder_name": row.get("drive_folder_name"),
                    "processed_at": row.get("processed_at"),
                    "extraction": row["extraction"],
                }
                for row in rows
            ]
        }
    except Exception as exc:
        return error_response(exc)
