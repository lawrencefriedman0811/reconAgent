/* global Office */

function readCommonPayload() {
  return {
    entity: document.getElementById("entity").value.trim(),
    period: document.getElementById("period").value.trim(),
  };
}

function apiBaseUrl() {
  return document.getElementById("apiBaseUrl").value.trim().replace(/\/$/, "");
}

function setResponse(value) {
  const text =
    typeof value === "string" ? value : JSON.stringify(value, null, 2);
  document.getElementById("responseOutput").textContent = text;
}

async function postJson(path, body) {
  const response = await fetch(`${apiBaseUrl()}${path}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });

  const text = await response.text();
  let payload;
  try {
    payload = JSON.parse(text);
  } catch {
    payload = { raw: text };
  }

  if (!response.ok) {
    throw new Error(JSON.stringify(payload, null, 2));
  }
  return payload;
}

async function onReconcile() {
  setResponse("Running /reconcile...");
  try {
    const result = await postJson("/reconcile", readCommonPayload());
    setResponse(result);
  } catch (error) {
    setResponse(`Reconcile failed:\n${error.message}`);
  }
}

async function onValidate() {
  setResponse("Running /validate...");
  try {
    const result = await postJson("/validate", readCommonPayload());
    setResponse(result);
  } catch (error) {
    setResponse(`Validate failed:\n${error.message}`);
  }
}

async function onWriteback() {
  setResponse("Running /writeback...");
  try {
    const updatesRaw = document.getElementById("updatesJson").value.trim() || "[]";
    const updates = JSON.parse(updatesRaw);
    if (!Array.isArray(updates)) {
      throw new Error("Writeback updates must be a JSON array.");
    }

    const result = await postJson("/writeback", {
      ...readCommonPayload(),
      updates,
    });
    setResponse(result);
  } catch (error) {
    setResponse(`Writeback failed:\n${error.message}`);
  }
}

function init() {
  document.getElementById("apiBaseUrl").value = window.location.origin;
  document.getElementById("reconcileBtn").addEventListener("click", onReconcile);
  document.getElementById("validateBtn").addEventListener("click", onValidate);
  document.getElementById("writebackBtn").addEventListener("click", onWriteback);
}

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(init);
} else {
  document.addEventListener("DOMContentLoaded", init);
}
