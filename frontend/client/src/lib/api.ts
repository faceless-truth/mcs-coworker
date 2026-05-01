/**
 * MCS CoWorker API Client
 *
 * The app only ever runs inside the desktop shell (pywebview serving from
 * http://127.0.0.1:7842/). Every call hits the real Flask backend. No mock
 * fallbacks — callers handle fetch failures by showing an empty state or
 * toast, which is the right UX anyway.
 */

const BASE = "http://127.0.0.1:7842";

function getApiToken(): string {
  return (window as any).__API_TOKEN__ || "";
}

async function apiFetch<T>(path: string, opts?: RequestInit): Promise<T> {
  const token = getApiToken();
  const res = await fetch(`${BASE}${path}`, {
    headers: {
      "Content-Type": "application/json",
      ...(token ? { Authorization: `Bearer ${token}` } : {}),
    },
    ...opts,
  });
  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    throw new Error(body.error || `HTTP ${res.status}`);
  }
  const json = await res.json();
  return json.data ?? json;
}

// ── Health ────────────────────────────────────────────────────────────────────
export async function fetchHealth() {
  return apiFetch("/api/health");
}

// ── System Status ─────────────────────────────────────────────────────────────
export async function fetchSystemStatus() {
  return apiFetch("/api/system/status");
}

// ── Plugins ───────────────────────────────────────────────────────────────────
export async function fetchPlugins() {
  return apiFetch("/api/plugins");
}

export async function enablePlugin(pluginId: string, enabled: boolean) {
  const endpoint = enabled ? "enable" : "disable";
  return apiFetch(`/api/plugins/${pluginId}/${endpoint}`, {
    method: "POST",
    body: JSON.stringify({ enabled }),
  });
}

export async function runPlugin(pluginId: string) {
  return apiFetch(`/api/plugins/${pluginId}/run`, { method: "POST" });
}

export async function deletePlugin(pluginId: string) {
  return apiFetch(`/api/plugins/${pluginId}`, { method: "DELETE" });
}

// ── Activity ──────────────────────────────────────────────────────────────────
export async function fetchActivity(limit = 100) {
  return apiFetch(`/api/activity?limit=${limit}`);
}

// ── Approvals ─────────────────────────────────────────────────────────────────
export async function fetchApprovals() {
  return apiFetch("/api/approvals");
}

export async function approveAction(actionId: string) {
  return apiFetch(`/api/approvals/${actionId}/approve`, { method: "POST" });
}

export async function rejectAction(actionId: string) {
  return apiFetch(`/api/approvals/${actionId}/reject`, { method: "POST" });
}

// ── Memory ────────────────────────────────────────────────────────────────────
export async function fetchMemory(query = "", limit = 50) {
  const q = query ? `&q=${encodeURIComponent(query)}` : "";
  return apiFetch(`/api/memory?limit=${limit}${q}`);
}

export async function fetchMemoryStats() {
  return apiFetch("/api/memory/stats");
}

export interface MemoryDetail {
  id: string;
  content: string;
  metadata: Record<string, any>;
  collection: string;
}

export async function fetchMemoryDetail(recordId: string): Promise<MemoryDetail | null> {
  // Bypass apiFetch's `json.data ?? json` unwrap so we can read `entry`
  // directly off the response body.
  const token = getApiToken();
  const res = await fetch(`${BASE}/api/memory/${encodeURIComponent(recordId)}`, {
    headers: {
      "Content-Type": "application/json",
      ...(token ? { Authorization: `Bearer ${token}` } : {}),
    },
  });
  if (!res.ok) return null;
  const json = await res.json();
  return json.entry ?? null;
}

export async function deleteMemory(recordId: string) {
  return apiFetch(`/api/memory/${recordId}`, { method: "DELETE" });
}
export async function updateMemory(recordId: string, fields: { client_name?: string; entity_name?: string }) {
  return apiFetch(`/api/memory/${recordId}`, {
    method: "PATCH",
    body: JSON.stringify(fields),
  });
}

// ── Events ────────────────────────────────────────────────────────────────────
export async function fetchEvents(limit = 50) {
  return apiFetch(`/api/events?limit=${limit}`);
}

// ── KPI ───────────────────────────────────────────────────────────────────────
export async function fetchKPI() {
  return apiFetch("/api/kpi");
}

// ── Usage ─────────────────────────────────────────────────────────────────────
export async function fetchUsage() {
  return apiFetch("/api/usage");
}

// ── Settings ──────────────────────────────────────────────────────────────────
export async function fetchSettings() {
  return apiFetch("/api/settings");
}

export async function saveSettings(settings: Record<string, string>) {
  return apiFetch("/api/settings", { method: "POST", body: JSON.stringify(settings) });
}

export async function testConnection(service: "xpm" | "fusesign" | "teams") {
  return apiFetch(`/api/settings/test/${service}`, { method: "POST" });
}

export async function testSharepointConnection() {
  return apiFetch("/api/sharepoint/test", { method: "POST" });
}

// ── Xero OAuth ────────────────────────────────────────────────────────────────
export async function fetchXeroStatus() {
  return apiFetch("/api/xero/status");
}

export async function startXeroAuth() {
  return apiFetch("/api/xero/start-auth", { method: "POST" });
}

export async function disconnectXero() {
  return apiFetch("/api/xero/disconnect", { method: "POST" });
}

// ── Email Rules ───────────────────────────────────────────────────────────────
export async function fetchRules() {
  return apiFetch("/api/rules");
}

export async function saveRule(rule: Record<string, any>) {
  return apiFetch("/api/rules", { method: "POST", body: JSON.stringify(rule) });
}

export async function deleteRule(ruleId: number) {
  return apiFetch(`/api/rules/${ruleId}`, { method: "DELETE" });
}

// ── Staff & Notify ────────────────────────────────────────────────────────────
export async function fetchStaff() {
  return apiFetch("/api/staff");
}

export async function saveStaff(staff: Record<string, any>) {
  return apiFetch("/api/staff", { method: "POST", body: JSON.stringify(staff) });
}

export async function deleteStaff(staffId: number) {
  return apiFetch(`/api/staff/${staffId}`, { method: "DELETE" });
}

// ── Links & Forms ─────────────────────────────────────────────────────────────
export async function fetchLinks() {
  return apiFetch("/api/links");
}

export async function saveLink(link: Record<string, any>) {
  return apiFetch("/api/links", { method: "POST", body: JSON.stringify(link) });
}

export async function deleteLink(linkId: number) {
  return apiFetch(`/api/links/${linkId}`, { method: "DELETE" });
}

// ── Chat ──────────────────────────────────────────────────────────────────────
export interface ChatFileRef {
  id: string;
  name: string;
  type: string;
  size: number;
}

export async function sendChatMessage(
  messages: { role: string; content: string }[],
  agentId: string = "plugin_builder",
  files: ChatFileRef[] = [],
  clientName: string | null = null,
  entityName: string | null = null,
) {
  return apiFetch("/api/chat", {
    method: "POST",
    body: JSON.stringify({
      messages,
      agent_id: agentId,
      files,
      client_name: clientName,
      entity_name: entityName,
    }),
  });
}

export async function fetchClientNames(): Promise<string[]> {
  try {
    const token = getApiToken();
    const res = await fetch(`${BASE}/api/clients/names`, {
      headers: {
        "Content-Type": "application/json",
        ...(token ? { Authorization: `Bearer ${token}` } : {}),
      },
    });
    if (!res.ok) return [];
    const json = await res.json();
    return Array.isArray(json.clients) ? json.clients : [];
  } catch {
    return [];
  }
}

export async function fetchAgents() {
  const token = getApiToken();
  const res = await fetch(`${BASE}/api/agents`, {
    headers: {
      "Content-Type": "application/json",
      ...(token ? { Authorization: `Bearer ${token}` } : {}),
    },
  });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const json = await res.json();
  return json.agents ?? [];
}

export async function uploadChatFile(file: File): Promise<ChatFileRef> {
  const form = new FormData();
  form.append("file", file);
  const token = getApiToken();
  const res = await fetch(`${BASE}/api/chat/upload`, {
    method: "POST",
    body: form,
    headers: token ? { Authorization: `Bearer ${token}` } : {},
  });
  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    throw new Error(body.error || `Upload failed (HTTP ${res.status})`);
  }
  const json = await res.json();
  return json.file;
}

export async function fetchChatHistory() {
  return apiFetch("/api/chat/history");
}

export async function clearChatHistory() {
  return apiFetch("/api/chat/history", { method: "DELETE" });
}

// ── Knowledge Base ────────────────────────────────────────────────────────────
export interface KnowledgeEntry {
  id: number;
  category: string;
  title: string;
  content: string;
  enabled: number;
  created_at?: string;
  updated_at?: string;
}

export async function fetchKnowledge(): Promise<KnowledgeEntry[]> {
  return apiFetch<KnowledgeEntry[]>("/api/knowledge");
}

export async function createKnowledge(entry: Omit<KnowledgeEntry, "id" | "created_at" | "updated_at">) {
  return apiFetch("/api/knowledge", { method: "POST", body: JSON.stringify(entry) });
}

export async function updateKnowledge(id: number, entry: Partial<KnowledgeEntry>) {
  return apiFetch(`/api/knowledge/${id}`, { method: "PUT", body: JSON.stringify(entry) });
}

export async function deleteKnowledge(id: number) {
  return apiFetch(`/api/knowledge/${id}`, { method: "DELETE" });
}

// ── BAS Clients ───────────────────────────────────────────────────────────────
export interface BasClient {
  id: number;
  client_name: string;
  entity_name: string | null;
  abn: string | null;
  frequency: string | null;
  client_email: string | null;
  last_data_received: string | null;
  last_reminder_sent: string | null;
  status: string | null;
  notes: string | null;
  created_at?: string | null;
}

export async function fetchBasClients(): Promise<BasClient[]> {
  return apiFetch<BasClient[]>("/api/bas-clients");
}

export async function createBasClient(row: Partial<BasClient>): Promise<BasClient> {
  return apiFetch<BasClient>("/api/bas-clients", {
    method: "POST",
    body: JSON.stringify(row),
  });
}

export async function updateBasClient(id: number, row: Partial<BasClient>): Promise<BasClient> {
  return apiFetch<BasClient>(`/api/bas-clients/${id}`, {
    method: "PUT",
    body: JSON.stringify(row),
  });
}

export async function deleteBasClient(id: number) {
  return apiFetch(`/api/bas-clients/${id}`, { method: "DELETE" });
}

// ── Lessons ───────────────────────────────────────────────────────────────────
export async function fetchLessons() {
  return apiFetch("/api/lessons");
}

export async function addLesson(lesson: string, source = "") {
  return apiFetch("/api/lessons", { method: "POST", body: JSON.stringify({ lesson, source }) });
}

export async function deleteLesson(lessonId: number) {
  return apiFetch(`/api/lessons/${lessonId}`, { method: "DELETE" });
}

// ── SharePoint Indexer ────────────────────────────────────────────────────────
export interface SharepointIndexStatus {
  status: "idle" | "running" | "error";
  last_run: string;
  last_count: number;
  last_error: string;
}

export async function fetchSharepointIndexStatus(): Promise<SharepointIndexStatus> {
  return apiFetch<SharepointIndexStatus>("/api/sharepoint/index/status");
}

export async function runSharepointIndex(incremental = true) {
  return apiFetch("/api/sharepoint/index/run", {
    method: "POST",
    body: JSON.stringify({ incremental }),
  });
}
