/**
 * MCS CoWorker API Client
 *
 * The app only ever runs inside the desktop shell (pywebview serving from
 * http://127.0.0.1:7842/). Every call hits the real Flask backend. No mock
 * fallbacks — callers handle fetch failures by showing an empty state or
 * toast, which is the right UX anyway.
 */

const BASE = "http://127.0.0.1:7842";

async function apiFetch<T>(path: string, opts?: RequestInit): Promise<T> {
  const res = await fetch(`${BASE}${path}`, {
    headers: { "Content-Type": "application/json" },
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
export async function fetchMemory(query = "recent client interactions", limit = 50) {
  return apiFetch(`/api/memory?q=${encodeURIComponent(query)}&limit=${limit}`);
}

export async function deleteMemory(recordId: string) {
  return apiFetch(`/api/memory/${recordId}`, { method: "DELETE" });
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
export async function sendChatMessage(messages: { role: string; content: string }[]) {
  return apiFetch("/api/chat", { method: "POST", body: JSON.stringify({ messages }) });
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
