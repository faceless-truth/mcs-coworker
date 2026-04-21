/**
 * MCS CoWorker API Client
 * Auto-detects environment:
 *   - Desktop app (running on port 7842): calls http://127.0.0.1:7842/api/...
 *   - Web demo (browser):                returns mock data
 */

import * as mock from "./mockData";

// Detect if running inside the desktop app (Flask server on 7842)
// Works for: direct Flask serve, pywebview, and Electron (file:// or app://)
const IS_DESKTOP =
  typeof window !== "undefined" &&
  (
    window.location.port === "7842" ||
    (window as any).__pywebview__ !== undefined ||
    (window as any).__ELECTRON__ !== undefined ||
    window.location.protocol === "file:" ||
    window.location.protocol === "app:" ||
    (window.location.hostname === "127.0.0.1" && window.location.port === "7842")
  );

const BASE = IS_DESKTOP ? "http://127.0.0.1:7842" : "";

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
  if (!IS_DESKTOP) {
    return { status: "ok", uptime: "6h 42m", version: "2.4.1", fast_model: "claude-haiku-4-5", reasoning_model: "claude-sonnet-4-6" };
  }
  return apiFetch("/api/health");
}

// ── System Status ─────────────────────────────────────────────────────────────
export async function fetchSystemStatus() {
  if (!IS_DESKTOP) {
    return { heartbeat: "Active", heartbeatTick: 2847, fastModel: "claude-haiku-4-5", reasoningModel: "claude-sonnet-4-6", memoryRecords: 1247, costToday: "$1.84", uptime: "6h 42m", version: "2.4.1", updateAvailable: false };
  }
  return apiFetch("/api/system/status");
}

// ── Plugins ───────────────────────────────────────────────────────────────────
export async function fetchPlugins() {
  if (!IS_DESKTOP) return mock.mockPlugins;
  return apiFetch("/api/plugins");
}

export async function enablePlugin(pluginId: string, enabled: boolean) {
  if (!IS_DESKTOP) return { plugin_id: pluginId, enabled };
  const endpoint = enabled ? "enable" : "disable";
  return apiFetch(`/api/plugins/${pluginId}/${endpoint}`, { method: "POST", body: JSON.stringify({ enabled }) });
}

export async function runPlugin(pluginId: string) {
  if (!IS_DESKTOP) return { plugin_id: pluginId, triggered: true };
  return apiFetch(`/api/plugins/${pluginId}/run`, { method: "POST" });
}

export async function deletePlugin(pluginId: string) {
  if (!IS_DESKTOP) return { deleted: pluginId };
  return apiFetch(`/api/plugins/${pluginId}`, { method: "DELETE" });
}

// ── Activity ──────────────────────────────────────────────────────────────────
export async function fetchActivity(limit = 100) {
  if (!IS_DESKTOP) return mock.mockActivityLog;
  return apiFetch(`/api/activity?limit=${limit}`);
}

// ── Approvals ─────────────────────────────────────────────────────────────────
export async function fetchApprovals() {
  if (!IS_DESKTOP) return mock.mockApprovals;
  return apiFetch("/api/approvals");
}

export async function approveAction(actionId: string) {
  if (!IS_DESKTOP) return { action_id: actionId, approved: true };
  return apiFetch(`/api/approvals/${actionId}/approve`, { method: "POST" });
}

export async function rejectAction(actionId: string) {
  if (!IS_DESKTOP) return { action_id: actionId, rejected: true };
  return apiFetch(`/api/approvals/${actionId}/reject`, { method: "POST" });
}

// ── Memory ────────────────────────────────────────────────────────────────────
export async function fetchMemory(query = "recent client interactions", limit = 50) {
  if (!IS_DESKTOP) return mock.mockMemory;
  return apiFetch(`/api/memory?q=${encodeURIComponent(query)}&limit=${limit}`);
}

export async function deleteMemory(recordId: string) {
  if (!IS_DESKTOP) return { deleted: recordId };
  return apiFetch(`/api/memory/${recordId}`, { method: "DELETE" });
}

// ── Events ────────────────────────────────────────────────────────────────────
export async function fetchEvents(limit = 50) {
  if (!IS_DESKTOP) return mock.mockEvents;
  return apiFetch(`/api/events?limit=${limit}`);
}

// ── KPI ───────────────────────────────────────────────────────────────────────
export async function fetchKPI() {
  if (!IS_DESKTOP) return mock.mockKpis;
  return apiFetch("/api/kpi");
}

// ── Usage ─────────────────────────────────────────────────────────────────────
export async function fetchUsage() {
  if (!IS_DESKTOP) return mock.mockUsageData;
  return apiFetch("/api/usage");
}

// ── Settings ──────────────────────────────────────────────────────────────────
export async function fetchSettings() {
  if (!IS_DESKTOP) return mock.mockSettings;
  return apiFetch("/api/settings");
}

export async function saveSettings(settings: Record<string, string>) {
  if (!IS_DESKTOP) return { saved: Object.keys(settings) };
  return apiFetch("/api/settings", { method: "POST", body: JSON.stringify(settings) });
}

export async function testConnection(service: "xpm" | "fusesign" | "teams") {
  if (!IS_DESKTOP) return { connected: false, service, error: "Demo mode — no real connection" };
  return apiFetch(`/api/settings/test/${service}`, { method: "POST" });
}

// ── Xero OAuth ────────────────────────────────────────────────────────────────
export async function fetchXeroStatus() {
  if (!IS_DESKTOP) return { configured: true, authorised: false, tenant_id: null, staff_name: null };
  return apiFetch("/api/xero/status");
}

export async function startXeroAuth() {
  if (!IS_DESKTOP) return { started: false, error: "Demo mode" };
  return apiFetch("/api/xero/start-auth", { method: "POST" });
}

export async function disconnectXero() {
  if (!IS_DESKTOP) return { disconnected: true };
  return apiFetch("/api/xero/disconnect", { method: "POST" });
}

// ── Email Rules ───────────────────────────────────────────────────────────────
export async function fetchRules() {
  if (!IS_DESKTOP) return mock.mockRules ?? [];
  return apiFetch("/api/rules");
}

export async function saveRule(rule: Record<string, any>) {
  if (!IS_DESKTOP) return [];
  return apiFetch("/api/rules", { method: "POST", body: JSON.stringify(rule) });
}

export async function deleteRule(ruleId: number) {
  if (!IS_DESKTOP) return {};
  return apiFetch(`/api/rules/${ruleId}`, { method: "DELETE" });
}

// ── Staff & Notify ────────────────────────────────────────────────────────────
export async function fetchStaff() {
  if (!IS_DESKTOP) return mock.mockStaff ?? [];
  return apiFetch("/api/staff");
}

export async function saveStaff(staff: Record<string, any>) {
  if (!IS_DESKTOP) return [];
  return apiFetch("/api/staff", { method: "POST", body: JSON.stringify(staff) });
}

export async function deleteStaff(staffId: number) {
  if (!IS_DESKTOP) return {};
  return apiFetch(`/api/staff/${staffId}`, { method: "DELETE" });
}

// ── Links & Forms ─────────────────────────────────────────────────────────────
export async function fetchLinks() {
  if (!IS_DESKTOP) return mock.mockLinks ?? [];
  return apiFetch("/api/links");
}

export async function saveLink(link: Record<string, any>) {
  if (!IS_DESKTOP) return [];
  return apiFetch("/api/links", { method: "POST", body: JSON.stringify(link) });
}

export async function deleteLink(linkId: number) {
  if (!IS_DESKTOP) return {};
  return apiFetch(`/api/links/${linkId}`, { method: "DELETE" });
}

// ── Chat ──────────────────────────────────────────────────────────────────────
export async function sendChatMessage(messages: { role: string; content: string }[]) {
  if (!IS_DESKTOP) {
    await new Promise(r => setTimeout(r, 1800));
    return {
      content: getMockChatResponse(messages[messages.length - 1]?.content || ""),
      model: "claude-sonnet-4-6",
      tier: 2,
      usage: { input_tokens: 450, output_tokens: 380 },
    };
  }
  return apiFetch("/api/chat", { method: "POST", body: JSON.stringify({ messages }) });
}

export async function fetchChatHistory() {
  if (!IS_DESKTOP) return [];
  return apiFetch("/api/chat/history");
}

export async function clearChatHistory() {
  if (!IS_DESKTOP) return {};
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
  if (!IS_DESKTOP) return [];
  return apiFetch<KnowledgeEntry[]>("/api/knowledge");
}

export async function createKnowledge(entry: Omit<KnowledgeEntry, "id" | "created_at" | "updated_at">) {
  if (!IS_DESKTOP) return { id: 0 };
  return apiFetch("/api/knowledge", { method: "POST", body: JSON.stringify(entry) });
}

export async function updateKnowledge(id: number, entry: Partial<KnowledgeEntry>) {
  if (!IS_DESKTOP) return { id, updated: true };
  return apiFetch(`/api/knowledge/${id}`, { method: "PUT", body: JSON.stringify(entry) });
}

export async function deleteKnowledge(id: number) {
  if (!IS_DESKTOP) return { id, deleted: true };
  return apiFetch(`/api/knowledge/${id}`, { method: "DELETE" });
}

// ── Lessons ───────────────────────────────────────────────────────────────────
export async function fetchLessons() {
  if (!IS_DESKTOP) return [];
  return apiFetch("/api/lessons");
}

export async function addLesson(lesson: string, source = "") {
  if (!IS_DESKTOP) return [];
  return apiFetch("/api/lessons", { method: "POST", body: JSON.stringify({ lesson, source }) });
}

export async function deleteLesson(lessonId: number) {
  if (!IS_DESKTOP) return {};
  return apiFetch(`/api/lessons/${lessonId}`, { method: "DELETE" });
}

// ── Mock chat responses ───────────────────────────────────────────────────────
function getMockChatResponse(userMsg: string): string {
  const lower = userMsg.toLowerCase();
  if (lower.includes("wip") || lower.includes("report")) {
    return `I'll build that WIP Report plugin for you using XPM and Teams.\n\n**Tier 2 Plugin: WIP Ageing Report**\n\n\`\`\`python\nclass WIPReportPlugin(AgentPlugin):\n    PLUGIN_ID = "plugin_wip_report"\n    NAME = "WIP Ageing Report"\n    SCHEDULE = Schedule.weekly_on(0, 8)  # Monday 8am\n\n    def run(self, context: PluginContext) -> PluginResult:\n        jobs = context.gateway.xpm.list_jobs(status="active")\n        stale = [j for j in jobs if j.get("days_since_update", 0) > 14]\n        summary = context.claude_reason.messages.create(\n            model=self.get_claude_model_reasoning(),\n            messages=[{"role": "user", "content": f"Summarise stale WIP: {stale}"}]\n        )\n        context.gateway.teams.send_alert(title="Weekly WIP Report", body=summary.content[0].text)\n        return PluginResult(success=True, message=f"{len(stale)} stale jobs reported")\n\`\`\`\n\n✅ Plugin saved · Runs every Monday at 8:00am · Uses Claude Sonnet`;
  }
  if (lower.includes("engagement") || lower.includes("letter")) {
    return `**Tier 2 Plugin: Engagement Letter Generator**\n\nWatches for new client emails and auto-drafts engagement letters via FuseSign.\n\n\`\`\`python\nclass EngagementLetterPlugin(AgentPlugin):\n    PLUGIN_ID = "plugin_engagement_letter"\n    SCHEDULE = Schedule.every_minutes(15)\n\n    def run(self, context: PluginContext) -> PluginResult:\n        emails = context.graph.get_unread_emails(folder="Inbox", limit=10)\n        new_clients = [e for e in emails if "new client" in e["subject"].lower()]\n        for email in new_clients:\n            client = context.gateway.xpm.get_client_by_email(email["from"])\n            draft = context.claude_reason.messages.create(\n                model=self.get_claude_model_reasoning(),\n                messages=[{"role": "user", "content": f"Draft engagement letter for {client}"}]\n            )\n            context.gateway.fusesign.create_envelope(\n                subject=f"Engagement Letter — {client['name']}",\n                body=draft.content[0].text,\n                recipients=[{"email": email["from"], "name": client["name"]}]\n            )\n        return PluginResult(success=True, message=f"Processed {len(new_clients)} new clients")\n\`\`\`\n\n✅ Plugin created and ready to activate`;
  }
  return `I can help you build that automation. Could you give me a bit more detail about:\n\n1. **Trigger** — What should start this automation? (email received, time of day, XPM job update)\n2. **Action** — What should it do? (draft email, update XPM, send Teams message, create FuseSign envelope)\n3. **Output** — Where should results go? (Teams channel, email draft, activity log)\n\nOnce I have those details I'll write the full plugin code for you.`;
}
