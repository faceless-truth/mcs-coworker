// Design: Refined Dark Professional — Settings page
// Configure AI models, integrations, business hours, and behaviour

import { useEffect, useState } from "react";
import { CheckCircle2, Eye, EyeOff, Save, TestTube, ExternalLink, Unlink, Loader2, Plus, Trash2, Pencil, X } from "lucide-react";
import { toast } from "sonner";
import {
  fetchXeroStatus, startXeroAuth, disconnectXero, testConnection,
  fetchSettings, saveSettings,
  fetchKnowledge, createKnowledge, updateKnowledge, deleteKnowledge,
  type KnowledgeEntry,
} from "@/lib/api";

const KB_CATEGORIES = ["Pricing", "Checklists", "Procedures", "Firm Policies", "Staff Info", "Other"];

function ModeToggle() {
  const [mode, setMode] = useState<"reception" | "accountant" | null>(null);
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    (async () => {
      try {
        const s: any = await fetchSettings();
        setMode(s.reception_mode === "1" ? "reception" : "accountant");
      } catch {
        setMode("accountant");
      }
    })();
  }, []);

  const change = async (next: "reception" | "accountant") => {
    if (next === mode || saving) return;
    setSaving(true);
    const prev = mode;
    setMode(next);
    try {
      await saveSettings({ reception_mode: next === "reception" ? "1" : "0" });
      toast.success(`Mode set to ${next === "reception" ? "Reception" : "Accountant"}`, {
        description: "The Plugins page will refresh to show only the relevant plugins.",
      });
      window.dispatchEvent(new CustomEvent("reception-mode-changed", { detail: next }));
    } catch (e: any) {
      setMode(prev);
      toast.error("Failed to change mode", { description: e.message });
    } finally {
      setSaving(false);
    }
  };

  if (mode === null) {
    return (
      <div className="bg-white rounded-lg border border-border shadow-sm p-5 flex items-center gap-2 text-sm text-muted-foreground">
        <Loader2 className="w-4 h-4 animate-spin" /> Loading mode...
      </div>
    );
  }

  const pillBase = "flex-1 px-4 py-2 text-sm font-medium rounded-md transition-all";
  const active = "text-white shadow-sm";
  const inactive = "bg-transparent text-muted-foreground hover:text-foreground";

  return (
    <div className="bg-white rounded-lg border border-border shadow-sm overflow-hidden">
      <div className="px-5 py-3.5 border-b border-border bg-slate-50">
        <div className="text-sm font-semibold text-foreground">Mode</div>
        <div className="text-xs text-muted-foreground mt-0.5">
          Reception mode shows plugins for the front desk (ASIC, NOA, BAS, etc). Accountant mode shows plugins for individual accountants (meeting prep, client outreach, etc). Universal plugins like Morning Briefing run in both modes.
        </div>
      </div>
      <div className="p-5">
        <div className="inline-flex items-center p-1 rounded-lg bg-slate-100 w-full max-w-md">
          <button
            onClick={() => change("reception")}
            disabled={saving}
            className={`${pillBase} ${mode === "reception" ? active : inactive}`}
            style={mode === "reception" ? { background: "oklch(0.5 0.2 250)" } : undefined}
            data-testid="mode-toggle-reception"
          >
            Reception
          </button>
          <button
            onClick={() => change("accountant")}
            disabled={saving}
            className={`${pillBase} ${mode === "accountant" ? active : inactive}`}
            style={mode === "accountant" ? { background: "oklch(0.5 0.2 250)" } : undefined}
            data-testid="mode-toggle-accountant"
          >
            Accountant
          </button>
        </div>
      </div>
    </div>
  );
}

function KnowledgeBaseSection() {
  const [entries, setEntries] = useState<KnowledgeEntry[]>([]);
  const [loading, setLoading] = useState(true);
  const [editing, setEditing] = useState<KnowledgeEntry | null>(null);
  const [showForm, setShowForm] = useState(false);

  const load = async () => {
    try {
      const rows = await fetchKnowledge();
      setEntries(rows || []);
    } catch (e: any) {
      toast.error("Failed to load knowledge base", { description: e.message });
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); }, []);

  const startAdd = () => {
    setEditing({ id: 0, category: KB_CATEGORIES[0], title: "", content: "", enabled: 1 });
    setShowForm(true);
  };

  const startEdit = (entry: KnowledgeEntry) => {
    setEditing({ ...entry });
    setShowForm(true);
  };

  const save = async () => {
    if (!editing) return;
    const { id, category, title, content, enabled } = editing;
    if (!category.trim() || !title.trim()) {
      toast.error("Category and title are required");
      return;
    }
    try {
      if (id === 0) {
        await createKnowledge({ category, title, content, enabled });
        toast.success("Knowledge entry added");
      } else {
        await updateKnowledge(id, { category, title, content, enabled });
        toast.success("Knowledge entry updated");
      }
      setShowForm(false);
      setEditing(null);
      load();
    } catch (e: any) {
      toast.error("Save failed", { description: e.message });
    }
  };

  const remove = async (entry: KnowledgeEntry) => {
    if (!confirm(`Delete "${entry.title}"?`)) return;
    try {
      await deleteKnowledge(entry.id);
      toast.success("Deleted");
      load();
    } catch (e: any) {
      toast.error("Delete failed", { description: e.message });
    }
  };

  const grouped = entries.reduce<Record<string, KnowledgeEntry[]>>((acc, e) => {
    (acc[e.category] ||= []).push(e);
    return acc;
  }, {});
  const categories = Object.keys(grouped).sort();

  return (
    <div className="bg-white rounded-lg border border-border shadow-sm overflow-hidden">
      <div className="px-5 py-3.5 border-b border-border bg-slate-50 flex items-center justify-between">
        <div>
          <div className="text-sm font-semibold text-foreground">Knowledge Base</div>
          <div className="text-xs text-muted-foreground mt-0.5">
            The Smart Email Responder uses these entries to answer client questions about pricing, checklists, and procedures.
          </div>
        </div>
        <button
          onClick={startAdd}
          className="flex items-center gap-1.5 px-3 py-2 rounded-md text-xs font-medium text-white hover:opacity-90"
          style={{ background: "oklch(0.5 0.2 250)" }}
        >
          <Plus className="w-3.5 h-3.5" /> Add Knowledge
        </button>
      </div>
      <div className="p-5 space-y-4">
        {loading ? (
          <div className="text-sm text-muted-foreground">Loading...</div>
        ) : categories.length === 0 ? (
          <div className="text-sm text-muted-foreground">No entries yet. Click "Add Knowledge" to create one.</div>
        ) : (
          categories.map(cat => (
            <div key={cat}>
              <div className="text-xs font-semibold uppercase tracking-wide text-muted-foreground mb-2">{cat}</div>
              <div className="space-y-2">
                {grouped[cat].map(e => (
                  <div key={e.id} className="flex items-start gap-3 p-3 rounded-md border border-border bg-slate-50">
                    <div className="flex-1 min-w-0">
                      <div className="text-sm font-medium text-foreground">{e.title}</div>
                      <div className="text-xs text-muted-foreground mt-1 line-clamp-2 whitespace-pre-wrap">
                        {e.content || <span className="italic">(empty — click edit to fill in)</span>}
                      </div>
                    </div>
                    <button onClick={() => startEdit(e)} className="p-1.5 rounded hover:bg-slate-200" title="Edit">
                      <Pencil className="w-3.5 h-3.5 text-muted-foreground" />
                    </button>
                    <button onClick={() => remove(e)} className="p-1.5 rounded hover:bg-rose-50" title="Delete">
                      <Trash2 className="w-3.5 h-3.5 text-rose-500" />
                    </button>
                  </div>
                ))}
              </div>
            </div>
          ))
        )}
      </div>

      {showForm && editing && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4"
          style={{ background: "rgba(0,0,0,0.45)" }}
          onClick={e => { if (e.target === e.currentTarget) { setShowForm(false); setEditing(null); } }}>
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-lg overflow-hidden">
            <div className="flex items-center justify-between px-5 py-3.5 border-b border-border bg-slate-50">
              <div className="text-sm font-semibold text-foreground">
                {editing.id === 0 ? "Add Knowledge" : "Edit Knowledge"}
              </div>
              <button onClick={() => { setShowForm(false); setEditing(null); }}
                className="p-1 rounded hover:bg-slate-200">
                <X className="w-4 h-4" />
              </button>
            </div>
            <div className="p-5 space-y-4">
              <div>
                <label className="block text-xs font-medium text-foreground mb-1">Category</label>
                <select
                  value={editing.category}
                  onChange={e => setEditing({ ...editing, category: e.target.value })}
                  className="w-full px-3 py-2 text-sm border border-border rounded-md bg-white"
                >
                  {KB_CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>
              <div>
                <label className="block text-xs font-medium text-foreground mb-1">Title</label>
                <input
                  value={editing.title}
                  onChange={e => setEditing({ ...editing, title: e.target.value })}
                  className="w-full px-3 py-2 text-sm border border-border rounded-md bg-white"
                  placeholder="e.g. Standard Fees"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-foreground mb-1">Content</label>
                <textarea
                  value={editing.content}
                  onChange={e => setEditing({ ...editing, content: e.target.value })}
                  rows={8}
                  className="w-full px-3 py-2 text-sm border border-border rounded-md bg-white font-mono"
                  placeholder="Multi-line content that Claude will use as context when replying to emails..."
                />
              </div>
            </div>
            <div className="px-5 py-3 border-t border-border flex items-center justify-end gap-2 bg-slate-50">
              <button onClick={() => { setShowForm(false); setEditing(null); }}
                className="px-4 py-2 text-xs font-medium rounded-md border border-border hover:bg-white">
                Cancel
              </button>
              <button onClick={save}
                className="px-4 py-2 text-xs font-medium rounded-md text-white hover:opacity-90"
                style={{ background: "oklch(0.5 0.2 250)" }}>
                Save
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

interface XeroStatus {
  configured: boolean;
  authorised: boolean;
  tenant_id: string | null;
  staff_name: string | null;
}

// Empty defaults — real values flow in from /api/settings on mount. Sensitive
// fields come back masked from the backend; only non-masked edits are saved.
const EMPTY_SETTINGS = {
  anthropicApiKey: "",
  fastModel: "",
  reasoningModel: "",
  outlookEmail: "",
  xpmApiKey: "",
  xeroClientId: "",
  xeroClientSecret: "",
  fuseSignApiKey: "",
  teamsWebhook: "",
  confidenceThreshold: 0.75,
  heartbeatInterval: 300,
  autoUpdate: false,
  draftMode: false,
};

export default function Settings() {
  const [settings, setSettings] = useState(EMPTY_SETTINGS);
  const [showKey, setShowKey] = useState(false);
  const [saved, setSaved] = useState(false);
  const [saving, setSaving] = useState(false);
  const [xero, setXero] = useState<XeroStatus>({ configured: false, authorised: false, tenant_id: null, staff_name: null });
  const [xeroLoading, setXeroLoading] = useState(false);

  // Load real settings once on mount.
  useEffect(() => {
    (async () => {
      try {
        const s: any = await fetchSettings();
        if (s && typeof s === "object") {
          setSettings(prev => ({
            ...prev,
            anthropicApiKey:     s.anthropic_api_key     ?? prev.anthropicApiKey,
            fastModel:           s.fast_model            ?? prev.fastModel,
            reasoningModel:      s.reasoning_model       ?? prev.reasoningModel,
            outlookEmail:        s.outlook_email         ?? prev.outlookEmail,
            xeroClientId:        s.xero_client_id        ?? prev.xeroClientId,
            xeroClientSecret:    s.xero_client_secret    ?? prev.xeroClientSecret,
            fuseSignApiKey:      s.fusesign_api_key      ?? prev.fuseSignApiKey,
            teamsWebhook:        s.teams_webhook_url     ?? prev.teamsWebhook,
            confidenceThreshold: s.confidence_threshold !== undefined ? parseFloat(s.confidence_threshold) : prev.confidenceThreshold,
            heartbeatInterval:   s.heartbeat_interval_seconds !== undefined ? parseInt(s.heartbeat_interval_seconds, 10) : prev.heartbeatInterval,
            autoUpdate:          s.auto_update_enabled === "1" || s.auto_update_enabled === true,
            draftMode:           s.draft_mode === "1" || s.draft_mode === true,
          }));
        }
      } catch {
        // Backend not ready — keep empty defaults.
      }
    })();
    fetchXeroStatus().then(s => setXero(s as XeroStatus)).catch(() => {});
  }, []);

  const save = async () => {
    setSaving(true);
    // Map camelCase UI keys to the snake_case keys the /api/settings POST
    // handler whitelists. Any field still containing bullets ("••••") is a
    // masked placeholder — the backend's own handler filters those out too.
    const payload: Record<string, string> = {
      anthropic_api_key:           settings.anthropicApiKey,
      outlook_email:               settings.outlookEmail,
      fusesign_api_key:            settings.fuseSignApiKey,
      teams_webhook_url:           settings.teamsWebhook,
      fast_model:                  settings.fastModel,
      reasoning_model:             settings.reasoningModel,
      confidence_threshold:        String(settings.confidenceThreshold),
      heartbeat_interval_seconds:  String(settings.heartbeatInterval),
      draft_mode:                  settings.draftMode ? "1" : "0",
      auto_update_enabled:         settings.autoUpdate ? "1" : "0",
      xero_client_id:              settings.xeroClientId,
      xero_client_secret:          settings.xeroClientSecret,
    };
    try {
      await saveSettings(payload);
      setSaved(true);
      toast.success("Settings saved", { description: "Changes will take effect on the next plugin run." });
      setTimeout(() => setSaved(false), 2000);
    } catch (e: any) {
      toast.error("Save failed", { description: e?.message });
    } finally {
      setSaving(false);
    }
  };

  const handleConnectXero = async () => {
    setXeroLoading(true);
    try {
      await startXeroAuth();
      toast.info("Xero login opened in browser", { description: "Complete the login, then return here." });
      // Poll for auth completion
      let attempts = 0;
      const poll = setInterval(async () => {
        attempts++;
        const status = await fetchXeroStatus() as XeroStatus;
        if (status.authorised) {
          setXero(status);
          clearInterval(poll);
          setXeroLoading(false);
          toast.success(`Connected to Xero${status.staff_name ? ` as ${status.staff_name}` : ""}`);
        }
        if (attempts > 60) { clearInterval(poll); setXeroLoading(false); }
      }, 3000);
    } catch (e: any) {
      toast.error(e.message || "Failed to start Xero auth");
      setXeroLoading(false);
    }
  };

  const handleDisconnectXero = async () => {
    try {
      await disconnectXero();
      setXero({ configured: true, authorised: false, tenant_id: null, staff_name: null });
      toast.success("Disconnected from Xero");
    } catch {
      toast.error("Failed to disconnect");
    }
  };

  const handleTestConnection = async (service: "fusesign" | "teams") => {
    try {
      const result = await testConnection(service) as any;
      if (result.connected) {
        toast.success(`${service === "fusesign" ? "FuseSign" : "Teams"} connection successful`);
      } else {
        toast.error(`${service} connection failed`, { description: result.error });
      }
    } catch (e: any) {
      toast.error(e.message || "Connection test failed");
    }
  };

  const Section = ({ title, description, children }: { title: string; description?: string; children: React.ReactNode }) => (
    <div className="bg-white rounded-lg border border-border shadow-sm overflow-hidden">
      <div className="px-5 py-3.5 border-b border-border bg-slate-50">
        <div className="text-sm font-semibold text-foreground">{title}</div>
        {description && <div className="text-xs text-muted-foreground mt-0.5">{description}</div>}
      </div>
      <div className="p-5 space-y-4">{children}</div>
    </div>
  );

  const Field = ({ label, hint, children }: { label: string; hint?: string; children: React.ReactNode }) => (
    <div>
      <label className="block text-xs font-medium text-foreground mb-1">{label}</label>
      {hint && <div className="text-xs text-muted-foreground mb-1.5">{hint}</div>}
      {children}
    </div>
  );

  const inputClass = "w-full px-3 py-2 text-sm border border-border rounded-md bg-white focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-400 font-mono";

  return (
    <div className="p-6 space-y-5">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-xl font-semibold text-foreground">Settings</h1>
          <p className="text-sm text-muted-foreground mt-0.5">Configure CoWorker's AI models, integrations, and behaviour</p>
        </div>
        <button
          onClick={save}
          disabled={saving}
          className="flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium text-white transition-all hover:opacity-90 disabled:opacity-60"
          style={{ background: "oklch(0.5 0.2 250)" }}
        >
          {saved ? <CheckCircle2 className="w-4 h-4" /> : <Save className="w-4 h-4" />}
          {saving ? "Saving…" : saved ? "Saved!" : "Save Changes"}
        </button>
      </div>

      {/* Mode toggle — gates which plugins show up */}
      <ModeToggle />

      {/* Knowledge base — used by the Smart Email Responder */}
      <KnowledgeBaseSection />

      {/* Claude AI */}
      <Section title="Claude AI — Dual Model" description="Configure the AI models used for fast triage and deep reasoning">
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          <Field label="Fast Model (Haiku)" hint="Used for triage, classification, and drafting">
            <input className={inputClass} value={settings.fastModel} onChange={e => setSettings(s => ({ ...s, fastModel: e.target.value }))} />
          </Field>
          <Field label="Reasoning Model (Sonnet)" hint="Used for complex analysis and decisions">
            <input className={inputClass} value={settings.reasoningModel} onChange={e => setSettings(s => ({ ...s, reasoningModel: e.target.value }))} />
          </Field>
        </div>
        <Field label="Anthropic API Key">
          <div className="relative">
            <input
              type={showKey ? "text" : "password"}
              className={inputClass + " pr-10"}
              value={settings.anthropicApiKey}
              onChange={e => setSettings(s => ({ ...s, anthropicApiKey: e.target.value }))}
            />
            <button onClick={() => setShowKey(!showKey)} className="absolute right-3 top-1/2 -translate-y-1/2 text-muted-foreground hover:text-foreground">
              {showKey ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
            </button>
          </div>
        </Field>
      </Section>

      {/* Microsoft 365 */}
      <Section title="Microsoft 365" description="The Outlook mailbox CoWorker monitors for incoming emails">
        <Field label="Outlook Mailbox">
          <input className={inputClass} value={settings.outlookEmail} onChange={e => setSettings(s => ({ ...s, outlookEmail: e.target.value }))} />
        </Field>
      </Section>

      {/* Xero XPM — OAuth */}
      <Section title="XPM / Xero Practice Manager" description="Connect via OAuth 2.0 to enable WIP summaries, job lookups, client notes, and meeting prep">
        <Field label="Xero Client ID" hint="From your Xero app in developer.xero.com (leave blank to use XERO_CLIENT_ID env var)">
          <input
            type="text"
            className={inputClass}
            value={settings.xeroClientId}
            placeholder="Not configured"
            onChange={e => setSettings(s => ({ ...s, xeroClientId: e.target.value }))}
          />
        </Field>
        <Field label="Xero Client Secret" hint="Stored per-install. Rotate in the Xero developer portal after setting.">
          <input
            type="password"
            className={inputClass}
            value={settings.xeroClientSecret}
            placeholder="Not configured"
            onChange={e => setSettings(s => ({ ...s, xeroClientSecret: e.target.value }))}
          />
        </Field>
        <div className="flex items-center justify-between p-4 rounded-lg border border-border bg-slate-50">
          <div className="flex items-center gap-3">
            {/* Xero logo placeholder */}
            <div className="w-9 h-9 rounded-lg flex items-center justify-center text-white text-sm font-bold flex-shrink-0"
              style={{ background: "oklch(0.55 0.18 155)" }}>
              X
            </div>
            <div>
              <div className="text-sm font-medium text-foreground">
                {xero.authorised
                  ? `Connected${xero.staff_name ? ` · ${xero.staff_name}` : ""}`
                  : xero.configured ? "Not connected" : "Not configured"}
              </div>
              <div className="text-xs text-muted-foreground mt-0.5">
                {xero.authorised
                  ? `Tenant: ${xero.tenant_id?.slice(0, 8) ?? "—"}… · Clients scoped to your account`
                  : "Click Connect to authorise CoWorker with your Xero account"}
              </div>
            </div>
          </div>
          <div className="flex items-center gap-2">
            {xero.authorised ? (
              <button
                onClick={handleDisconnectXero}
                className="flex items-center gap-1.5 px-3 py-2 rounded-md text-xs font-medium border border-rose-200 text-rose-600 hover:bg-rose-50 transition-colors"
              >
                <Unlink className="w-3.5 h-3.5" />
                Disconnect
              </button>
            ) : (
              <button
                onClick={handleConnectXero}
                disabled={xeroLoading || !xero.configured}
                className="flex items-center gap-1.5 px-4 py-2 rounded-md text-xs font-medium text-white transition-all hover:opacity-90 disabled:opacity-50"
                style={{ background: "oklch(0.55 0.18 155)" }}
              >
                {xeroLoading ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <ExternalLink className="w-3.5 h-3.5" />}
                {xeroLoading ? "Waiting…" : "Connect Xero"}
              </button>
            )}
          </div>
        </div>
        {xero.authorised && (
          <div className="flex items-center gap-2 text-xs text-emerald-700 bg-emerald-50 border border-emerald-100 rounded-md px-3 py-2">
            <CheckCircle2 className="w-3.5 h-3.5 flex-shrink-0" />
            XPM connected — client data scoped to your manager account. Token auto-refreshes.
          </div>
        )}
      </Section>

      {/* FuseSign */}
      <Section title="FuseSign" description="Document signing integration for engagement letters and tax returns">
        <Field label="FuseSign API Key">
          <div className="flex gap-2">
            <input
              type="password"
              className={inputClass}
              value={settings.fuseSignApiKey}
              placeholder="Not configured"
              onChange={e => setSettings(s => ({ ...s, fuseSignApiKey: e.target.value }))}
            />
            <button
              onClick={() => handleTestConnection("fusesign")}
              className="flex items-center gap-1.5 px-3 py-2 rounded-md text-xs font-medium border border-border hover:border-blue-300 hover:text-blue-600 transition-all flex-shrink-0"
            >
              <TestTube className="w-3.5 h-3.5" />
              Test
            </button>
          </div>
        </Field>
      </Section>

      {/* Microsoft Teams */}
      <Section title="Microsoft Teams" description="Incoming webhook for automated alerts and morning briefings">
        <Field label="Teams Webhook URL">
          <div className="flex gap-2">
            <input
              type="password"
              className={inputClass}
              value={settings.teamsWebhook}
              onChange={e => setSettings(s => ({ ...s, teamsWebhook: e.target.value }))}
            />
            <button
              onClick={() => handleTestConnection("teams")}
              className="flex items-center gap-1.5 px-3 py-2 rounded-md text-xs font-medium border border-border hover:border-blue-300 hover:text-blue-600 transition-all flex-shrink-0"
            >
              <TestTube className="w-3.5 h-3.5" />
              Test
            </button>
          </div>
        </Field>
      </Section>

      {/* Autonomy */}
      <Section title="Autonomy & Behaviour" description="Control how independently CoWorker acts on your behalf">
        <Field label="Confidence Threshold" hint={`Actions below ${Math.round(settings.confidenceThreshold * 100)}% confidence require human approval`}>
          <div className="flex items-center gap-4">
            <input
              type="range" min={0.5} max={1} step={0.05}
              value={settings.confidenceThreshold}
              onChange={e => setSettings(s => ({ ...s, confidenceThreshold: parseFloat(e.target.value) }))}
              className="flex-1"
            />
            <span className="text-sm font-mono font-semibold text-foreground w-12">{Math.round(settings.confidenceThreshold * 100)}%</span>
          </div>
        </Field>
        <Field label="Heartbeat Interval" hint="How often the scheduler checks for work (seconds)">
          <div className="flex items-center gap-4">
            <input
              type="range" min={60} max={600} step={60}
              value={settings.heartbeatInterval}
              onChange={e => setSettings(s => ({ ...s, heartbeatInterval: parseInt(e.target.value) }))}
              className="flex-1"
            />
            <span className="text-sm font-mono font-semibold text-foreground w-16">{settings.heartbeatInterval}s</span>
          </div>
        </Field>
        <div className="flex items-center justify-between pt-2">
          <div>
            <div className="text-xs font-medium text-foreground">Draft Mode</div>
            <div className="text-xs text-muted-foreground">All emails drafted but not sent — requires manual approval</div>
          </div>
          <button
            onClick={() => setSettings(s => ({ ...s, draftMode: !s.draftMode }))}
            className="relative w-10 h-6 rounded-full transition-all"
            style={{ background: settings.draftMode ? "oklch(0.5 0.2 250)" : "oklch(0.85 0.005 240)" }}
          >
            <span className={`absolute top-1 w-4 h-4 bg-white rounded-full shadow transition-all ${settings.draftMode ? "left-5" : "left-1"}`} />
          </button>
        </div>
        <div className="flex items-center justify-between">
          <div>
            <div className="text-xs font-medium text-foreground">Auto-Update Plugins</div>
            <div className="text-xs text-muted-foreground">Automatically apply plugin updates from GitHub releases</div>
          </div>
          <button
            onClick={() => setSettings(s => ({ ...s, autoUpdate: !s.autoUpdate }))}
            className="relative w-10 h-6 rounded-full transition-all"
            style={{ background: settings.autoUpdate ? "oklch(0.5 0.2 250)" : "oklch(0.85 0.005 240)" }}
          >
            <span className={`absolute top-1 w-4 h-4 bg-white rounded-full shadow transition-all ${settings.autoUpdate ? "left-5" : "left-1"}`} />
          </button>
        </div>
      </Section>
    </div>
  );
}
