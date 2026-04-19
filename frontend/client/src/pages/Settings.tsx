// Design: Refined Dark Professional — Settings page
// Configure AI models, integrations, business hours, and behaviour

import { mockSettings } from "@/lib/mockData";
import { useEffect, useState } from "react";
import { CheckCircle2, Eye, EyeOff, Save, TestTube, ExternalLink, Unlink, Loader2 } from "lucide-react";
import { toast } from "sonner";
import { fetchXeroStatus, startXeroAuth, disconnectXero, testConnection } from "@/lib/api";

interface XeroStatus {
  configured: boolean;
  authorised: boolean;
  tenant_id: string | null;
  staff_name: string | null;
}

export default function Settings() {
  const [settings, setSettings] = useState(mockSettings);
  const [showKey, setShowKey] = useState(false);
  const [saved, setSaved] = useState(false);
  const [xero, setXero] = useState<XeroStatus>({ configured: false, authorised: false, tenant_id: null, staff_name: null });
  const [xeroLoading, setXeroLoading] = useState(false);

  useEffect(() => {
    fetchXeroStatus().then(s => setXero(s as XeroStatus)).catch(() => {});
  }, []);

  const save = () => {
    setSaved(true);
    toast.success("Settings saved", { description: "Changes will take effect on the next plugin run." });
    setTimeout(() => setSaved(false), 2000);
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
          className="flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium text-white transition-all hover:opacity-90"
          style={{ background: "oklch(0.5 0.2 250)" }}
        >
          {saved ? <CheckCircle2 className="w-4 h-4" /> : <Save className="w-4 h-4" />}
          {saved ? "Saved!" : "Save Changes"}
        </button>
      </div>

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
