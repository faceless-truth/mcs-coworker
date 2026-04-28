// Design: Refined Dark Professional — Plugins page
// Shows all plugins with status, controls, and category filtering
// Data fetched live from /api/plugins — NO mock data

import { useState, useEffect } from "react";
import { CheckCircle2, Clock, Pause, Play, XCircle } from "lucide-react";
import { toast } from "sonner";

const API_BASE = "http://127.0.0.1:7842";
// Backend emits one of these three values on plugin.category — see PluginCategory
// in plugin_base.py. Keep this list in sync with that enum.
const CATEGORIES = [
  { value: "all",        label: "All" },
  { value: "universal",  label: "Universal" },
  { value: "reception",  label: "Reception" },
  { value: "accountant", label: "Accountant" },
];

export default function Plugins() {
  const [filter, setFilter] = useState("all");
  const [plugins, setPlugins] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);
  const [isPassive, setIsPassive] = useState(false);

  useEffect(() => {
    const fetchPlugins = async () => {
      try {
        const res = await fetch(`${API_BASE}/api/plugins`);
        if (res.ok) {
          const json = await res.json();
          const data = (json && typeof json === "object" && "data" in json) ? json.data : json;
          setPlugins(Array.isArray(data) ? data : []);
        }
      } catch (_) {
        // backend not ready yet
      } finally {
        setLoading(false);
      }
    };
    const fetchMode = async () => {
      try {
        const res = await fetch(`${API_BASE}/api/processing-mode`);
        if (res.ok) {
          const json = await res.json();
          setIsPassive((json?.mode || "active").toLowerCase() === "passive");
        }
      } catch (_) {
        // best effort
      }
    };
    fetchPlugins();
    fetchMode();
    const interval = setInterval(() => { fetchPlugins(); fetchMode(); }, 30000);
    const onModeChanged = () => fetchPlugins();
    const onProcModeChanged = (e: Event) => {
      const detail = (e as CustomEvent).detail as string;
      setIsPassive((detail || "active").toLowerCase() === "passive");
    };
    window.addEventListener("reception-mode-changed", onModeChanged);
    window.addEventListener("processing-mode-changed", onProcModeChanged);
    return () => {
      clearInterval(interval);
      window.removeEventListener("reception-mode-changed", onModeChanged);
      window.removeEventListener("processing-mode-changed", onProcModeChanged);
    };
  }, []);

  const filtered = filter === "all" ? plugins : plugins.filter((p: any) => p.category === filter);

  const togglePlugin = async (id: string) => {
    const plugin = plugins.find((p: any) => p.id === id);
    if (!plugin) return;
    const newEnabled = !plugin.enabled;
    // Optimistic update
    setPlugins(prev => prev.map((p: any) =>
      p.id === id ? { ...p, enabled: newEnabled, status: newEnabled ? "idle" : "disabled" } : p
    ));
    toast(newEnabled ? `${plugin.name} enabled` : `${plugin.name} disabled`, {
      description: newEnabled ? "Plugin will run on next scheduled tick." : "Plugin paused — no further runs.",
    });
    try {
      await fetch(`${API_BASE}/api/plugins/${id}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ enabled: newEnabled }),
      });
    } catch (_) {
      // best effort
    }
  };

  const runNow = async (id: string) => {
    if (isPassive) {
      toast.error("Disabled in passive mode", {
        description: "Switch to Active in Settings to run plugins on this machine.",
      });
      return;
    }
    const plugin = plugins.find((p: any) => p.id === id);
    if (!plugin) return;
    toast(`Running ${plugin.name}...`, { description: "Plugin triggered manually — check Activity Log for results." });
    try {
      await fetch(`${API_BASE}/api/plugins/${id}/run`, { method: "POST" });
    } catch (_) {
      // best effort
    }
  };

  return (
    <div className="p-6 space-y-5">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-xl font-semibold text-foreground">Plugins</h1>
          <p className="text-sm text-muted-foreground mt-0.5">
            {loading ? "Loading..." : `${plugins.filter((p: any) => p.enabled).length} of ${plugins.length} plugins active`}
          </p>
        </div>
      </div>

      {/* Category filter */}
      <div className="flex flex-wrap gap-2">
        {CATEGORIES.map(cat => (
          <button
            key={cat.value}
            onClick={() => setFilter(cat.value)}
            className={`px-3 py-1.5 rounded-full text-xs font-medium transition-all ${
              filter === cat.value
                ? "text-white"
                : "bg-white border border-border text-muted-foreground hover:border-blue-300 hover:text-blue-600"
            }`}
            style={filter === cat.value ? { background: "oklch(0.5 0.2 250)" } : {}}
          >
            {cat.label}
          </button>
        ))}
      </div>

      {/* Plugin grid */}
      {loading ? (
        <div className="py-12 text-center text-sm text-muted-foreground">Loading plugins...</div>
      ) : plugins.length === 0 ? (
        <div className="py-12 text-center text-sm text-muted-foreground">No plugins registered yet.</div>
      ) : (
        <div className="grid grid-cols-1 lg:grid-cols-2 xl:grid-cols-3 gap-4">
          {filtered.map((plugin: any) => (
            <div key={plugin.id} className={`plugin-card ${!plugin.enabled ? "opacity-60" : ""}`}>
              <div className="flex items-start justify-between mb-3">
                <div className="flex items-center gap-2">
                  <span className={`status-dot ${plugin.status ?? "idle"}`} />
                  <span className="text-sm font-semibold text-foreground">{plugin.name}</span>
                </div>
                <span className="badge neutral">{plugin.category ?? "—"}</span>
              </div>

              <p className="text-xs text-muted-foreground mb-4 leading-relaxed">{plugin.description ?? ""}</p>

              <div className="grid grid-cols-2 gap-2 mb-4 text-xs">
                <div>
                  <div className="text-muted-foreground">Last run</div>
                  <div className="font-medium text-foreground">{plugin.lastRun ?? plugin.last_run ?? "Never"}</div>
                </div>
                <div>
                  <div className="text-muted-foreground">Next run</div>
                  <div className="font-medium text-foreground">{plugin.nextRun ?? plugin.next_run ?? "—"}</div>
                </div>
                <div>
                  <div className="text-muted-foreground">Runs today</div>
                  <div className="font-medium text-foreground">{plugin.runsToday ?? plugin.runs_today ?? 0}</div>
                </div>
                <div>
                  <div className="text-muted-foreground">Success rate</div>
                  {(() => {
                    const rate = plugin.successRate ?? plugin.success_rate ?? 100;
                    return (
                      <div className={`font-medium ${rate >= 95 ? "text-emerald-600" : rate >= 80 ? "text-amber-600" : "text-rose-600"}`}>
                        {rate}%
                      </div>
                    );
                  })()}
                </div>
              </div>

              <div className="flex items-center gap-2 pt-3 border-t border-border">
                <span className={`badge ${(plugin.model ?? "").includes("sonnet") ? "info" : "neutral"}`}>
                  {plugin.model ?? "—"}
                </span>
                <div className="flex-1" />
                <button
                  onClick={() => runNow(plugin.id)}
                  disabled={!plugin.enabled || isPassive}
                  title={isPassive ? "Disabled in passive mode" : undefined}
                  className="flex items-center gap-1.5 px-2.5 py-1.5 rounded text-xs font-medium border border-border hover:border-blue-300 hover:text-blue-600 transition-all disabled:opacity-40 disabled:cursor-not-allowed"
                >
                  <Play className="w-3 h-3" />
                  Run now
                </button>
                <button
                  onClick={() => togglePlugin(plugin.id)}
                  className={`flex items-center gap-1.5 px-2.5 py-1.5 rounded text-xs font-medium transition-all ${
                    plugin.enabled
                      ? "border border-border hover:border-rose-300 hover:text-rose-600"
                      : "text-white"
                  }`}
                  style={!plugin.enabled ? { background: "oklch(0.5 0.2 250)" } : {}}
                >
                  {plugin.enabled ? <Pause className="w-3 h-3" /> : <Play className="w-3 h-3" />}
                  {plugin.enabled ? "Disable" : "Enable"}
                </button>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
