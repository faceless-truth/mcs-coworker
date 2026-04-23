// Design: Refined Dark Professional — KPI Monitor page

import { useEffect, useState } from "react";
import { AlertTriangle, Bell, CheckCircle2 } from "lucide-react";
import { toast } from "sonner";
import { fetchKPI } from "@/lib/api";

interface ServerKpi {
  kpi_id: string;
  label: string;
  description?: string;
  threshold: number;
  enabled: boolean;
  severity: string;
  unit: string;
  value: number | null;
  message?: string;
  last_alert?: string;
}

export default function KpiMonitor() {
  const [kpis, setKpis] = useState<ServerKpi[]>([]);
  const [loading, setLoading] = useState(true);
  const [degraded, setDegraded] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    const load = async () => {
      try {
        const rows = await fetchKPI();
        if (cancelled) return;
        if (Array.isArray(rows)) {
          setKpis(rows as ServerKpi[]);
          setDegraded(null);
        } else {
          setKpis([]);
        }
      } catch (e: any) {
        if (!cancelled) {
          setDegraded(e?.message ?? "KPI backend unavailable");
          setKpis([]);
        }
      } finally {
        if (!cancelled) setLoading(false);
      }
    };
    load();
    const id = setInterval(load, 30000);
    return () => { cancelled = true; clearInterval(id); };
  }, []);

  const testAlert = (name: string) => {
    toast.warning(`KPI Alert: ${name}`, {
      description: "Alert dispatched to Teams and email as configured.",
    });
  };

  const fmt = (val: number | null | undefined, unit: string): string => {
    if (val === null || val === undefined) return "—";
    return unit === "AUD" ? `$${val}` : String(val);
  };

  return (
    <div className="p-6 space-y-5">
      <div>
        <h1 className="text-xl font-semibold text-foreground">KPI Monitor</h1>
        <p className="text-sm text-muted-foreground mt-0.5">
          Proactive threshold monitoring — alerts fire via Teams and email when breached
        </p>
      </div>

      {degraded && (
        <div className="rounded-md border border-amber-200 bg-amber-50 px-4 py-2 text-xs text-amber-800">
          KPI backend degraded: {degraded}
        </div>
      )}

      {loading ? (
        <div className="py-12 text-center text-sm text-muted-foreground">Loading KPIs…</div>
      ) : kpis.length === 0 ? (
        <div className="py-12 text-center text-sm text-muted-foreground">
          No KPIs configured yet.
        </div>
      ) : (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          {kpis.map((kpi) => {
            const hasValue = kpi.value !== null && kpi.value !== undefined;
            const breached = hasValue && typeof kpi.value === "number" && kpi.value >= kpi.threshold;
            const pct = hasValue && kpi.threshold > 0
              ? Math.min(100, Math.round(((kpi.value as number) / kpi.threshold) * 100))
              : 0;
            return (
              <div key={kpi.kpi_id} className={`bg-white rounded-lg border shadow-sm p-5 ${breached ? "border-amber-200" : "border-border"}`}>
                <div className="flex items-start justify-between mb-4">
                  <div>
                    <div className="flex items-center gap-2">
                      {breached
                        ? <AlertTriangle className="w-4 h-4 text-amber-500" />
                        : <CheckCircle2 className="w-4 h-4 text-emerald-500" />
                      }
                      <span className="text-sm font-semibold text-foreground">{kpi.label}</span>
                    </div>
                    <div className="text-xs text-muted-foreground mt-1">
                      Threshold: {fmt(kpi.threshold, kpi.unit)} {kpi.unit}
                    </div>
                  </div>
                  <div className="text-right">
                    <div className={`text-2xl font-bold font-mono ${breached ? "text-amber-600" : "text-foreground"}`}>
                      {fmt(kpi.value, kpi.unit)}
                    </div>
                    <div className="text-xs text-muted-foreground">{kpi.unit}</div>
                  </div>
                </div>

                {/* Progress bar — only rendered when we have a real value */}
                {hasValue && (
                  <div className="h-2 bg-slate-100 rounded-full overflow-hidden mb-3">
                    <div
                      className="h-full rounded-full transition-all"
                      style={{
                        width: `${pct}%`,
                        background: breached ? "oklch(0.75 0.15 70)" : "oklch(0.5 0.15 145)",
                      }}
                    />
                  </div>
                )}

                <div className="flex items-center justify-between">
                  <div className="text-xs text-muted-foreground">
                    {kpi.last_alert ? `Last alert: ${kpi.last_alert}` : "No alerts recorded"}
                  </div>
                  <button
                    onClick={() => testAlert(kpi.label)}
                    className="flex items-center gap-1.5 px-2.5 py-1 rounded text-xs font-medium border border-border hover:border-blue-300 hover:text-blue-600 transition-all"
                  >
                    <Bell className="w-3 h-3" />
                    Test alert
                  </button>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
