// Design: Refined Dark Professional — KPI Monitor page

import { mockKpis } from "@/lib/mockData";
import { AlertTriangle, Bell, CheckCircle2, TrendingDown, TrendingUp } from "lucide-react";
import { toast } from "sonner";

export default function KpiMonitor() {
  const testAlert = (name: string) => {
    toast.warning(`KPI Alert: ${name}`, {
      description: "Alert dispatched to Teams and email as configured.",
    });
  };

  return (
    <div className="p-6 space-y-5">
      <div>
        <h1 className="text-xl font-semibold text-foreground">KPI Monitor</h1>
        <p className="text-sm text-muted-foreground mt-0.5">
          Proactive threshold monitoring — alerts fire via Teams and email when breached
        </p>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        {mockKpis.map((kpi) => {
          const pct = Math.min(100, Math.round((kpi.value / kpi.threshold) * 100));
          const isWarning = kpi.status === "warning";
          return (
            <div key={kpi.name} className={`bg-white rounded-lg border shadow-sm p-5 ${isWarning ? "border-amber-200" : "border-border"}`}>
              <div className="flex items-start justify-between mb-4">
                <div>
                  <div className="flex items-center gap-2">
                    {isWarning
                      ? <AlertTriangle className="w-4 h-4 text-amber-500" />
                      : <CheckCircle2 className="w-4 h-4 text-emerald-500" />
                    }
                    <span className="text-sm font-semibold text-foreground">{kpi.name}</span>
                  </div>
                  <div className="text-xs text-muted-foreground mt-1">
                    Threshold: {typeof kpi.threshold === "number" && kpi.unit === "AUD" ? `$${kpi.threshold}` : kpi.threshold} {kpi.unit}
                  </div>
                </div>
                <div className="text-right">
                  <div className={`text-2xl font-bold font-mono ${isWarning ? "text-amber-600" : "text-foreground"}`}>
                    {typeof kpi.value === "number" && kpi.unit === "AUD" ? `$${kpi.value}` : kpi.value}
                  </div>
                  <div className="text-xs text-muted-foreground">{kpi.unit}</div>
                </div>
              </div>

              {/* Progress bar */}
              <div className="h-2 bg-slate-100 rounded-full overflow-hidden mb-3">
                <div
                  className="h-full rounded-full transition-all"
                  style={{
                    width: `${pct}%`,
                    background: isWarning ? "oklch(0.75 0.15 70)" : "oklch(0.5 0.15 145)",
                  }}
                />
              </div>

              <div className="flex items-center justify-between">
                <div className="flex items-center gap-1 text-xs text-muted-foreground">
                  {kpi.trend.startsWith("+") ? (
                    <TrendingUp className="w-3 h-3 text-rose-400" />
                  ) : kpi.trend.startsWith("-") ? (
                    <TrendingDown className="w-3 h-3 text-emerald-500" />
                  ) : null}
                  {kpi.trend} vs yesterday
                </div>
                <button
                  onClick={() => testAlert(kpi.name)}
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
    </div>
  );
}
