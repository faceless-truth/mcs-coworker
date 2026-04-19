// Design: Refined Dark Professional — AI Usage page
// All data fetched from /api/usage — NO mock data

import { useState, useEffect } from "react";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell } from "recharts";
import { DollarSign, Hash, Zap } from "lucide-react";

const API_BASE = "http://127.0.0.1:7842";

export default function Usage() {
  const [usage, setUsage] = useState<any>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const fetchUsage = async () => {
      try {
        const res = await fetch(`${API_BASE}/api/usage`);
        if (res.ok) {
          const data = await res.json();
          setUsage(data);
        }
      } catch (_) {}
      finally {
        setLoading(false);
      }
    };
    fetchUsage();
    const interval = setInterval(fetchUsage, 30000);
    return () => clearInterval(interval);
  }, []);

  const todayCost = usage?.todayCost ?? usage?.today_cost ?? 0;
  const monthlyCost = usage?.monthlyCost ?? usage?.monthly_cost ?? 0;
  const monthlyBudget = usage?.monthlyBudget ?? usage?.monthly_budget ?? 100;
  const totalCalls = usage?.totalCalls ?? usage?.total_calls ?? 0;
  const totalTokensIn = usage?.totalTokensIn ?? usage?.total_tokens_in ?? 0;
  const totalTokensOut = usage?.totalTokensOut ?? usage?.total_tokens_out ?? 0;
  const byDay: any[] = usage?.byDay ?? usage?.by_day ?? [];
  const byPlugin: any[] = usage?.byPlugin ?? usage?.by_plugin ?? [];

  const budgetPct = monthlyBudget > 0 ? Math.min(100, Math.round((monthlyCost / monthlyBudget) * 100)) : 0;

  return (
    <div className="p-6 space-y-6">
      <div>
        <h1 className="text-xl font-semibold text-foreground">AI Usage</h1>
        <p className="text-sm text-muted-foreground mt-0.5">Token consumption and cost tracking across all plugins</p>
      </div>

      {/* Summary cards */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
        <div className="stat-card">
          <div className="flex items-center gap-2 mb-2">
            <DollarSign className="w-4 h-4 text-muted-foreground" />
            <span className="text-xs text-muted-foreground uppercase tracking-wide">Today</span>
          </div>
          <div className="text-2xl font-bold font-mono text-foreground">
            {loading ? "—" : `$${Number(todayCost).toFixed(2)}`}
          </div>
          <div className="text-xs text-muted-foreground mt-1">AUD</div>
        </div>
        <div className="stat-card">
          <div className="flex items-center gap-2 mb-2">
            <DollarSign className="w-4 h-4 text-muted-foreground" />
            <span className="text-xs text-muted-foreground uppercase tracking-wide">This month</span>
          </div>
          <div className="text-2xl font-bold font-mono text-foreground">
            {loading ? "—" : `$${Number(monthlyCost).toFixed(2)}`}
          </div>
          <div className="text-xs text-muted-foreground mt-1">of ${monthlyBudget} budget</div>
        </div>
        <div className="stat-card">
          <div className="flex items-center gap-2 mb-2">
            <Hash className="w-4 h-4 text-muted-foreground" />
            <span className="text-xs text-muted-foreground uppercase tracking-wide">API calls</span>
          </div>
          <div className="text-2xl font-bold font-mono text-foreground">
            {loading ? "—" : Number(totalCalls).toLocaleString()}
          </div>
          <div className="text-xs text-muted-foreground mt-1">this month</div>
        </div>
        <div className="stat-card">
          <div className="flex items-center gap-2 mb-2">
            <Zap className="w-4 h-4 text-muted-foreground" />
            <span className="text-xs text-muted-foreground uppercase tracking-wide">Tokens</span>
          </div>
          <div className="text-2xl font-bold font-mono text-foreground">
            {loading ? "—" : `${((totalTokensIn + totalTokensOut) / 1000000).toFixed(1)}M`}
          </div>
          <div className="text-xs text-muted-foreground mt-1">in + out</div>
        </div>
      </div>

      {/* Budget bar */}
      <div className="bg-white rounded-lg border border-border shadow-sm p-4">
        <div className="flex items-center justify-between mb-2">
          <span className="text-sm font-medium text-foreground">Monthly Budget</span>
          <span className="text-sm font-mono font-semibold text-foreground">
            ${Number(monthlyCost).toFixed(2)} / ${monthlyBudget}
          </span>
        </div>
        <div className="h-3 bg-slate-100 rounded-full overflow-hidden">
          <div
            className="h-full rounded-full transition-all"
            style={{
              width: `${budgetPct}%`,
              background: budgetPct > 80 ? "oklch(0.65 0.22 35)" : budgetPct > 60 ? "oklch(0.75 0.15 70)" : "oklch(0.5 0.2 250)",
            }}
          />
        </div>
        <div className="text-xs text-muted-foreground mt-1">{budgetPct}% used</div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Daily cost chart */}
        <div className="bg-white rounded-lg border border-border shadow-sm">
          <div className="px-4 py-3 border-b border-border">
            <span className="text-sm font-semibold text-foreground">Daily Cost (AUD)</span>
          </div>
          <div className="p-4">
            {byDay.length === 0 ? (
              <div className="h-[180px] flex items-center justify-center text-xs text-muted-foreground">
                {loading ? "Loading..." : "No cost data available yet"}
              </div>
            ) : (
              <ResponsiveContainer width="100%" height={180}>
                <BarChart data={byDay} barSize={28}>
                  <XAxis dataKey="date" tick={{ fontSize: 11, fill: "#94a3b8" }} axisLine={false} tickLine={false} />
                  <YAxis tick={{ fontSize: 11, fill: "#94a3b8" }} axisLine={false} tickLine={false} tickFormatter={v => `$${v}`} />
                  <Tooltip formatter={(v: number) => [`$${v.toFixed(2)}`, "Cost"]} contentStyle={{ fontSize: 12, borderRadius: 6 }} />
                  <Bar dataKey="cost" radius={[4, 4, 0, 0]}>
                    {byDay.map((_: any, i: number) => (
                      <Cell key={i} fill={i === byDay.length - 1 ? "oklch(0.5 0.2 250)" : "oklch(0.75 0.1 250)"} />
                    ))}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            )}
          </div>
        </div>

        {/* By plugin */}
        <div className="bg-white rounded-lg border border-border shadow-sm">
          <div className="px-4 py-3 border-b border-border">
            <span className="text-sm font-semibold text-foreground">Cost by Plugin</span>
          </div>
          <div className="divide-y divide-border">
            {byPlugin.length === 0 ? (
              <div className="px-4 py-8 text-center text-xs text-muted-foreground">
                {loading ? "Loading..." : "No plugin cost data yet"}
              </div>
            ) : byPlugin.map((p: any) => {
              const pct = monthlyCost > 0 ? Math.round((p.cost / monthlyCost) * 100) : 0;
              return (
                <div key={p.name} className="flex items-center gap-3 px-4 py-3">
                  <div className="flex-1 min-w-0">
                    <div className="flex items-center justify-between mb-1">
                      <span className="text-xs font-medium text-foreground truncate">{p.name}</span>
                      <span className="text-xs font-mono font-semibold text-foreground ml-2">${Number(p.cost).toFixed(2)}</span>
                    </div>
                    <div className="h-1.5 bg-slate-100 rounded-full overflow-hidden">
                      <div className="h-full rounded-full" style={{ width: `${pct}%`, background: "oklch(0.5 0.2 250)" }} />
                    </div>
                  </div>
                  <span className={`badge ${(p.model ?? "").includes("sonnet") ? "info" : "neutral"} flex-shrink-0`}>{p.model ?? "—"}</span>
                </div>
              );
            })}
          </div>
        </div>
      </div>
    </div>
  );
}
