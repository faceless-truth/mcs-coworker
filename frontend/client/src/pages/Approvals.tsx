// Design: Refined Dark Professional — Approvals page

import { useEffect, useState } from "react";
import { CheckCircle2, ChevronDown, ChevronUp, XCircle } from "lucide-react";
import { toast } from "sonner";
import { fetchApprovals, approveAction, rejectAction } from "@/lib/api";

interface ServerApproval {
  action_id: number | string;
  plugin_id: string;
  action_type: string;
  description: string;
  payload: Record<string, any>;
  confidence: number;
  status: string;
  created_at: string;
  expires_at?: string;
}

interface ViewApproval {
  id: string;
  plugin: string;
  action: string;
  detail: string;
  confidence: number;
  created: string;
  priority: "high" | "medium" | "low";
  draftPreview: string;
}

function normalise(raw: ServerApproval): ViewApproval {
  const payload = raw.payload || {};
  const detailParts: string[] = [];
  if (payload.recipient || payload.to) detailParts.push(String(payload.recipient || payload.to));
  if (payload.subject) detailParts.push(String(payload.subject));
  if (payload.balance) detailParts.push(`$${payload.balance}`);
  if (payload.days_overdue) detailParts.push(`${payload.days_overdue} days overdue`);
  const priority: "high" | "medium" | "low" =
    raw.confidence < 0.5 ? "high" : raw.confidence < 0.75 ? "medium" : "low";
  return {
    id: String(raw.action_id),
    plugin: raw.plugin_id || raw.action_type || "unknown",
    action: raw.description || raw.action_type || "Pending action",
    detail: detailParts.join(" · ") || "—",
    confidence: typeof raw.confidence === "number" ? raw.confidence : 0,
    created: raw.created_at || "",
    priority,
    draftPreview:
      String(payload.body_preview || payload.body || payload.draft || "(no preview available)"),
  };
}

export default function Approvals() {
  const [approvals, setApprovals] = useState<ViewApproval[]>([]);
  const [loading, setLoading] = useState(true);
  const [expanded, setExpanded] = useState<string | null>(null);
  const [threshold, setThreshold] = useState(0.75);

  const load = async () => {
    try {
      const rows = (await fetchApprovals()) as ServerApproval[] | null;
      setApprovals(Array.isArray(rows) ? rows.map(normalise) : []);
    } catch (e: any) {
      toast.error("Failed to load approvals", { description: e?.message });
      setApprovals([]);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    load();
    const id = setInterval(load, 15000);
    return () => clearInterval(id);
  }, []);

  const approve = async (id: string) => {
    const a = approvals.find(x => x.id === id);
    try {
      await approveAction(id);
      toast.success(`Approved: ${a?.action ?? id}`, { description: "Action will execute shortly." });
      setApprovals(prev => prev.filter(x => x.id !== id));
    } catch (e: any) {
      toast.error("Approve failed", { description: e?.message });
    }
  };

  const reject = async (id: string) => {
    const a = approvals.find(x => x.id === id);
    try {
      await rejectAction(id);
      toast(`Rejected: ${a?.action ?? id}`, { description: "Action discarded." });
      setApprovals(prev => prev.filter(x => x.id !== id));
    } catch (e: any) {
      toast.error("Reject failed", { description: e?.message });
    }
  };

  return (
    <div className="p-6 space-y-5">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-xl font-semibold text-foreground">Approvals</h1>
          <p className="text-sm text-muted-foreground mt-0.5">
            Actions below the confidence threshold waiting for human review
          </p>
        </div>
        <div className="flex items-center gap-3 bg-white border border-border rounded-lg px-4 py-2.5 shadow-sm">
          <span className="text-xs text-muted-foreground">Confidence threshold</span>
          <input
            type="range" min={0.5} max={1} step={0.05} value={threshold}
            onChange={e => setThreshold(parseFloat(e.target.value))}
            className="w-24"
          />
          <span className="text-sm font-mono font-semibold text-foreground w-10">{Math.round(threshold * 100)}%</span>
        </div>
      </div>

      {loading ? (
        <div className="py-12 text-center text-sm text-muted-foreground">Loading approvals…</div>
      ) : approvals.length === 0 ? (
        <div className="flex flex-col items-center justify-center py-20 text-center">
          <CheckCircle2 className="w-12 h-12 text-emerald-400 mb-4" />
          <div className="text-lg font-semibold text-foreground">All clear</div>
          <div className="text-sm text-muted-foreground mt-1">No actions pending approval right now.</div>
        </div>
      ) : (
        <div className="space-y-3">
          {approvals.map((a) => {
            const isExpanded = expanded === a.id;
            const confPct = Math.round(a.confidence * 100);
            const confColor = a.confidence >= 0.70 ? "text-amber-600" : "text-rose-600";
            return (
              <div key={a.id} className="bg-white rounded-lg border border-border shadow-sm overflow-hidden">
                <div className="flex items-start gap-4 p-4">
                  <div className="flex-shrink-0 mt-0.5">
                    <div className={`w-2 h-2 rounded-full mt-1.5 ${a.priority === "high" ? "bg-rose-500" : a.priority === "medium" ? "bg-amber-500" : "bg-slate-400"}`} />
                  </div>
                  <div className="flex-1 min-w-0">
                    <div className="flex items-center gap-2 mb-1">
                      <span className="text-xs font-medium text-muted-foreground">{a.plugin}</span>
                      <span className="text-xs text-muted-foreground">·</span>
                      <span className="text-xs text-muted-foreground">{a.created}</span>
                    </div>
                    <div className="text-sm font-semibold text-foreground">{a.action}</div>
                    <div className="text-xs text-muted-foreground mt-0.5">{a.detail}</div>
                  </div>
                  <div className="flex items-center gap-3 flex-shrink-0">
                    <div className="text-right">
                      <div className={`text-sm font-mono font-bold ${confColor}`}>{confPct}%</div>
                      <div className="text-xs text-muted-foreground">confidence</div>
                    </div>
                    <button
                      onClick={() => setExpanded(isExpanded ? null : a.id)}
                      className="p-1.5 rounded hover:bg-slate-100 transition-colors"
                    >
                      {isExpanded ? <ChevronUp className="w-4 h-4 text-muted-foreground" /> : <ChevronDown className="w-4 h-4 text-muted-foreground" />}
                    </button>
                  </div>
                </div>

                {isExpanded && (
                  <div className="px-4 pb-4 border-t border-border pt-3">
                    <div className="text-xs font-medium text-muted-foreground mb-2">Draft preview</div>
                    <pre className="text-xs text-foreground bg-slate-50 rounded-md p-3 whitespace-pre-wrap font-mono leading-relaxed border border-border">
                      {a.draftPreview}
                    </pre>
                  </div>
                )}

                <div className="flex items-center gap-2 px-4 py-3 bg-slate-50 border-t border-border">
                  <button
                    onClick={() => approve(a.id)}
                    className="flex items-center gap-1.5 px-3 py-1.5 rounded text-xs font-medium text-white transition-all"
                    style={{ background: "oklch(0.5 0.15 145)" }}
                  >
                    <CheckCircle2 className="w-3.5 h-3.5" />
                    Approve & Send
                  </button>
                  <button
                    onClick={() => reject(a.id)}
                    className="flex items-center gap-1.5 px-3 py-1.5 rounded text-xs font-medium border border-border hover:border-rose-300 hover:text-rose-600 transition-all"
                  >
                    <XCircle className="w-3.5 h-3.5" />
                    Reject
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
