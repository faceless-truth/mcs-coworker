// Design: Refined Dark Professional — Memory Browser page

import { useEffect, useState } from "react";
import { Brain, Search, Trash2 } from "lucide-react";
import { toast } from "sonner";
import { fetchMemory, deleteMemory } from "@/lib/api";

interface ServerMemory {
  content: string;
  metadata: Record<string, any>;
  distance: number;
}

interface ViewMemory {
  id: string;
  client: string;
  type: string;
  summary: string;
  date: string;
  relevance: number;
}

function normalise(raw: ServerMemory, idx: number): ViewMemory {
  const meta = raw.metadata || {};
  return {
    id: String(meta.id ?? meta.doc_id ?? idx),
    client: String(meta.client_name ?? meta.client_email ?? meta.client ?? "Unknown"),
    type: String(meta.type ?? meta.interaction_type ?? "entry"),
    summary: raw.content || "",
    date: String(meta.date ?? meta.stored_at ?? ""),
    // distance is 0 (identical) → 2 (opposite). Invert to a 0..1 relevance.
    relevance: Math.max(0, Math.min(1, 1 - (raw.distance ?? 0) / 2)),
  };
}

export default function Memory() {
  const [records, setRecords] = useState<ViewMemory[]>([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");

  useEffect(() => {
    let cancelled = false;
    const load = async () => {
      try {
        const rows = (await fetchMemory(search || "recent client interactions", 50)) as ServerMemory[] | null;
        if (!cancelled) {
          setRecords(Array.isArray(rows) ? rows.map(normalise) : []);
        }
      } catch {
        if (!cancelled) setRecords([]);
      } finally {
        if (!cancelled) setLoading(false);
      }
    };
    load();
    return () => { cancelled = true; };
  }, [search]);

  const filtered = records.filter(r =>
    search === "" ||
    r.client.toLowerCase().includes(search.toLowerCase()) ||
    r.summary.toLowerCase().includes(search.toLowerCase())
  );

  const deleteRecord = async (id: string) => {
    try {
      await deleteMemory(id);
      setRecords(prev => prev.filter(r => r.id !== id));
      toast("Memory record deleted");
    } catch (e: any) {
      toast.error("Delete failed", { description: e?.message });
    }
  };

  return (
    <div className="p-6 space-y-5">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-xl font-semibold text-foreground">Memory</h1>
          <p className="text-sm text-muted-foreground mt-0.5">
            {loading ? "Loading…" : `${records.length} client interactions stored in the vector memory`}
          </p>
        </div>
      </div>

      <div className="relative">
        <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground" />
        <input
          value={search}
          onChange={e => setSearch(e.target.value)}
          placeholder="Search by client name or content..."
          className="w-full pl-9 pr-4 py-2 text-sm border border-border rounded-lg bg-white focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-400"
        />
      </div>

      <div className="bg-white rounded-lg border border-border shadow-sm divide-y divide-border">
        {filtered.map(mem => (
          <div key={mem.id} className="flex items-start gap-4 px-4 py-4 group">
            <div className="w-8 h-8 rounded-full flex items-center justify-center flex-shrink-0 mt-0.5"
              style={{ background: "oklch(0.94 0.05 300)" }}>
              <Brain className="w-4 h-4" style={{ color: "oklch(0.5 0.15 300)" }} />
            </div>
            <div className="flex-1 min-w-0">
              <div className="flex items-center gap-2 mb-1">
                <span className="text-sm font-semibold text-foreground">{mem.client}</span>
                <span className="badge info">{mem.type}</span>
                <span className="text-xs text-muted-foreground">{mem.date}</span>
              </div>
              <div className="text-sm text-muted-foreground leading-relaxed">{mem.summary}</div>
            </div>
            <div className="flex items-center gap-3 flex-shrink-0">
              <div className="text-right">
                <div className="text-sm font-mono font-semibold text-foreground">{Math.round(mem.relevance * 100)}%</div>
                <div className="text-xs text-muted-foreground">relevance</div>
              </div>
              <button
                onClick={() => deleteRecord(mem.id)}
                className="p-1.5 rounded opacity-0 group-hover:opacity-100 hover:bg-rose-50 hover:text-rose-500 transition-all"
              >
                <Trash2 className="w-3.5 h-3.5" />
              </button>
            </div>
          </div>
        ))}
        {!loading && filtered.length === 0 && (
          <div className="py-12 text-center text-sm text-muted-foreground">
            {search
              ? "No records match your search."
              : "No memory records yet — plugins will populate this as they run."}
          </div>
        )}
      </div>
    </div>
  );
}
