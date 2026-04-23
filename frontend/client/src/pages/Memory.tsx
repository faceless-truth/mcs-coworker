// Design: Refined Dark Professional — Memory Browser page

import { useEffect, useState } from "react";
import { Brain, Search, Trash2 } from "lucide-react";
import { toast } from "sonner";
import { fetchMemory, fetchMemoryStats, deleteMemory } from "@/lib/api";

interface ServerMemory {
  id?: string;
  content: string;
  metadata: Record<string, any>;
  distance?: number;
  relevance?: number | null;
}

interface ViewMemory {
  id: string;
  client: string;
  type: string;
  summary: string;
  date: string;
  relevance: number | null;
}

interface MemoryStats {
  interactions: number;
  lessons: number;
  documents: number;
}

function relativeTime(iso: string | undefined, epoch?: number): string {
  const ms = iso ? Date.parse(iso) : (epoch ? epoch * 1000 : NaN);
  if (!Number.isFinite(ms)) return "";
  const diff = Date.now() - ms;
  const min = Math.floor(diff / 60_000);
  if (min < 1) return "just now";
  if (min < 60) return `${min}m ago`;
  const hr = Math.floor(min / 60);
  if (hr < 24) return `${hr}h ago`;
  const d = Math.floor(hr / 24);
  if (d < 30) return `${d}d ago`;
  return new Date(ms).toLocaleDateString();
}

function normalise(raw: ServerMemory, idx: number): ViewMemory {
  const meta = raw.metadata || {};
  const rel =
    raw.relevance !== undefined && raw.relevance !== null
      ? raw.relevance
      : (typeof raw.distance === "number"
          ? Math.max(0, Math.min(1, 1 - raw.distance))
          : null);
  return {
    id: String(raw.id ?? meta.id ?? meta.doc_id ?? idx),
    client: String(meta.client_name ?? meta.client_email ?? meta.client ?? "Unknown"),
    type: String(meta.interaction_type ?? meta.type ?? "entry"),
    summary: (raw.content || "").slice(0, 300),
    date: relativeTime(meta.timestamp ?? meta.date ?? meta.stored_at, meta.epoch),
    relevance: rel,
  };
}

export default function Memory() {
  const [records, setRecords] = useState<ViewMemory[]>([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");
  const [debouncedSearch, setDebouncedSearch] = useState("");
  const [stats, setStats] = useState<MemoryStats | null>(null);

  // Debounce the search term so every keystroke doesn't hit the backend.
  useEffect(() => {
    const t = setTimeout(() => setDebouncedSearch(search), 300);
    return () => clearTimeout(t);
  }, [search]);

  useEffect(() => {
    let cancelled = false;
    const load = async () => {
      setLoading(true);
      try {
        const rows = (await fetchMemory(debouncedSearch, 50)) as ServerMemory[] | null;
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
  }, [debouncedSearch]);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const s = (await fetchMemoryStats()) as MemoryStats | null;
        if (!cancelled && s) setStats(s);
      } catch {
        // non-fatal; header just omits counts
      }
    })();
    return () => { cancelled = true; };
  }, [debouncedSearch]); // refresh stats after searches too (new entries may have landed)

  const deleteRecord = async (id: string) => {
    try {
      await deleteMemory(id);
      setRecords(prev => prev.filter(r => r.id !== id));
      toast("Memory record deleted");
    } catch (e: any) {
      toast.error("Delete failed", { description: e?.message });
    }
  };

  const isSearching = debouncedSearch.trim().length > 0;
  const headerText = stats
    ? `${stats.interactions} client interaction${stats.interactions === 1 ? "" : "s"}, ${stats.lessons} lesson${stats.lessons === 1 ? "" : "s"}, ${stats.documents} document${stats.documents === 1 ? "" : "s"} in semantic memory`
    : (loading ? "Loading…" : "Semantic memory");

  return (
    <div className="p-6 space-y-5">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-xl font-semibold text-foreground">Memory</h1>
          <p className="text-sm text-muted-foreground mt-0.5">{headerText}</p>
        </div>
      </div>

      <div className="relative">
        <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground" />
        <input
          value={search}
          onChange={e => setSearch(e.target.value)}
          placeholder="Search by meaning, client, or content..."
          className="w-full pl-9 pr-4 py-2 text-sm border border-border rounded-lg bg-white focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-400"
        />
      </div>

      <div className="bg-white rounded-lg border border-border shadow-sm divide-y divide-border">
        {records.map(mem => (
          <div key={mem.id} className="flex items-start gap-4 px-4 py-4 group">
            <div className="w-8 h-8 rounded-full flex items-center justify-center flex-shrink-0 mt-0.5"
              style={{ background: "oklch(0.94 0.05 300)" }}>
              <Brain className="w-4 h-4" style={{ color: "oklch(0.5 0.15 300)" }} />
            </div>
            <div className="flex-1 min-w-0">
              <div className="flex items-center gap-2 mb-1">
                <span className="text-sm font-semibold text-foreground">{mem.client}</span>
                <span className="badge info">{mem.type}</span>
                {mem.date && <span className="text-xs text-muted-foreground">{mem.date}</span>}
              </div>
              <div className="text-sm text-muted-foreground leading-relaxed">{mem.summary}</div>
            </div>
            <div className="flex items-center gap-3 flex-shrink-0">
              {isSearching && mem.relevance !== null && (
                <div className="text-right">
                  <div className="text-sm font-mono font-semibold text-foreground">
                    {Math.round(mem.relevance * 100)}%
                  </div>
                  <div className="text-xs text-muted-foreground">relevance</div>
                </div>
              )}
              <button
                onClick={() => deleteRecord(mem.id)}
                className="p-1.5 rounded opacity-0 group-hover:opacity-100 hover:bg-rose-50 hover:text-rose-500 transition-all"
              >
                <Trash2 className="w-3.5 h-3.5" />
              </button>
            </div>
          </div>
        ))}
        {!loading && records.length === 0 && (
          <div className="py-12 text-center text-sm text-muted-foreground">
            {isSearching
              ? "No records match your search."
              : "No memory records yet — plugins will populate this as they process emails and documents."}
          </div>
        )}
      </div>
    </div>
  );
}
