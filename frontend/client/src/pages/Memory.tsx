// Design: Refined Dark Professional — Memory Browser page

import { useEffect, useState } from "react";
import { Brain, Search, Trash2, Pencil, Check, X, Copy, User, Bot } from "lucide-react";
import { toast } from "sonner";
import { Streamdown } from "streamdown";
import {
  fetchMemory,
  fetchMemoryStats,
  deleteMemory,
  updateMemory,
  fetchMemoryDetail,
  type MemoryDetail,
} from "@/lib/api";

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
  entity: string;
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

interface EditState {
  client: string;
  entity: string;
}

interface ChatTurn {
  role: "user" | "assistant" | string;
  content: string;
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

function formatAbsoluteDate(meta: Record<string, any>): string {
  const raw = meta.timestamp ?? meta.date ?? meta.stored_at;
  let ms = NaN;
  if (typeof raw === "string") ms = Date.parse(raw);
  if (!Number.isFinite(ms) && typeof meta.epoch === "number") {
    ms = meta.epoch * 1000;
  }
  if (!Number.isFinite(ms)) return "";
  return new Date(ms).toLocaleString(undefined, {
    day: "numeric",
    month: "short",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function parseChatTurns(meta: Record<string, any>, fallbackContent: string): ChatTurn[] {
  // Prefer the structured `full_messages` blob written by save-conversation.
  const raw = meta.full_messages;
  if (typeof raw === "string" && raw.trim().startsWith("[")) {
    try {
      const arr = JSON.parse(raw);
      if (Array.isArray(arr)) {
        return arr
          .filter(m => m && typeof m === "object")
          .map(m => ({
            role: String((m as any).role || ""),
            content: String((m as any).content || ""),
          }));
      }
    } catch {
      // fall through to text parsing
    }
  }
  // Fallback for older entries that only have the "Q: ... / A: ..." summary.
  const turns: ChatTurn[] = [];
  let role: "user" | "assistant" | null = null;
  let buffer: string[] = [];
  const flush = () => {
    if (role && buffer.length) {
      turns.push({ role, content: buffer.join("\n").trim() });
    }
    buffer = [];
  };
  for (const line of (fallbackContent || "").split("\n")) {
    if (line.startsWith("Q: ")) {
      flush();
      role = "user";
      buffer.push(line.slice(3));
    } else if (line.startsWith("A: ")) {
      flush();
      role = "assistant";
      buffer.push(line.slice(3));
    } else {
      buffer.push(line);
    }
  }
  flush();
  return turns;
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
    entity: String(meta.entity_name ?? ""),
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
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editState, setEditState] = useState<EditState>({ client: "", entity: "" });
  const [saving, setSaving] = useState(false);
  const [expandedId, setExpandedId] = useState<string | null>(null);
  const [detailCache, setDetailCache] = useState<Record<string, MemoryDetail>>({});
  const [detailLoading, setDetailLoading] = useState<string | null>(null);

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
  }, [debouncedSearch]);

  const deleteRecord = async (id: string) => {
    try {
      await deleteMemory(id);
      setRecords(prev => prev.filter(r => r.id !== id));
      if (expandedId === id) setExpandedId(null);
      toast("Memory record deleted");
    } catch (e: any) {
      toast.error("Delete failed", { description: e?.message });
    }
  };

  const startEdit = (mem: ViewMemory) => {
    setEditingId(mem.id);
    setEditState({ client: mem.client, entity: mem.entity });
  };

  const cancelEdit = () => {
    setEditingId(null);
    setEditState({ client: "", entity: "" });
  };

  const saveEdit = async (id: string) => {
    if (!editState.client.trim()) {
      toast.error("Client name cannot be empty");
      return;
    }
    setSaving(true);
    try {
      await updateMemory(id, {
        client_name: editState.client.trim(),
        ...(editState.entity.trim() ? { entity_name: editState.entity.trim() } : {}),
      });
      setRecords(prev => prev.map(r =>
        r.id === id
          ? { ...r, client: editState.client.trim(), entity: editState.entity.trim() }
          : r
      ));
      toast.success("Memory record updated");
      setEditingId(null);
    } catch (e: any) {
      toast.error("Update failed", { description: e?.message });
    } finally {
      setSaving(false);
    }
  };

  const toggleExpand = async (id: string) => {
    if (editingId === id) return; // don't expand while editing
    if (expandedId === id) {
      setExpandedId(null);
      return;
    }
    setExpandedId(id);
    if (!detailCache[id]) {
      setDetailLoading(id);
      try {
        const detail = await fetchMemoryDetail(id);
        if (detail) {
          setDetailCache(prev => ({ ...prev, [id]: detail }));
        } else {
          toast.error("Could not load entry details");
        }
      } catch (e: any) {
        toast.error("Could not load entry", { description: e?.message });
      } finally {
        setDetailLoading(null);
      }
    }
  };

  const copyContent = async (text: string) => {
    try {
      await navigator.clipboard.writeText(text);
      toast.success("Copied to clipboard");
    } catch {
      toast.error("Copy failed");
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
        {records.map(mem => {
          const isExpanded = expandedId === mem.id;
          const detail = detailCache[mem.id];
          const isLoadingDetail = detailLoading === mem.id;
          return (
          <div key={mem.id}>
            <div
              className={`flex items-start gap-4 px-4 py-4 group ${editingId !== mem.id ? "cursor-pointer hover:bg-muted/40 transition-colors" : ""}`}
              onClick={() => editingId !== mem.id && toggleExpand(mem.id)}
            >
              <div className="w-8 h-8 rounded-full flex items-center justify-center flex-shrink-0 mt-0.5"
                style={{ background: "oklch(0.94 0.05 300)" }}>
                <Brain className="w-4 h-4" style={{ color: "oklch(0.5 0.15 300)" }} />
              </div>

              <div className="flex-1 min-w-0">
                {editingId === mem.id ? (
                  /* ── Edit mode ── */
                  <div className="space-y-2 mb-2" onClick={e => e.stopPropagation()}>
                    <div className="flex items-center gap-2">
                      <label className="text-xs text-muted-foreground w-16 flex-shrink-0">Client</label>
                      <input
                        autoFocus
                        value={editState.client}
                        onChange={e => setEditState(s => ({ ...s, client: e.target.value }))}
                        onKeyDown={e => { if (e.key === "Enter") saveEdit(mem.id); if (e.key === "Escape") cancelEdit(); }}
                        placeholder="Surname, First Name"
                        className="flex-1 text-sm border border-blue-300 rounded px-2 py-1 focus:outline-none focus:ring-2 focus:ring-blue-400/30"
                      />
                    </div>
                    <div className="flex items-center gap-2">
                      <label className="text-xs text-muted-foreground w-16 flex-shrink-0">Entity</label>
                      <input
                        value={editState.entity}
                        onChange={e => setEditState(s => ({ ...s, entity: e.target.value }))}
                        onKeyDown={e => { if (e.key === "Enter") saveEdit(mem.id); if (e.key === "Escape") cancelEdit(); }}
                        placeholder="e.g. 000 Pool Care Pty Ltd (optional)"
                        className="flex-1 text-sm border border-border rounded px-2 py-1 focus:outline-none focus:ring-2 focus:ring-blue-400/30"
                      />
                    </div>
                    <div className="flex items-center gap-2 pt-1">
                      <button
                        onClick={() => saveEdit(mem.id)}
                        disabled={saving}
                        className="flex items-center gap-1 text-xs px-2.5 py-1 bg-blue-600 text-white rounded hover:bg-blue-700 disabled:opacity-50 transition-colors"
                      >
                        <Check className="w-3 h-3" />
                        {saving ? "Saving…" : "Save"}
                      </button>
                      <button
                        onClick={cancelEdit}
                        className="flex items-center gap-1 text-xs px-2.5 py-1 border border-border rounded hover:bg-muted transition-colors"
                      >
                        <X className="w-3 h-3" />
                        Cancel
                      </button>
                    </div>
                  </div>
                ) : (
                  /* ── View mode ── */
                  <div className="flex items-center gap-2 mb-1">
                    <span className="text-sm font-semibold text-foreground">{mem.client}</span>
                    {mem.entity && (
                      <span className="text-xs text-muted-foreground">— {mem.entity}</span>
                    )}
                    <span className="badge info">{mem.type}</span>
                    {mem.date && <span className="text-xs text-muted-foreground">{mem.date}</span>}
                  </div>
                )}
                {editingId !== mem.id && (
                  <div className="text-sm text-muted-foreground leading-relaxed">{mem.summary}</div>
                )}
              </div>

              <div className="flex items-center gap-1.5 flex-shrink-0" onClick={e => e.stopPropagation()}>
                {isSearching && mem.relevance !== null && editingId !== mem.id && (
                  <div className="text-right mr-1">
                    <div className="text-sm font-mono font-semibold text-foreground">
                      {Math.round(mem.relevance * 100)}%
                    </div>
                    <div className="text-xs text-muted-foreground">relevance</div>
                  </div>
                )}
                {editingId !== mem.id && (
                  <>
                    {isExpanded ? (
                      <button
                        onClick={() => setExpandedId(null)}
                        className="p-1.5 rounded hover:bg-muted transition-all"
                        title="Collapse"
                      >
                        <X className="w-3.5 h-3.5" />
                      </button>
                    ) : (
                      <>
                        <button
                          onClick={() => startEdit(mem)}
                          className="p-1.5 rounded opacity-0 group-hover:opacity-100 hover:bg-blue-50 hover:text-blue-500 transition-all"
                          title="Edit client / entity"
                        >
                          <Pencil className="w-3.5 h-3.5" />
                        </button>
                        <button
                          onClick={() => deleteRecord(mem.id)}
                          className="p-1.5 rounded opacity-0 group-hover:opacity-100 hover:bg-rose-50 hover:text-rose-500 transition-all"
                          title="Delete record"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </>
                    )}
                  </>
                )}
              </div>
            </div>

            {isExpanded && (
              <div className="px-4 pb-5 pt-1 bg-muted/30 border-t border-border">
                {isLoadingDetail && !detail ? (
                  <div className="py-6 text-center text-sm text-muted-foreground">
                    Loading full content…
                  </div>
                ) : detail ? (
                  <ExpandedDetail
                    detail={detail}
                    type={mem.type}
                    onCopy={() => copyContent(detail.content || "")}
                  />
                ) : (
                  <div className="py-6 text-center text-sm text-muted-foreground">
                    Could not load full content.
                  </div>
                )}
              </div>
            )}
          </div>
          );
        })}
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

interface ExpandedDetailProps {
  detail: MemoryDetail;
  type: string;
  onCopy: () => void;
}

function ExpandedDetail({ detail, type, onCopy }: ExpandedDetailProps) {
  const meta = detail.metadata || {};
  const isChat = type === "ai_chat_conversation";
  const turns = isChat ? parseChatTurns(meta, detail.content) : [];
  const agentName = String(meta.agent_name ?? "");
  const messageCount = typeof meta.message_count === "number" ? meta.message_count : turns.length;
  const absoluteDate = formatAbsoluteDate(meta);

  return (
    <div className="rounded-lg bg-white border border-border p-4 space-y-3">
      <div className="flex flex-wrap gap-x-4 gap-y-1 text-xs text-muted-foreground">
        <span><span className="font-medium text-foreground">Type:</span> {type}</span>
        {agentName && <span><span className="font-medium text-foreground">Agent:</span> {agentName}</span>}
        {absoluteDate && <span><span className="font-medium text-foreground">Date:</span> {absoluteDate}</span>}
        {isChat && <span><span className="font-medium text-foreground">Messages:</span> {messageCount}</span>}
      </div>

      <div className="border-t border-border pt-3">
        {isChat && turns.length > 0 ? (
          <div className="space-y-3">
            {turns.map((turn, i) => (
              <ChatTurnView key={i} turn={turn} agentName={agentName} />
            ))}
          </div>
        ) : (
          <div className="prose prose-sm max-w-none text-sm text-foreground">
            <Streamdown>{detail.content || "(no content)"}</Streamdown>
          </div>
        )}
      </div>

      <div className="flex items-center gap-2 pt-2 border-t border-border">
        <button
          onClick={onCopy}
          className="flex items-center gap-1.5 text-xs px-2.5 py-1.5 border border-border rounded hover:bg-muted transition-colors"
        >
          <Copy className="w-3 h-3" />
          Copy
        </button>
      </div>
    </div>
  );
}

function ChatTurnView({ turn, agentName }: { turn: ChatTurn; agentName: string }) {
  const isUser = turn.role === "user";
  return (
    <div className="flex items-start gap-2.5">
      <div
        className="w-7 h-7 rounded-full flex items-center justify-center flex-shrink-0 mt-0.5"
        style={{ background: isUser ? "oklch(0.94 0.04 250)" : "oklch(0.94 0.05 160)" }}
      >
        {isUser
          ? <User className="w-3.5 h-3.5" style={{ color: "oklch(0.45 0.15 250)" }} />
          : <Bot className="w-3.5 h-3.5" style={{ color: "oklch(0.45 0.15 160)" }} />}
      </div>
      <div className="flex-1 min-w-0">
        <div className="text-xs font-medium text-muted-foreground mb-1">
          {isUser ? "You" : (agentName || "Assistant")}
        </div>
        {isUser ? (
          <div className="text-sm text-foreground whitespace-pre-wrap break-words">
            {turn.content}
          </div>
        ) : (
          <div className="prose prose-sm max-w-none text-sm text-foreground">
            <Streamdown>{turn.content}</Streamdown>
          </div>
        )}
      </div>
    </div>
  );
}
