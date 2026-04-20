// Design: Refined Dark Professional — Activity Log page with 3 sub-tabs
// All data fetched from real /api/* endpoints — NO mock data

import { useEffect, useRef, useState } from "react";
import { AlertTriangle, Brain, CheckCircle2, Clock, Radio, Search, Wifi, WifiOff } from "lucide-react";

const API_BASE = "http://127.0.0.1:7842";
const SSE_URL = `${API_BASE}/api/activity/stream`;

type SubTab = "email" | "memory" | "events";

interface ActivityEntry {
  id: string | number;
  plugin: string;
  action: string;
  status: "success" | "error" | "warning";
  time: string;
}

// ── SSE hook — connects to real backend stream ─────────────────────────────────
function useActivityStream(enabled: boolean) {
  const [entries, setEntries] = useState<ActivityEntry[]>([]);
  const [connected, setConnected] = useState(false);
  const [liveCount, setLiveCount] = useState(0);
  const esRef = useRef<EventSource | null>(null);

  // Also do an initial REST fetch for recent history
  useEffect(() => {
    const fetchHistory = async () => {
      try {
        const res = await fetch(`${API_BASE}/api/activity`);
        if (res.ok) {
          const json = await res.json();
          const data = (json && typeof json === "object" && "data" in json) ? json.data : json;
          if (Array.isArray(data) && data.length > 0) {
            setEntries(data);
          }
        }
      } catch (_) {}
    };
    fetchHistory();
  }, []);

  useEffect(() => {
    if (!enabled) return;

    const connect = () => {
      const es = new EventSource(SSE_URL);
      esRef.current = es;

      es.onopen = () => setConnected(true);

      es.onmessage = (e) => {
        try {
          const entry: ActivityEntry = JSON.parse(e.data);
          setEntries(prev => [entry, ...prev].slice(0, 200));
          setLiveCount(n => n + 1);
        } catch {
          // ignore malformed events
        }
      };

      es.addEventListener("ping", () => {
        // keepalive — no action needed
      });

      es.onerror = () => {
        setConnected(false);
        es.close();
        setTimeout(connect, 5000);
      };
    };

    connect();
    return () => {
      esRef.current?.close();
    };
  }, [enabled]);

  return { entries, connected, liveCount };
}

// ── Component ─────────────────────────────────────────────────────────────────
export default function Activity() {
  const [subTab, setSubTab] = useState<SubTab>("email");
  const [search, setSearch] = useState("");
  const [streaming, setStreaming] = useState(true);
  const [memory, setMemory] = useState<any[]>([]);
  const [events, setEvents] = useState<any[]>([]);

  const { entries, connected, liveCount } = useActivityStream(streaming);

  // Fetch memory and events when those tabs are opened
  useEffect(() => {
    const unwrap = (d: any) => (d && typeof d === "object" && "data" in d) ? d.data : d;
    if (subTab === "memory") {
      fetch(`${API_BASE}/api/memory`)
        .then(r => r.ok ? r.json() : [])
        .then(d => setMemory(Array.isArray(unwrap(d)) ? unwrap(d) : []))
        .catch(() => {});
    }
    if (subTab === "events") {
      fetch(`${API_BASE}/api/events`)
        .then(r => r.ok ? r.json() : [])
        .then(d => setEvents(Array.isArray(unwrap(d)) ? unwrap(d) : []))
        .catch(() => {});
    }
  }, [subTab]);

  const filteredLog = entries.filter(l =>
    search === "" ||
    l.plugin?.toLowerCase().includes(search.toLowerCase()) ||
    l.action?.toLowerCase().includes(search.toLowerCase())
  );

  return (
    <div className="p-6 space-y-5">
      <div className="flex items-start justify-between">
        <div>
          <h1 className="text-xl font-semibold text-foreground">Activity Log</h1>
          <p className="text-sm text-muted-foreground mt-0.5">Real-time view of all plugin actions, memory, and events</p>
        </div>
        {subTab === "email" && (
          <div className="flex items-center gap-2">
            <div className={`flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-medium border ${
              connected
                ? "border-emerald-200 text-emerald-700"
                : "border-slate-200 text-slate-500"
            }`}
              style={{ background: connected ? "oklch(0.97 0.04 145)" : "oklch(0.97 0 0)" }}>
              {connected
                ? <><Wifi className="w-3 h-3" /> Live</>
                : <><WifiOff className="w-3 h-3" /> Connecting...</>
              }
              {liveCount > 0 && (
                <span className="ml-1 px-1.5 py-0.5 rounded-full text-xs font-bold"
                  style={{ background: "oklch(0.5 0.15 145)", color: "white" }}>
                  +{liveCount}
                </span>
              )}
            </div>
            <button
              onClick={() => setStreaming(s => !s)}
              className={`px-2.5 py-1 rounded-lg text-xs font-medium border transition-colors ${
                streaming
                  ? "border-blue-200 text-blue-700 bg-blue-50 hover:bg-blue-100"
                  : "border-border text-muted-foreground bg-white hover:bg-slate-50"
              }`}
            >
              {streaming ? "Pause" : "Resume"}
            </button>
          </div>
        )}
      </div>

      {/* Sub-tabs */}
      <div className="flex gap-1 bg-slate-100 rounded-lg p-1 w-fit">
        {([
          { id: "email", label: "Email Activity", icon: CheckCircle2 },
          { id: "memory", label: "Memory Browser", icon: Brain },
          { id: "events", label: "Event Log", icon: Radio },
        ] as const).map(tab => {
          const Icon = tab.icon;
          return (
            <button
              key={tab.id}
              onClick={() => setSubTab(tab.id)}
              className={`flex items-center gap-2 px-3 py-2 rounded-md text-xs font-medium transition-all ${
                subTab === tab.id ? "bg-white shadow-sm text-foreground" : "text-muted-foreground hover:text-foreground"
              }`}
            >
              <Icon className="w-3.5 h-3.5" />
              {tab.label}
            </button>
          );
        })}
      </div>

      {subTab === "email" && (
        <div className="space-y-3">
          <div className="relative">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground" />
            <input
              value={search}
              onChange={e => setSearch(e.target.value)}
              placeholder="Search activity..."
              className="w-full pl-9 pr-4 py-2 text-sm border border-border rounded-lg bg-white focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-400"
            />
          </div>

          <div className="bg-white rounded-lg border border-border shadow-sm divide-y divide-border">
            {filteredLog.length === 0 ? (
              <div className="px-4 py-8 text-center text-sm text-muted-foreground">
                {entries.length === 0 ? "No activity yet — plugins will log here when they run." : "No results match your search."}
              </div>
            ) : filteredLog.slice(0, 50).map((log, i) => (
              <div
                key={log.id}
                className={`flex items-start gap-3 px-4 py-3 transition-colors ${
                  i === 0 && liveCount > 0 ? "bg-emerald-50/60" : ""
                }`}
              >
                <div className="flex-shrink-0 mt-0.5">
                  {log.status === "success" && <CheckCircle2 className="w-4 h-4 text-emerald-500" />}
                  {log.status === "error" && <AlertTriangle className="w-4 h-4 text-rose-500" />}
                  {log.status === "warning" && <Clock className="w-4 h-4 text-amber-500" />}
                </div>
                <div className="flex-1 min-w-0">
                  <div className="flex items-center gap-2 mb-0.5">
                    <span className="text-xs font-semibold text-foreground">{log.plugin}</span>
                    <span className={`badge ${log.status === "success" ? "success" : log.status === "error" ? "error" : "warning"}`}>
                      {log.status}
                    </span>
                    {i === 0 && liveCount > 0 && (
                      <span className="badge" style={{ background: "oklch(0.94 0.06 145)", color: "oklch(0.35 0.12 145)" }}>
                        new
                      </span>
                    )}
                  </div>
                  <div className="text-sm text-muted-foreground">{log.action}</div>
                </div>
                <span className="text-xs text-muted-foreground flex-shrink-0 font-mono">{log.time}</span>
              </div>
            ))}
          </div>
        </div>
      )}

      {subTab === "memory" && (
        <div className="space-y-3">
          <div className="bg-white rounded-lg border border-border shadow-sm divide-y divide-border">
            {memory.length === 0 ? (
              <div className="px-4 py-8 text-center text-sm text-muted-foreground">No memory records yet.</div>
            ) : memory.map((mem: any) => (
              <div key={mem.id} className="flex items-start gap-4 px-4 py-4">
                <Brain className="w-4 h-4 text-purple-500 flex-shrink-0 mt-0.5" />
                <div className="flex-1 min-w-0">
                  <div className="flex items-center gap-2 mb-1">
                    <span className="text-sm font-semibold text-foreground">{mem.client}</span>
                    <span className="badge info">{mem.type}</span>
                  </div>
                  <div className="text-sm text-muted-foreground">{mem.summary}</div>
                  <div className="text-xs text-muted-foreground mt-1">{mem.date}</div>
                </div>
                {mem.relevance != null && (
                  <div className="text-right flex-shrink-0">
                    <div className="text-xs font-mono font-semibold text-foreground">{Math.round(mem.relevance * 100)}%</div>
                    <div className="text-xs text-muted-foreground">relevance</div>
                  </div>
                )}
              </div>
            ))}
          </div>
        </div>
      )}

      {subTab === "events" && (
        <div className="bg-white rounded-lg border border-border shadow-sm divide-y divide-border">
          {events.length === 0 ? (
            <div className="px-4 py-8 text-center text-sm text-muted-foreground">No events logged yet.</div>
          ) : events.map((evt: any) => (
            <div key={evt.id} className="flex items-center gap-4 px-4 py-3">
              <span className="text-xs font-mono text-muted-foreground w-20 flex-shrink-0">{evt.time}</span>
              <span className="text-xs font-mono font-medium text-blue-600 flex-shrink-0">{evt.type}</span>
              <span className="text-xs text-muted-foreground flex-shrink-0">from {evt.source}</span>
              <span className="text-xs text-muted-foreground truncate">{evt.payload}</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
