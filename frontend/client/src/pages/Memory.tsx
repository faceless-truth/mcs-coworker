// Design: Refined Dark Professional — Memory Browser page

import { mockMemory } from "@/lib/mockData";
import { useState } from "react";
import { Brain, Search, Trash2 } from "lucide-react";
import { toast } from "sonner";

export default function Memory() {
  const [records, setRecords] = useState(mockMemory);
  const [search, setSearch] = useState("");

  const filtered = records.filter(r =>
    search === "" ||
    r.client.toLowerCase().includes(search.toLowerCase()) ||
    r.summary.toLowerCase().includes(search.toLowerCase())
  );

  const deleteRecord = (id: string) => {
    setRecords(prev => prev.filter(r => r.id !== id));
    toast("Memory record deleted", { description: "Record removed from the vector store." });
  };

  return (
    <div className="p-6 space-y-5">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-xl font-semibold text-foreground">Memory</h1>
          <p className="text-sm text-muted-foreground mt-0.5">
            {records.length} client interactions stored in the vector memory (ChromaDB)
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
        {filtered.length === 0 && (
          <div className="py-12 text-center text-sm text-muted-foreground">No records match your search.</div>
        )}
      </div>
    </div>
  );
}
