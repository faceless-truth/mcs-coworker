// Design: Refined Dark Professional — Email Rules page
// Manage email routing rules that determine how CoWorker processes incoming emails

import { useEffect, useState } from "react";
import { fetchRules, saveRule, deleteRule } from "@/lib/api";
import { Filter, Plus, Trash2, ToggleLeft, ToggleRight, Mail } from "lucide-react";
import { toast } from "sonner";

interface Rule {
  id?: number;
  name: string;
  pattern: string;
  action: string;
  enabled: boolean;
  priority: number;
}

const EMPTY_RULE: Rule = { name: "", pattern: "", action: "", enabled: true, priority: 99 };

export default function EmailRules() {
  const [rules, setRules] = useState<Rule[]>([]);
  const [loading, setLoading] = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState<Rule>(EMPTY_RULE);
  const [saving, setSaving] = useState(false);

  const load = async () => {
    try {
      const data = await fetchRules() as Rule[];
      setRules(data);
    } catch {
      toast.error("Failed to load email rules");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); }, []);

  const handleSave = async () => {
    if (!form.name.trim() || !form.pattern.trim() || !form.action.trim()) {
      toast.error("Please fill in all fields");
      return;
    }
    setSaving(true);
    try {
      const updated = await saveRule(form) as Rule[];
      setRules(updated);
      setShowForm(false);
      setForm(EMPTY_RULE);
      toast.success("Rule saved");
    } catch (e: any) {
      toast.error(e.message || "Failed to save rule");
    } finally {
      setSaving(false);
    }
  };

  const handleDelete = async (id: number) => {
    try {
      await deleteRule(id);
      setRules(r => r.filter(x => x.id !== id));
      toast.success("Rule deleted");
    } catch {
      toast.error("Failed to delete rule");
    }
  };

  const inputClass = "w-full px-3 py-2 text-sm border border-border rounded-md bg-white focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-400";

  return (
    <div className="p-6 space-y-5">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-xl font-semibold text-foreground">Email Rules</h1>
          <p className="text-sm text-muted-foreground mt-0.5">
            Define how CoWorker routes and processes incoming emails
          </p>
        </div>
        <button
          onClick={() => { setShowForm(true); setForm(EMPTY_RULE); }}
          className="flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium text-white transition-all hover:opacity-90"
          style={{ background: "oklch(0.5 0.2 250)" }}
        >
          <Plus className="w-4 h-4" />
          Add Rule
        </button>
      </div>

      {/* Add Rule Form */}
      {showForm && (
        <div className="bg-white rounded-lg border border-blue-200 shadow-sm p-5 space-y-4"
          style={{ borderLeft: "3px solid oklch(0.5 0.2 250)" }}>
          <div className="text-sm font-semibold text-foreground">New Email Rule</div>
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
            <div>
              <label className="block text-xs font-medium text-foreground mb-1">Rule Name</label>
              <input className={inputClass} placeholder="e.g. ATO Correspondence" value={form.name}
                onChange={e => setForm(f => ({ ...f, name: e.target.value }))} />
            </div>
            <div>
              <label className="block text-xs font-medium text-foreground mb-1">Match Pattern</label>
              <input className={inputClass} placeholder="from:ato.gov.au OR subject:BAS" value={form.pattern}
                onChange={e => setForm(f => ({ ...f, pattern: e.target.value }))} />
            </div>
            <div>
              <label className="block text-xs font-medium text-foreground mb-1">Action</label>
              <input className={inputClass} placeholder="plugin:bas_reminder OR label:ATO" value={form.action}
                onChange={e => setForm(f => ({ ...f, action: e.target.value }))} />
            </div>
          </div>
          <div className="flex items-center gap-3">
            <button
              onClick={handleSave}
              disabled={saving}
              className="px-4 py-2 rounded-lg text-sm font-medium text-white transition-all hover:opacity-90 disabled:opacity-50"
              style={{ background: "oklch(0.5 0.2 250)" }}
            >
              {saving ? "Saving…" : "Save Rule"}
            </button>
            <button
              onClick={() => setShowForm(false)}
              className="px-4 py-2 rounded-lg text-sm font-medium border border-border hover:bg-slate-50 transition-colors"
            >
              Cancel
            </button>
          </div>
        </div>
      )}

      {/* Rules List */}
      {loading ? (
        <div className="bg-white rounded-lg border border-border p-8 text-center text-sm text-muted-foreground">
          Loading rules…
        </div>
      ) : rules.length === 0 ? (
        <div className="bg-white rounded-lg border border-border p-12 text-center">
          <Filter className="w-8 h-8 mx-auto mb-3 text-muted-foreground opacity-40" />
          <div className="text-sm font-medium text-foreground mb-1">No email rules yet</div>
          <div className="text-xs text-muted-foreground">Add rules to control how CoWorker routes incoming emails</div>
        </div>
      ) : (
        <div className="bg-white rounded-lg border border-border shadow-sm overflow-hidden">
          <div className="px-5 py-3 border-b border-border bg-slate-50 flex items-center justify-between">
            <span className="text-sm font-semibold text-foreground">Active Rules</span>
            <span className="text-xs text-muted-foreground">{rules.length} rule{rules.length !== 1 ? "s" : ""}</span>
          </div>
          <div className="divide-y divide-border">
            {rules.map((rule) => (
              <div key={rule.id} className="flex items-center gap-4 px-5 py-4 hover:bg-slate-50 transition-colors group">
                {/* Priority badge */}
                <div className="w-6 h-6 rounded-full flex items-center justify-center text-xs font-bold flex-shrink-0"
                  style={{ background: "oklch(0.94 0.05 250)", color: "oklch(0.45 0.2 250)" }}>
                  {rule.priority}
                </div>

                {/* Rule info */}
                <div className="flex-1 min-w-0">
                  <div className="flex items-center gap-2">
                    <span className="text-sm font-medium text-foreground">{rule.name}</span>
                    {rule.enabled ? (
                      <span className="badge success">Active</span>
                    ) : (
                      <span className="badge neutral">Disabled</span>
                    )}
                  </div>
                  <div className="flex items-center gap-3 mt-0.5">
                    <span className="text-xs text-muted-foreground font-mono truncate">
                      <span className="text-blue-500">match:</span> {rule.pattern}
                    </span>
                    <span className="text-xs text-muted-foreground">→</span>
                    <span className="text-xs text-muted-foreground font-mono truncate">
                      <span className="text-emerald-600">action:</span> {rule.action}
                    </span>
                  </div>
                </div>

                {/* Actions */}
                <div className="flex items-center gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                  <button
                    onClick={() => handleDelete(rule.id!)}
                    className="w-7 h-7 rounded-md flex items-center justify-center text-muted-foreground hover:text-rose-600 hover:bg-rose-50 transition-colors"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                  </button>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Help text */}
      <div className="bg-blue-50 rounded-lg border border-blue-100 p-4">
        <div className="flex items-start gap-3">
          <Mail className="w-4 h-4 mt-0.5 flex-shrink-0" style={{ color: "oklch(0.5 0.2 250)" }} />
          <div className="text-xs text-blue-700 space-y-1">
            <div className="font-semibold">Pattern syntax</div>
            <div><code className="bg-blue-100 px-1 rounded">from:domain.com</code> — match sender domain</div>
            <div><code className="bg-blue-100 px-1 rounded">subject:keyword</code> — match subject line</div>
            <div><code className="bg-blue-100 px-1 rounded">plugin:plugin_id</code> — trigger a specific plugin</div>
            <div><code className="bg-blue-100 px-1 rounded">label:name</code> — apply a label</div>
          </div>
        </div>
      </div>
    </div>
  );
}
