// Design: Refined Dark Professional — Staff & Notify page
// Manage staff members and their notification preferences

import { useEffect, useState } from "react";
import { fetchStaff, saveStaff, deleteStaff } from "@/lib/api";
import { Plus, Trash2, Users, Bell, BellOff } from "lucide-react";
import { toast } from "sonner";

interface StaffMember {
  id?: number;
  name: string;
  email: string;
  role: string;
  notify_approvals: boolean;
  notify_briefing: boolean;
  notify_alerts: boolean;
}

const EMPTY_STAFF: StaffMember = {
  name: "", email: "", role: "",
  notify_approvals: false, notify_briefing: false, notify_alerts: false,
};

function Toggle({ value, onChange }: { value: boolean; onChange: (v: boolean) => void }) {
  return (
    <button
      onClick={() => onChange(!value)}
      className={`relative w-9 h-5 rounded-full transition-all flex-shrink-0`}
      style={{ background: value ? "oklch(0.5 0.2 250)" : "oklch(0.85 0.005 240)" }}
    >
      <span className={`absolute top-0.5 w-4 h-4 bg-white rounded-full shadow transition-all ${value ? "left-4" : "left-0.5"}`} />
    </button>
  );
}

export default function StaffNotify() {
  const [staff, setStaff] = useState<StaffMember[]>([]);
  const [loading, setLoading] = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState<StaffMember>(EMPTY_STAFF);
  const [saving, setSaving] = useState(false);

  const load = async () => {
    try {
      const data = await fetchStaff() as StaffMember[];
      setStaff(data);
    } catch {
      toast.error("Failed to load staff");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); }, []);

  const handleSave = async () => {
    if (!form.name.trim() || !form.email.trim()) {
      toast.error("Name and email are required");
      return;
    }
    setSaving(true);
    try {
      const updated = await saveStaff(form) as StaffMember[];
      setStaff(updated);
      setShowForm(false);
      setForm(EMPTY_STAFF);
      toast.success("Staff member saved");
    } catch (e: any) {
      toast.error(e.message || "Failed to save");
    } finally {
      setSaving(false);
    }
  };

  const handleDelete = async (id: number) => {
    try {
      await deleteStaff(id);
      setStaff(s => s.filter(x => x.id !== id));
      toast.success("Staff member removed");
    } catch {
      toast.error("Failed to remove staff member");
    }
  };

  const inputClass = "w-full px-3 py-2 text-sm border border-border rounded-md bg-white focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-400";

  return (
    <div className="p-6 space-y-5">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-xl font-semibold text-foreground">Staff & Notify</h1>
          <p className="text-sm text-muted-foreground mt-0.5">
            Manage staff members and their notification preferences
          </p>
        </div>
        <button
          onClick={() => { setShowForm(true); setForm(EMPTY_STAFF); }}
          className="flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium text-white transition-all hover:opacity-90"
          style={{ background: "oklch(0.5 0.2 250)" }}
        >
          <Plus className="w-4 h-4" />
          Add Staff
        </button>
      </div>

      {/* Add Staff Form */}
      {showForm && (
        <div className="bg-white rounded-lg border border-blue-200 shadow-sm p-5 space-y-4"
          style={{ borderLeft: "3px solid oklch(0.5 0.2 250)" }}>
          <div className="text-sm font-semibold text-foreground">Add Staff Member</div>
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
            <div>
              <label className="block text-xs font-medium text-foreground mb-1">Full Name</label>
              <input className={inputClass} placeholder="e.g. Sarah Chen" value={form.name}
                onChange={e => setForm(f => ({ ...f, name: e.target.value }))} />
            </div>
            <div>
              <label className="block text-xs font-medium text-foreground mb-1">Email</label>
              <input className={inputClass} type="email" placeholder="sarah@mcs.com.au" value={form.email}
                onChange={e => setForm(f => ({ ...f, email: e.target.value }))} />
            </div>
            <div>
              <label className="block text-xs font-medium text-foreground mb-1">Role</label>
              <input className={inputClass} placeholder="e.g. Senior Accountant" value={form.role}
                onChange={e => setForm(f => ({ ...f, role: e.target.value }))} />
            </div>
          </div>
          <div className="flex items-center gap-6">
            <div className="flex items-center gap-2">
              <Toggle value={form.notify_approvals} onChange={v => setForm(f => ({ ...f, notify_approvals: v }))} />
              <span className="text-xs text-foreground">Approval notifications</span>
            </div>
            <div className="flex items-center gap-2">
              <Toggle value={form.notify_briefing} onChange={v => setForm(f => ({ ...f, notify_briefing: v }))} />
              <span className="text-xs text-foreground">Morning briefing</span>
            </div>
            <div className="flex items-center gap-2">
              <Toggle value={form.notify_alerts} onChange={v => setForm(f => ({ ...f, notify_alerts: v }))} />
              <span className="text-xs text-foreground">System alerts</span>
            </div>
          </div>
          <div className="flex items-center gap-3">
            <button
              onClick={handleSave}
              disabled={saving}
              className="px-4 py-2 rounded-lg text-sm font-medium text-white transition-all hover:opacity-90 disabled:opacity-50"
              style={{ background: "oklch(0.5 0.2 250)" }}
            >
              {saving ? "Saving…" : "Save"}
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

      {/* Staff List */}
      {loading ? (
        <div className="bg-white rounded-lg border border-border p-8 text-center text-sm text-muted-foreground">
          Loading staff…
        </div>
      ) : staff.length === 0 ? (
        <div className="bg-white rounded-lg border border-border p-12 text-center">
          <Users className="w-8 h-8 mx-auto mb-3 text-muted-foreground opacity-40" />
          <div className="text-sm font-medium text-foreground mb-1">No staff members yet</div>
          <div className="text-xs text-muted-foreground">Add staff to configure who receives CoWorker notifications</div>
        </div>
      ) : (
        <div className="bg-white rounded-lg border border-border shadow-sm overflow-hidden">
          <div className="px-5 py-3 border-b border-border bg-slate-50">
            <div className="grid grid-cols-[1fr_1fr_1fr_auto_auto_auto_auto] gap-4 text-xs font-semibold text-muted-foreground uppercase tracking-wide">
              <span>Name</span>
              <span>Email</span>
              <span>Role</span>
              <span className="text-center">Approvals</span>
              <span className="text-center">Briefing</span>
              <span className="text-center">Alerts</span>
              <span></span>
            </div>
          </div>
          <div className="divide-y divide-border">
            {staff.map((member) => (
              <div key={member.id}
                className="grid grid-cols-[1fr_1fr_1fr_auto_auto_auto_auto] gap-4 items-center px-5 py-4 hover:bg-slate-50 transition-colors group">
                <div className="flex items-center gap-3">
                  <div className="w-8 h-8 rounded-full flex items-center justify-center text-xs font-bold text-white flex-shrink-0"
                    style={{ background: "oklch(0.5 0.2 250)" }}>
                    {member.name.split(" ").map(n => n[0]).join("").slice(0, 2)}
                  </div>
                  <span className="text-sm font-medium text-foreground truncate">{member.name}</span>
                </div>
                <span className="text-sm text-muted-foreground truncate">{member.email}</span>
                <span className="text-sm text-muted-foreground">{member.role}</span>
                <div className="flex justify-center">
                  {member.notify_approvals
                    ? <Bell className="w-4 h-4" style={{ color: "oklch(0.5 0.2 250)" }} />
                    : <BellOff className="w-4 h-4 text-muted-foreground opacity-30" />}
                </div>
                <div className="flex justify-center">
                  {member.notify_briefing
                    ? <Bell className="w-4 h-4" style={{ color: "oklch(0.5 0.15 145)" }} />
                    : <BellOff className="w-4 h-4 text-muted-foreground opacity-30" />}
                </div>
                <div className="flex justify-center">
                  {member.notify_alerts
                    ? <Bell className="w-4 h-4" style={{ color: "oklch(0.55 0.2 35)" }} />
                    : <BellOff className="w-4 h-4 text-muted-foreground opacity-30" />}
                </div>
                <button
                  onClick={() => handleDelete(member.id!)}
                  className="w-7 h-7 rounded-md flex items-center justify-center text-muted-foreground hover:text-rose-600 hover:bg-rose-50 transition-colors opacity-0 group-hover:opacity-100"
                >
                  <Trash2 className="w-3.5 h-3.5" />
                </button>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}
