// Design: Refined Dark Professional — dark slate sidebar with cobalt blue accents
// Fixed 240px sidebar with logo, nav groups, and bottom system status strip
// System status strip fetches LIVE data from /api/system/status — NO mock data

import { useState, useEffect } from "react";
import {
  Activity,
  Bell,
  Brain,
  CheckSquare,
  ChevronRight,
  Cpu,
  FileText,
  Filter,
  Gauge,
  LayoutDashboard,
  MessageSquare,
  Plug,
  Settings,
  Users,
  Zap,
} from "lucide-react";

const API_BASE = "http://127.0.0.1:7842";

interface SystemStatus {
  version?: string;
  heartbeatTick?: number;
  heartbeat_tick?: number;
  costToday?: string;
  cost_today?: string;
  fastModel?: string;
  fast_model?: string;
  reasoningModel?: string;
  reasoning_model?: string;
  memoryRecords?: number;
  memory_records?: number;
}

function useSystemStatus() {
  const [status, setStatus] = useState<SystemStatus | null>(null);

  useEffect(() => {
    const fetchStatus = async () => {
      try {
        const res = await fetch(`${API_BASE}/api/system/status`);
        if (res.ok) {
          const data = await res.json();
          setStatus(data);
        }
      } catch (_) {
        // backend not ready yet — keep null
      }
    };
    fetchStatus();
    const interval = setInterval(fetchStatus, 30000);
    return () => clearInterval(interval);
  }, []);

  return status;
}

interface SidebarProps {
  activePage: string;
  onNavigate: (page: string) => void;
  pendingApprovals: number;
}

const navGroups = [
  {
    label: "Overview",
    items: [
      { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
      { id: "activity", label: "Activity Log", icon: Activity },
      { id: "approvals", label: "Approvals", icon: CheckSquare, badge: true },
    ],
  },
  {
    label: "Automation",
    items: [
      { id: "plugins", label: "Plugins", icon: Plug },
      { id: "rules", label: "Email Rules", icon: Filter },
      { id: "chat", label: "AI Chat", icon: MessageSquare },
    ],
  },
  {
    label: "Intelligence",
    items: [
      { id: "memory", label: "Memory", icon: Brain },
      { id: "kpi", label: "KPI Monitor", icon: Gauge },
      { id: "usage", label: "AI Usage", icon: Zap },
    ],
  },
  {
    label: "System",
    items: [
      { id: "staff", label: "Staff & Notify", icon: Users },
      { id: "settings", label: "Settings", icon: Settings },
    ],
  },
];

export default function Sidebar({ activePage, onNavigate, pendingApprovals }: SidebarProps) {
  const sys = useSystemStatus();

  const version = sys?.version ?? "—";
  const heartbeatTick = sys?.heartbeatTick ?? sys?.heartbeat_tick;
  const costToday = sys?.costToday ?? sys?.cost_today ?? "—";
  const fastModel = sys?.fastModel ?? sys?.fast_model ?? "—";
  const reasoningModel = sys?.reasoningModel ?? sys?.reasoning_model ?? "—";
  const memoryRecords = sys?.memoryRecords ?? sys?.memory_records;

  return (
    <aside
      className="flex flex-col h-full"
      style={{
        width: 240,
        minWidth: 240,
        background: "oklch(0.175 0.02 245)",
        borderRight: "1px solid oklch(0.25 0.02 245)",
      }}
    >
      {/* Logo */}
      <div className="flex items-center gap-3 px-4 py-5" style={{ borderBottom: "1px solid oklch(0.25 0.02 245)" }}>
        <img
          src="https://d2xsxph8kpxj0f.cloudfront.net/310519663335455300/LByvWCytss4VKinErwUWjP/Outlook-ezrf1yby_2674f4e9.png"
          alt="MC&S"
          className="w-8 h-8 rounded-lg object-cover flex-shrink-0"
        />
        <div>
          <div className="text-sm font-semibold" style={{ color: "oklch(0.92 0.005 240)" }}>
            MCS CoWorker
          </div>
          <div className="text-xs" style={{ color: "oklch(0.5 0.01 240)" }}>
            {version !== "—" ? `v${version}` : "—"}
          </div>
        </div>
      </div>

      {/* Nav */}
      <nav className="flex-1 overflow-y-auto px-2 py-3 space-y-5">
        {navGroups.map((group) => (
          <div key={group.label}>
            <div
              className="px-3 mb-1 text-xs font-semibold uppercase tracking-wider"
              style={{ color: "oklch(0.42 0.01 240)" }}
            >
              {group.label}
            </div>
            <div className="space-y-0.5">
              {group.items.map((item) => {
                const Icon = item.icon;
                const isActive = activePage === item.id;
                const hasBadge = item.badge && pendingApprovals > 0;
                return (
                  <button
                    key={item.id}
                    onClick={() => onNavigate(item.id)}
                    className={`nav-item w-full text-left ${isActive ? "active" : ""}`}
                  >
                    <Icon className="w-4 h-4 flex-shrink-0" />
                    <span className="flex-1">{item.label}</span>
                    {hasBadge && (
                      <span
                        className="flex items-center justify-center w-5 h-5 rounded-full text-xs font-bold"
                        style={{ background: "oklch(0.65 0.22 35)", color: "white" }}
                      >
                        {pendingApprovals}
                      </span>
                    )}
                    {isActive && !hasBadge && (
                      <ChevronRight className="w-3 h-3 opacity-50" />
                    )}
                  </button>
                );
              })}
            </div>
          </div>
        ))}
      </nav>

      {/* System status strip — live data from /api/system/status */}
      <div
        className="px-3 py-3 space-y-2"
        style={{ borderTop: "1px solid oklch(0.25 0.02 245)" }}
      >
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2">
            <span className="status-dot running" />
            <span className="text-xs" style={{ color: "oklch(0.55 0.01 240)" }}>
              {heartbeatTick != null ? `Heartbeat #${heartbeatTick.toLocaleString()}` : "Connecting..."}
            </span>
          </div>
          <span className="text-xs font-mono" style={{ color: "oklch(0.55 0.01 240)" }}>
            {costToday}
          </span>
        </div>
        <div className="flex items-center gap-2">
          <Cpu className="w-3 h-3" style={{ color: "oklch(0.42 0.01 240)" }} />
          <span className="text-xs truncate" style={{ color: "oklch(0.42 0.01 240)" }}>
            {fastModel !== "—" ? `${fastModel} · ${reasoningModel}` : "Loading models..."}
          </span>
        </div>
        <div className="flex items-center gap-2">
          <FileText className="w-3 h-3" style={{ color: "oklch(0.42 0.01 240)" }} />
          <span className="text-xs" style={{ color: "oklch(0.42 0.01 240)" }}>
            {memoryRecords != null ? `${memoryRecords.toLocaleString()} memory records` : "Loading..."}
          </span>
        </div>
      </div>
    </aside>
  );
}
