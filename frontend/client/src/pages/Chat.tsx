// Design: Refined Dark Professional — AI Chat page (Tier 2 plugin builder)

import { useState, useRef, useEffect } from "react";
import { Bot, Send, Sparkles, User, Zap } from "lucide-react";

interface Message {
  id: number;
  role: "user" | "assistant";
  content: string;
  tier?: 1 | 2;
}

const examples = [
  { label: "Email triage rule", prompt: "Create a plugin that flags emails from the ATO as HIGH priority and notifies the team via Teams", tier: 1 },
  { label: "WIP report", prompt: "Build a plugin that queries XPM every Monday morning and emails a WIP ageing report to the partners", tier: 2 },
  { label: "Client onboarding", prompt: "Create a plugin that detects new client emails and automatically drafts an engagement letter using XPM data", tier: 2 },
  { label: "Debtor reminder", prompt: "Build a plugin that checks XPM for invoices over 30 days and drafts polite follow-up emails", tier: 2 },
];

const mockResponses: Record<string, string> = {
  default: `I'll build that plugin for you. Let me analyse the requirements...\n\n**Tier 2 Plugin: XPM WIP Report**\n\nHere's what I'll create:\n\n\`\`\`python\nclass WIPReportPlugin(AgentPlugin):\n    PLUGIN_ID = "plugin_wip_report"\n    NAME = "WIP Report"\n    SCHEDULE = Schedule.weekly_on(0, 8)  # Monday 8am\n    \n    def run(self, context: PluginContext) -> PluginResult:\n        # 1. Fetch jobs from XPM\n        if not context.gateway.xpm.is_configured:\n            return PluginResult(success=False, message="XPM not configured")\n        \n        jobs = context.gateway.xpm.list_jobs(status="active")\n        stale = [j for j in jobs if j["days_since_update"] > 14]\n        \n        # 2. Use Sonnet to draft summary\n        summary = context.claude_reason.messages.create(\n            model=self.get_claude_model_reasoning(),\n            messages=[{"role": "user", "content": f"Summarise these stale jobs: {stale}"}]\n        )\n        \n        # 3. Send via Teams\n        context.gateway.teams.send_alert(\n            title="Weekly WIP Report",\n            body=summary.content[0].text\n        )\n        \n        return PluginResult(success=True, message=f"Reported {len(stale)} stale jobs")\n\`\`\`\n\n✅ Plugin created and saved to \`plugins/plugin_wip_report.py\`\n✅ Will run automatically next Monday at 8:00am\n✅ Uses Sonnet for intelligent summarisation\n✅ Sends to Teams via configured webhook`,
};

export default function Chat() {
  const [messages, setMessages] = useState<Message[]>([
    {
      id: 0,
      role: "assistant",
      content: "Hi! I'm the MCS CoWorker AI assistant. I can help you build new automation plugins, answer questions about your practice, or analyse client data.\n\n**Tier 1** — I can create plugins from templates (email rules, auto-replies, document processing).\n\n**Tier 2** — I can write full custom Python plugins using XPM, FuseSign, Teams, vector memory, and both Claude models.\n\nWhat would you like to build?",
    },
  ]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const bottomRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages]);

  const send = async (text?: string) => {
    const msg = text || input;
    if (!msg.trim()) return;
    setInput("");

    const userMsg: Message = { id: Date.now(), role: "user", content: msg };
    setMessages(prev => [...prev, userMsg]);
    setLoading(true);

    // Detect tier
    const tier2Keywords = ["xpm", "fusesign", "teams", "memory", "workflow", "report", "wip", "debtor", "engagement", "onboard"];
    const isTier2 = tier2Keywords.some(k => msg.toLowerCase().includes(k));

    await new Promise(r => setTimeout(r, 1800));

    const response = mockResponses.default;
    const assistantMsg: Message = {
      id: Date.now() + 1,
      role: "assistant",
      content: response,
      tier: isTier2 ? 2 : 1,
    };
    setMessages(prev => [...prev, assistantMsg]);
    setLoading(false);
  };

  const renderContent = (content: string) => {
    const parts = content.split(/(```[\s\S]*?```)/g);
    return parts.map((part, i) => {
      if (part.startsWith("```")) {
        const code = part.replace(/```\w*\n?/, "").replace(/```$/, "");
        return (
          <pre key={i} className="bg-slate-900 text-slate-100 rounded-lg p-3 text-xs font-mono overflow-x-auto my-2 leading-relaxed">
            {code}
          </pre>
        );
      }
      return (
        <span key={i} className="whitespace-pre-wrap">
          {part.split(/(\*\*[^*]+\*\*)/g).map((chunk, j) =>
            chunk.startsWith("**") ? (
              <strong key={j}>{chunk.slice(2, -2)}</strong>
            ) : chunk
          )}
        </span>
      );
    });
  };

  return (
    <div className="flex flex-col h-full">
      {/* Header */}
      <div className="flex items-center justify-between px-6 py-4 border-b border-border bg-white">
        <div>
          <h1 className="text-xl font-semibold text-foreground">AI Chat</h1>
          <p className="text-sm text-muted-foreground">Tier 2 Plugin Builder · Powered by Claude Sonnet</p>
        </div>
        <div className="flex items-center gap-2 px-3 py-1.5 rounded-full text-xs font-medium" style={{ background: "oklch(0.94 0.05 250)", color: "oklch(0.4 0.15 250)" }}>
          <Sparkles className="w-3.5 h-3.5" />
          Tier 2 Active
        </div>
      </div>

      {/* Messages */}
      <div className="flex-1 overflow-y-auto p-6 space-y-4">
        {messages.map(msg => (
          <div key={msg.id} className={`flex gap-3 ${msg.role === "user" ? "flex-row-reverse" : ""}`}>
            <div className={`flex-shrink-0 w-8 h-8 rounded-full flex items-center justify-center ${
              msg.role === "assistant" ? "" : "bg-slate-200"
            }`} style={msg.role === "assistant" ? { background: "oklch(0.5 0.2 250)" } : {}}>
              {msg.role === "assistant" ? <Bot className="w-4 h-4 text-white" /> : <User className="w-4 h-4 text-slate-600" />}
            </div>
            <div className={`max-w-[75%] rounded-xl px-4 py-3 text-sm leading-relaxed ${
              msg.role === "user"
                ? "bg-slate-100 text-foreground rounded-tr-sm"
                : "bg-white border border-border shadow-sm text-foreground rounded-tl-sm"
            }`}>
              {msg.tier === 2 && (
                <div className="flex items-center gap-1.5 mb-2 text-xs font-medium" style={{ color: "oklch(0.5 0.2 250)" }}>
                  <Zap className="w-3 h-3" />
                  Tier 2 — using Claude Sonnet
                </div>
              )}
              {renderContent(msg.content)}
            </div>
          </div>
        ))}

        {loading && (
          <div className="flex gap-3">
            <div className="w-8 h-8 rounded-full flex items-center justify-center" style={{ background: "oklch(0.5 0.2 250)" }}>
              <Bot className="w-4 h-4 text-white" />
            </div>
            <div className="bg-white border border-border shadow-sm rounded-xl rounded-tl-sm px-4 py-3">
              <div className="flex gap-1.5 items-center">
                {[0, 1, 2].map(i => (
                  <div key={i} className="w-2 h-2 rounded-full bg-slate-300 animate-bounce" style={{ animationDelay: `${i * 0.15}s` }} />
                ))}
              </div>
            </div>
          </div>
        )}
        <div ref={bottomRef} />
      </div>

      {/* Examples */}
      {messages.length <= 1 && (
        <div className="px-6 pb-3">
          <div className="text-xs text-muted-foreground mb-2">Try an example:</div>
          <div className="flex flex-wrap gap-2">
            {examples.map(ex => (
              <button
                key={ex.label}
                onClick={() => send(ex.prompt)}
                className="flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-medium border border-border bg-white hover:border-blue-300 hover:text-blue-600 transition-all"
              >
                {ex.tier === 2 && <Zap className="w-3 h-3" />}
                {ex.label}
              </button>
            ))}
          </div>
        </div>
      )}

      {/* Input */}
      <div className="px-6 pb-6 pt-2">
        <div className="flex gap-3 bg-white border border-border rounded-xl shadow-sm p-3">
          <textarea
            value={input}
            onChange={e => setInput(e.target.value)}
            onKeyDown={e => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); send(); } }}
            placeholder="Describe a plugin to build, or ask about your practice data..."
            rows={2}
            className="flex-1 text-sm resize-none focus:outline-none text-foreground placeholder:text-muted-foreground"
          />
          <button
            onClick={() => send()}
            disabled={!input.trim() || loading}
            className="flex-shrink-0 w-9 h-9 rounded-lg flex items-center justify-center text-white transition-all disabled:opacity-40"
            style={{ background: "oklch(0.5 0.2 250)" }}
          >
            <Send className="w-4 h-4" />
          </button>
        </div>
        <div className="text-xs text-muted-foreground mt-2 text-center">
          Shift+Enter for new line · Tier 2 keywords (xpm, teams, memory, wip) auto-switch to Sonnet
        </div>
      </div>
    </div>
  );
}
