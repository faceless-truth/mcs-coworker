// Design: Refined Dark Professional — AI Chat page (specialist agents)

import { useState, useRef, useEffect } from "react";
import { Bot, Send, User, Paperclip, X, FileText, Download, Copy } from "lucide-react";
import { toast } from "sonner";
import {
  sendChatMessage,
  fetchAgents,
  uploadChatFile,
  type ChatFileRef,
} from "@/lib/api";

interface Message {
  id: number;
  role: "user" | "assistant";
  content: string;
  files?: ChatFileRef[];
}

interface Agent {
  id: string;
  name: string;
  description: string;
  icon: string;
  category: string;
  supports_files: boolean;
  file_types: string[];
  model_preference?: string;
}

// Example prompts shown as quick-start pills, keyed by agent id. When an
// agent isn't in the map we fall back to its description.
const AGENT_EXAMPLES: Record<string, string[]> = {
  plugin_builder: [
    "Create a plugin that flags ATO emails as HIGH priority and notifies Teams",
    "Build a WIP ageing report that emails the partners every Monday",
    "Detect new client emails and draft an engagement letter using XPM data",
    "Check XPM for 30-day-old invoices and draft polite follow-up emails",
  ],
  gst: [
    "Is a new residential property subject to GST?",
    "GST on commercial lease assignments",
    "Input tax credits for mixed-use assets",
  ],
  smsf: [
    "Can an SMSF lend money to a member?",
    "ECPI calculation for a fund in pension phase",
    "In-house asset rules for related party leases",
  ],
  div7a: [
    "Calculate DS from this balance sheet",
    "UPE left on trust — Div 7A implications?",
    "$150k shareholder loan, no agreement — options?",
  ],
  trusts: [
    "Streaming franked dividends to a corporate beneficiary",
    "s100A risk on a family trust distribution",
    "Vesting date extension — resettlement risk?",
  ],
  individual_tax: [
    "Teacher with rental property — deduction checklist",
    "Nurse claiming car and uniform expenses",
    "Home office deductions for an IT contractor",
  ],
  ato_portal: [
    "GIC remission for domestic violence hardship",
    "FTL penalty remission — all lodgements now current",
    "Payment arrangement for $12,000 tax debt",
  ],
  net_wealth: [
    "Extract transactions from this managed account statement",
    "Reconcile these WRAP and Annual statements",
    "Generate BGL import CSV from this data",
  ],
};

const CATEGORY_LABELS: Record<string, string> = {
  general: "General",
  tax: "Tax",
  documents: "Documents",
  compliance: "Compliance",
};

function formatBytes(n: number): string {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / 1024 / 1024).toFixed(1)} MB`;
}

export default function Chat() {
  const [agents, setAgents] = useState<Agent[]>([]);
  const [selectedAgentId, setSelectedAgentId] = useState<string>("plugin_builder");
  const [messages, setMessages] = useState<Message[]>([]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const [attachedFiles, setAttachedFiles] = useState<ChatFileRef[]>([]);
  const [uploading, setUploading] = useState(false);
  const bottomRef = useRef<HTMLDivElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const selectedAgent = agents.find(a => a.id === selectedAgentId);

  useEffect(() => {
    (async () => {
      try {
        const list = (await fetchAgents()) as Agent[];
        if (Array.isArray(list) && list.length > 0) setAgents(list);
      } catch {
        // Leave agents empty — UI falls back to plugin_builder-only behaviour.
      }
    })();
  }, []);

  // Seed a greeting whenever the selected agent changes.
  useEffect(() => {
    const agent = agents.find(a => a.id === selectedAgentId);
    if (!agent) return;
    const greeting =
      agent.id === "plugin_builder"
        ? "Hi! I'm the MCS CoWorker assistant. I can help you build new automation plugins, answer questions about your practice, or analyse client data. What would you like to build?"
        : `${agent.icon} ${agent.name} here. ${agent.description}${agent.supports_files ? " You can attach relevant source documents (PDFs, Excel, CSV, Word) using the paperclip button." : ""} How can I help?`;
    setMessages([{ id: 0, role: "assistant", content: greeting }]);
    setAttachedFiles([]);
    setInput("");
  }, [selectedAgentId, agents]);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages]);

  const handleAgentChange = (newId: string) => {
    if (newId === selectedAgentId) return;
    const agent = agents.find(a => a.id === newId);
    setSelectedAgentId(newId);
    if (agent) toast(`Switched to ${agent.name} — conversation cleared.`);
  };

  const handleFilePick = async (ev: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(ev.target.files || []);
    if (files.length === 0) return;
    if (attachedFiles.length + files.length > 5) {
      toast.error("Max 5 files per message");
      ev.target.value = "";
      return;
    }
    setUploading(true);
    try {
      for (const f of files) {
        const ref = await uploadChatFile(f);
        setAttachedFiles(prev => [...prev, ref]);
      }
    } catch (e: any) {
      toast.error("Upload failed", { description: e?.message ?? "" });
    } finally {
      setUploading(false);
      ev.target.value = "";
    }
  };

  const removeAttached = (id: string) => {
    setAttachedFiles(prev => prev.filter(f => f.id !== id));
  };

  const send = async (text?: string) => {
    const msg = text || input;
    if (!msg.trim() && attachedFiles.length === 0) return;
    setInput("");

    const currentFiles = attachedFiles;
    const userMsg: Message = {
      id: Date.now(),
      role: "user",
      content: msg,
      files: currentFiles.length ? currentFiles : undefined,
    };
    const nextMessages = [...messages, userMsg];
    setMessages(nextMessages);
    setAttachedFiles([]);
    setLoading(true);

    try {
      const history = nextMessages.map(m => ({ role: m.role, content: m.content }));
      const resp: any = await sendChatMessage(history, selectedAgentId, currentFiles);
      const assistantMsg: Message = {
        id: Date.now() + 1,
        role: "assistant",
        content: resp?.content ?? "(no response)",
      };
      setMessages(prev => [...prev, assistantMsg]);
    } catch (e: any) {
      toast.error("Chat failed", {
        description: e?.message ?? "Check your Anthropic API key in Settings.",
      });
    } finally {
      setLoading(false);
    }
  };

  const buildExportMarkdown = (): string => {
    const agentName = selectedAgent?.name || "Assistant";
    const now = new Date();
    const dateStr = now.toLocaleString("en-AU", {
      dateStyle: "long",
      timeStyle: "short",
    });

    // Skip the auto-greeting: start from the first user message.
    const firstUserIdx = messages.findIndex(m => m.role === "user");
    const convo = firstUserIdx >= 0 ? messages.slice(firstUserIdx) : [];

    let md = `# ${agentName} — Chat Export\n`;
    md += `Date: ${dateStr}\n`;
    md += `Agent: ${agentName}\n\n`;
    md += `---\n\n`;

    for (let i = 0; i < convo.length; i++) {
      const m = convo[i];
      const label = m.role === "user" ? "User" : agentName;
      md += `**${label}:** ${m.content}\n\n`;
      const isPairEnd = m.role === "assistant";
      const isLast = i === convo.length - 1;
      if (isPairEnd || isLast) {
        md += `---\n\n`;
      }
    }

    return md;
  };

  const buildExportFilename = (): string => {
    const now = new Date();
    const pad = (n: number) => String(n).padStart(2, "0");
    const ts =
      `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}` +
      `_${pad(now.getHours())}${pad(now.getMinutes())}`;
    const safeName = (selectedAgent?.name || "chat")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "");
    return `${safeName || "chat"}_${ts}.md`;
  };

  const handleExport = () => {
    try {
      const md = buildExportMarkdown();
      const blob = new Blob([md], { type: "text/markdown;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = buildExportFilename();
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      toast.success("Conversation exported");
    } catch (e: any) {
      toast.error("Export failed", { description: e?.message ?? "" });
    }
  };

  const handleCopyAll = async () => {
    try {
      await navigator.clipboard.writeText(buildExportMarkdown());
      toast.success("Conversation copied to clipboard");
    } catch (e: any) {
      toast.error("Copy failed", { description: e?.message ?? "" });
    }
  };

  const hasConversation = messages.some(m => m.role === "user");

  const renderContent = (content: string) => {
    const parts = content.split(/(```[\s\S]*?```)/g);
    return parts.map((part, i) => {
      if (part.startsWith("```")) {
        const code = part.replace(/```\w*\n?/, "").replace(/```$/, "");
        return (
          <pre
            key={i}
            className="bg-slate-900 text-slate-100 rounded-lg p-3 text-xs font-mono overflow-x-auto my-2 leading-relaxed"
          >
            {code}
          </pre>
        );
      }
      return (
        <span key={i} className="whitespace-pre-wrap">
          {part.split(/(\*\*[^*]+\*\*)/g).map((chunk, j) =>
            chunk.startsWith("**") ? <strong key={j}>{chunk.slice(2, -2)}</strong> : chunk,
          )}
        </span>
      );
    });
  };

  // Group agents by category for the selector.
  const grouped: Record<string, Agent[]> = {};
  for (const a of agents) {
    (grouped[a.category] = grouped[a.category] || []).push(a);
  }
  const categoryOrder = ["general", "tax", "documents", "compliance"];
  const orderedCats = categoryOrder.filter(c => grouped[c]?.length);

  const exampleList = AGENT_EXAMPLES[selectedAgentId] || [];
  const acceptAttr = (selectedAgent?.file_types || []).join(",");

  return (
    <div className="flex flex-col h-full">
      {/* Header */}
      <div className="flex items-center justify-between px-6 py-4 border-b border-border bg-white">
        <div>
          <h1 className="text-xl font-semibold text-foreground flex items-center gap-2">
            {selectedAgent && <span>{selectedAgent.icon}</span>}
            {selectedAgent?.name || "AI Chat"}
          </h1>
          <p className="text-sm text-muted-foreground">
            {selectedAgent?.description || "Select a specialist to begin"}
            {selectedAgent?.supports_files && (
              <span className="ml-2 inline-flex items-center gap-1 text-xs">
                <Paperclip className="w-3 h-3" /> accepts files
              </span>
            )}
          </p>
        </div>
        {hasConversation && (
          <div className="flex items-center gap-2">
            <button
              onClick={handleCopyAll}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium border border-border bg-white text-slate-700 hover:border-blue-300 hover:text-blue-600 transition-all"
              title="Copy full conversation to clipboard"
            >
              <Copy className="w-3.5 h-3.5" />
              Copy all
            </button>
            <button
              onClick={handleExport}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium border border-border bg-white text-slate-700 hover:border-blue-300 hover:text-blue-600 transition-all"
              title="Download conversation as Markdown"
            >
              <Download className="w-3.5 h-3.5" />
              Export
            </button>
          </div>
        )}
      </div>

      {/* Agent selector */}
      {agents.length > 0 && (
        <div className="px-6 py-3 border-b border-border bg-slate-50 flex flex-wrap items-center gap-3">
          {orderedCats.map(cat => (
            <div key={cat} className="flex items-center gap-2">
              <span className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
                {CATEGORY_LABELS[cat] || cat}
              </span>
              <div className="flex flex-wrap gap-1.5">
                {grouped[cat].map(a => {
                  const active = a.id === selectedAgentId;
                  return (
                    <button
                      key={a.id}
                      onClick={() => handleAgentChange(a.id)}
                      className={`flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-medium border transition-all ${
                        active
                          ? "border-blue-400 bg-blue-50 text-blue-700"
                          : "border-border bg-white hover:border-blue-300 hover:text-blue-600"
                      }`}
                      title={a.description}
                    >
                      <span>{a.icon}</span>
                      {a.name}
                    </button>
                  );
                })}
              </div>
            </div>
          ))}
        </div>
      )}

      {/* Messages */}
      <div className="flex-1 overflow-y-auto p-6 space-y-4">
        {messages.map(msg => (
          <div key={msg.id} className={`flex gap-3 ${msg.role === "user" ? "flex-row-reverse" : ""}`}>
            <div
              className={`flex-shrink-0 w-8 h-8 rounded-full flex items-center justify-center ${
                msg.role === "assistant" ? "" : "bg-slate-200"
              }`}
              style={msg.role === "assistant" ? { background: "oklch(0.5 0.2 250)" } : {}}
            >
              {msg.role === "assistant" ? (
                <Bot className="w-4 h-4 text-white" />
              ) : (
                <User className="w-4 h-4 text-slate-600" />
              )}
            </div>
            <div
              className={`max-w-[75%] rounded-xl px-4 py-3 text-sm leading-relaxed ${
                msg.role === "user"
                  ? "bg-slate-100 text-foreground rounded-tr-sm"
                  : "bg-white border border-border shadow-sm text-foreground rounded-tl-sm"
              }`}
            >
              {msg.files && msg.files.length > 0 && (
                <div className="flex flex-wrap gap-1.5 mb-2">
                  {msg.files.map(f => (
                    <span
                      key={f.id}
                      className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs bg-white border border-slate-200 text-slate-700"
                    >
                      <FileText className="w-3 h-3" />
                      {f.name}
                    </span>
                  ))}
                </div>
              )}
              {renderContent(msg.content)}
            </div>
          </div>
        ))}

        {loading && (
          <div className="flex gap-3">
            <div
              className="w-8 h-8 rounded-full flex items-center justify-center"
              style={{ background: "oklch(0.5 0.2 250)" }}
            >
              <Bot className="w-4 h-4 text-white" />
            </div>
            <div className="bg-white border border-border shadow-sm rounded-xl rounded-tl-sm px-4 py-3">
              <div className="flex gap-1.5 items-center">
                {[0, 1, 2].map(i => (
                  <div
                    key={i}
                    className="w-2 h-2 rounded-full bg-slate-300 animate-bounce"
                    style={{ animationDelay: `${i * 0.15}s` }}
                  />
                ))}
              </div>
            </div>
          </div>
        )}
        <div ref={bottomRef} />
      </div>

      {/* Examples */}
      {messages.length <= 1 && exampleList.length > 0 && (
        <div className="px-6 pb-3">
          <div className="text-xs text-muted-foreground mb-2">Try an example:</div>
          <div className="flex flex-wrap gap-2">
            {exampleList.map(ex => (
              <button
                key={ex}
                onClick={() => send(ex)}
                className="px-3 py-1.5 rounded-full text-xs font-medium border border-border bg-white hover:border-blue-300 hover:text-blue-600 transition-all"
              >
                {ex}
              </button>
            ))}
          </div>
        </div>
      )}

      {/* Attached-file chips */}
      {attachedFiles.length > 0 && (
        <div className="px-6 pb-2">
          <div className="flex flex-wrap gap-1.5">
            {attachedFiles.map(f => (
              <span
                key={f.id}
                className="inline-flex items-center gap-1.5 px-2 py-1 rounded-lg text-xs bg-blue-50 border border-blue-200 text-blue-800"
              >
                <FileText className="w-3 h-3" />
                {f.name}
                <span className="text-blue-500">· {formatBytes(f.size)}</span>
                <button
                  onClick={() => removeAttached(f.id)}
                  className="ml-1 hover:text-blue-900"
                  title="Remove"
                >
                  <X className="w-3 h-3" />
                </button>
              </span>
            ))}
          </div>
        </div>
      )}

      {/* Input */}
      <div className="px-6 pb-6 pt-2">
        <div className="flex gap-3 bg-white border border-border rounded-xl shadow-sm p-3">
          {selectedAgent?.supports_files && (
            <>
              <input
                ref={fileInputRef}
                type="file"
                multiple
                accept={acceptAttr}
                onChange={handleFilePick}
                className="hidden"
              />
              <button
                onClick={() => fileInputRef.current?.click()}
                disabled={uploading || attachedFiles.length >= 5}
                className="flex-shrink-0 w-9 h-9 rounded-lg flex items-center justify-center text-slate-500 hover:text-blue-600 hover:bg-slate-100 transition-all disabled:opacity-40"
                title={uploading ? "Uploading..." : "Attach files"}
              >
                <Paperclip className="w-4 h-4" />
              </button>
            </>
          )}
          <textarea
            value={input}
            onChange={e => setInput(e.target.value)}
            onKeyDown={e => {
              if (e.key === "Enter" && !e.shiftKey) {
                e.preventDefault();
                send();
              }
            }}
            placeholder={
              selectedAgent?.id === "plugin_builder"
                ? "Describe a plugin to build, or ask about your practice data..."
                : `Ask ${selectedAgent?.name || "the assistant"}...`
            }
            rows={2}
            className="flex-1 text-sm resize-none focus:outline-none text-foreground placeholder:text-muted-foreground"
          />
          <button
            onClick={() => send()}
            disabled={(!input.trim() && attachedFiles.length === 0) || loading}
            className="flex-shrink-0 w-9 h-9 rounded-lg flex items-center justify-center text-white transition-all disabled:opacity-40"
            style={{ background: "oklch(0.5 0.2 250)" }}
          >
            <Send className="w-4 h-4" />
          </button>
        </div>
        <div className="text-xs text-muted-foreground mt-2 text-center">
          Shift+Enter for new line
        </div>
      </div>
    </div>
  );
}
