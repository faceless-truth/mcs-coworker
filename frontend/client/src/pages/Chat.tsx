// Design: Refined Dark Professional — AI Chat page (specialist agents)

import { useState, useRef, useEffect, useMemo } from "react";
import { Bot, Send, User, Paperclip, X, FileText, Download, Copy, Check, FolderUp } from "lucide-react";
import { toast } from "sonner";
import {
  sendChatMessage,
  fetchAgents,
  fetchClientNames,
  fetchSettings,
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
  general: [
    "Help me draft an email to a client",
    "Summarise this document",
    "What are the key EOFY 2026 dates?",
    "Quick question about a client situation",
  ],
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
  tax_structure: [
    "New client starting a consulting business — best structure?",
    "Restructure from sole trader to trust — client earns $180k",
    "IT contractor, single client, $200k profit — PSI implications?",
    "Family with 3 adult beneficiaries, business profit $300k",
  ],
  payroll: [
    "New receptionist in a law firm — what award and rate?",
    "Casual kitchen hand in a cafe — pay rates and penalties",
    "Is this contractor arrangement legitimate?",
    "Qualified carpenter, full-time, Melbourne — full pay rate card",
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
  const [selectedAgentId, setSelectedAgentId] = useState<string>("general");
  const [messages, setMessages] = useState<Message[]>([]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const [attachedFiles, setAttachedFiles] = useState<ChatFileRef[]>([]);
  const [uploading, setUploading] = useState(false);
  const [uploadingFiles, setUploadingFiles] = useState<{ name: string; size: number }[]>([]);
  const [dragActive, setDragActive] = useState(false);
  const [exportMenuOpen, setExportMenuOpen] = useState(false);
  const [exportingType, setExportingType] = useState<"transcript" | "summary" | "recommendation" | "sharepoint" | null>(null);
  const [sharepointConfigured, setSharepointConfigured] = useState(false);
  const [clientName, setClientName] = useState("");
  const [entityName, setEntityName] = useState("");
  const [clientNamesList, setClientNamesList] = useState<string[]>([]);
  const [showClientSuggestions, setShowClientSuggestions] = useState(false);
  const bottomRef = useRef<HTMLDivElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const selectedAgent = agents.find(a => a.id === selectedAgentId);
  const isSpecialist =
    !!selectedAgent &&
    selectedAgent.id !== "plugin_builder" &&
    selectedAgent.id !== "general";

  const clientSuggestions = useMemo(() => {
    const q = clientName.trim().toLowerCase();
    if (!q) return [];
    return clientNamesList
      .filter(n => n.toLowerCase().includes(q))
      .slice(0, 8);
  }, [clientName, clientNamesList]);

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

  // Cache the known client list once for autocomplete.
  useEffect(() => {
    (async () => {
      try {
        const names = await fetchClientNames();
        setClientNamesList(names);
      } catch {
        // Autocomplete falls back to empty — typing a new name still works.
      }
    })();
  }, []);

  // SharePoint config is now hardcoded for MC&S, so always treat it as configured.
  useEffect(() => {
    setSharepointConfigured(true);
  }, []);

  // Seed a greeting whenever the selected agent changes.
  useEffect(() => {
    const agent = agents.find(a => a.id === selectedAgentId);
    if (!agent) return;
    const isSpec =
      agent.id !== "plugin_builder" && agent.id !== "general";
    let greeting: string;
    if (agent.id === "plugin_builder") {
      greeting = "Hi! I'm the MCS CoWorker assistant. I can help you build new automation plugins, answer questions about your practice, or analyse client data. What would you like to build?";
    } else if (isSpec && clientName.trim()) {
      greeting = `${agent.icon} ${agent.name} here. Working on ${clientName.trim()}. How can I help?`;
    } else if (isSpec) {
      greeting = `${agent.icon} ${agent.name} here. ${agent.description}${agent.supports_files ? " You can attach relevant source documents (PDFs, Excel, CSV, Word) using the paperclip button." : ""} Is this query for a specific client?`;
    } else {
      greeting = `${agent.icon} ${agent.name} here. ${agent.description}${agent.supports_files ? " You can attach relevant source documents (PDFs, Excel, CSV, Word) using the paperclip button." : ""} How can I help?`;
    }
    setMessages([{ id: 0, role: "assistant", content: greeting }]);
    setAttachedFiles([]);
    setInput("");
    // Clear client context when switching agents.
    setClientName("");
    setEntityName("");
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

  const uploadFiles = async (files: File[]) => {
    if (files.length === 0) return;
    if (attachedFiles.length + files.length > 5) {
      toast.error("Max 5 files per message");
      return;
    }
    setUploading(true);
    setUploadingFiles(files.map(f => ({ name: f.name, size: f.size })));
    try {
      for (const f of files) {
        try {
          const ref = await uploadChatFile(f);
          setAttachedFiles(prev => [...prev, ref]);
        } catch (e: any) {
          toast.error(`Upload failed: ${f.name}`, { description: e?.message ?? "" });
        } finally {
          setUploadingFiles(prev => prev.filter(p => p.name !== f.name));
        }
      }
    } finally {
      setUploading(false);
      setUploadingFiles([]);
    }
  };

  const handleFilePick = async (ev: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(ev.target.files || []);
    await uploadFiles(files);
    ev.target.value = "";
  };

  const handleDragOver = (e: React.DragEvent) => {
    if (!selectedAgent?.supports_files) return;
    e.preventDefault();
    e.stopPropagation();
    if (!dragActive) setDragActive(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    // Only deactivate when leaving the outer container.
    if (e.currentTarget === e.target) setDragActive(false);
  };

  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    if (!selectedAgent?.supports_files) return;
    const files = Array.from(e.dataTransfer.files || []);
    if (files.length === 0) return;
    // Filter by accepted extensions.
    const accepted = (selectedAgent.file_types || []).map(t => t.toLowerCase());
    const valid = accepted.length === 0
      ? files
      : files.filter(f => {
          const dot = f.name.lastIndexOf(".");
          const ext = dot >= 0 ? f.name.slice(dot).toLowerCase() : "";
          return accepted.includes(ext);
        });
    if (valid.length < files.length) {
      toast.error(`${files.length - valid.length} file(s) skipped — unsupported type`);
    }
    await uploadFiles(valid);
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
      const resp: any = await sendChatMessage(
        history,
        selectedAgentId,
        currentFiles,
        clientName.trim() || null,
        entityName.trim() || null,
      );
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

  const handleExport = async (
    exportType: "transcript" | "summary" | "recommendation" = "transcript",
  ) => {
    const agentName = selectedAgent?.name || "Assistant";
    const firstUserIdx = messages.findIndex(m => m.role === "user");
    const convo = firstUserIdx >= 0 ? messages.slice(firstUserIdx) : [];
    if (convo.length === 0) {
      toast.error("Nothing to export yet");
      return;
    }
    setExportingType(exportType);
    if (exportType === "summary") {
      toast("Generating summary…", {
        description: "Claude is condensing the conversation — this takes a few seconds.",
      });
    }
    try {
      const res = await fetch("http://127.0.0.1:7842/api/chat/export", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          agent_name: agentName,
          agent_id: selectedAgentId,
          messages: convo.map(m => ({ role: m.role, content: m.content })),
          client_name: clientName.trim() || null,
          entity_name: entityName.trim() || null,
          export_type: exportType,
        }),
      });
      const body = await res.json().catch(() => ({}));
      if (!res.ok || !body.ok) {
        throw new Error(body.error || `HTTP ${res.status}`);
      }
      toast.success(`Saved to Downloads: ${body.filename}`);
    } catch (e: any) {
      toast.error("Export failed", { description: e?.message ?? "" });
    } finally {
      setExportingType(null);
      setExportMenuOpen(false);
    }
  };

  const handleSharepointExport = async () => {
    const agentName = selectedAgent?.name || "Assistant";
    const firstUserIdx = messages.findIndex(m => m.role === "user");
    const convo = firstUserIdx >= 0 ? messages.slice(firstUserIdx) : [];
    if (convo.length === 0) {
      toast.error("Nothing to export yet");
      return;
    }
    if (!clientName.trim()) {
      toast.error("Enter a client name to save to SharePoint");
      return;
    }
    if (!sharepointConfigured) {
      toast.error("Configure SharePoint in Settings first");
      return;
    }
    // Pick the most useful export type for the file record:
    //   recommendation if the specialist produced one, otherwise transcript.
    const exportType = hasStructuredRecommendation ? "recommendation" : "transcript";
    setExportingType("sharepoint");
    try {
      const res = await fetch("http://127.0.0.1:7842/api/chat/export/sharepoint", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          agent_name: agentName,
          agent_id: selectedAgentId,
          messages: convo.map(m => ({ role: m.role, content: m.content })),
          client_name: clientName.trim(),
          entity_name: entityName.trim() || null,
          export_type: exportType,
        }),
      });
      const body = await res.json().catch(() => ({}));
      if (!res.ok || !body.ok) {
        throw new Error(body.error || `HTTP ${res.status}`);
      }
      toast.success("Saved to SharePoint", {
        description: body.filename,
        action: body.url
          ? { label: "Open", onClick: () => window.open(body.url, "_blank") }
          : undefined,
      });
    } catch (e: any) {
      toast.error("SharePoint save failed", { description: e?.message ?? "" });
    } finally {
      setExportingType(null);
      setExportMenuOpen(false);
    }
  };

  // Scan every assistant message for structured headings — the
  // recommendation may have been produced before a follow-up turn,
  // so checking only the last message would falsely grey the export.
  const hasStructuredRecommendation = messages.some(m => {
    if (m.role !== "assistant" || typeof m.content !== "string") return false;
    const headingCount =
      (m.content.match(/\n#{1,3}\s/g) || []).length +
      (/^#{1,3}\s/.test(m.content) ? 1 : 0);
    return headingCount >= 2;
  });

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
  const categoryOrder = ["general", "tax", "compliance", "documents"];
  const orderedCats = categoryOrder.filter(c => grouped[c]?.length);

  const exampleList = AGENT_EXAMPLES[selectedAgentId] || [];
  const acceptAttr = (selectedAgent?.file_types || []).join(",");

  return (
    <div
      className="flex flex-col h-full relative"
      onDragOver={handleDragOver}
      onDragEnter={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
    >
      {dragActive && selectedAgent?.supports_files && (
        <div
          className="absolute inset-0 z-50 flex items-center justify-center pointer-events-none"
          style={{ background: "rgba(59, 130, 246, 0.08)" }}
        >
          <div className="bg-white border-2 border-dashed border-blue-400 rounded-2xl px-10 py-8 shadow-lg text-center">
            <Paperclip className="w-8 h-8 text-blue-500 mx-auto mb-2" />
            <div className="text-base font-semibold text-blue-700">
              Drop files here
            </div>
            <div className="text-xs text-blue-600/70 mt-1">
              PDFs, Excel, Word, CSV, images — up to 5 per message
            </div>
          </div>
        </div>
      )}

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
            <div className="relative">
              <button
                onClick={() => setExportMenuOpen(o => !o)}
                onBlur={() => setTimeout(() => setExportMenuOpen(false), 150)}
                disabled={exportingType !== null}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium border border-border bg-white text-slate-700 hover:border-blue-300 hover:text-blue-600 transition-all disabled:opacity-60"
                title="Download as Word document"
              >
                <Download className="w-3.5 h-3.5" />
                {exportingType === "summary" ? "Generating summary…"
                  : exportingType === "sharepoint" ? "Saving to SharePoint…"
                  : exportingType ? "Exporting…"
                  : "Export"}
                <span className="text-slate-400">▾</span>
              </button>
              {exportMenuOpen && (
                <div className="absolute right-0 mt-1 z-30 w-64 bg-white border border-border rounded-lg shadow-lg overflow-hidden">
                  <button
                    type="button"
                    onMouseDown={e => { e.preventDefault(); handleExport("transcript"); }}
                    className="w-full text-left px-3 py-2 text-xs hover:bg-blue-50 border-b border-border"
                  >
                    <div className="font-semibold text-slate-700">Download Transcript (.docx)</div>
                    <div className="text-slate-500">Full conversation as a Word doc</div>
                  </button>
                  <button
                    type="button"
                    onMouseDown={e => { e.preventDefault(); handleExport("summary"); }}
                    className="w-full text-left px-3 py-2 text-xs hover:bg-blue-50 border-b border-border"
                  >
                    <div className="font-semibold text-slate-700">Download Summary (.docx)</div>
                    <div className="text-slate-500">Claude-generated summary for the client file</div>
                  </button>
                  <button
                    type="button"
                    onMouseDown={e => {
                      if (!hasStructuredRecommendation) return;
                      e.preventDefault();
                      handleExport("recommendation");
                    }}
                    disabled={!hasStructuredRecommendation}
                    title={hasStructuredRecommendation
                      ? undefined
                      : "No structured recommendation to export — ask the specialist to produce one first."}
                    className="w-full text-left px-3 py-2 text-xs hover:bg-blue-50 border-b border-border disabled:opacity-50 disabled:hover:bg-white disabled:cursor-not-allowed"
                  >
                    <div className="font-semibold text-slate-700">Download Recommendation (.docx)</div>
                    <div className="text-slate-500">
                      {hasStructuredRecommendation
                        ? "Specialist's structured output as a formatted document"
                        : "Available once the specialist produces a structured doc"}
                    </div>
                  </button>
                  {(() => {
                    const sharepointDisabled = !clientName.trim() || !sharepointConfigured;
                    const sharepointTooltip = !clientName.trim()
                      ? "Enter a client name to save to SharePoint"
                      : !sharepointConfigured
                      ? "Configure SharePoint in Settings first"
                      : undefined;
                    return (
                      <button
                        type="button"
                        onMouseDown={e => {
                          if (sharepointDisabled) return;
                          e.preventDefault();
                          handleSharepointExport();
                        }}
                        disabled={sharepointDisabled}
                        title={sharepointTooltip}
                        className="w-full text-left px-3 py-2 text-xs hover:bg-blue-50 disabled:opacity-50 disabled:hover:bg-white disabled:cursor-not-allowed"
                      >
                        <div className="font-semibold text-slate-700 flex items-center gap-1.5">
                          <FolderUp className="w-3.5 h-3.5" />
                          Save to Client Folder (SharePoint)
                        </div>
                        <div className="text-slate-500">
                          {sharepointDisabled
                            ? sharepointTooltip
                            : `Files into /${clientName.trim()}/CoWorker Exports/`}
                        </div>
                      </button>
                    );
                  })()}
                </div>
              )}
            </div>
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

      {/* Client context bar (specialist agents only) */}
      {isSpecialist && (
        <div className="px-6 py-2.5 border-b border-border bg-amber-50/40 flex flex-wrap items-center gap-3 text-sm">
          <span className="inline-flex items-center gap-1.5 text-xs font-semibold uppercase tracking-wide text-muted-foreground">
            📋 Client
          </span>

          <div className="relative flex-1 min-w-[220px] max-w-md">
            <input
              type="text"
              value={clientName}
              onChange={e => {
                setClientName(e.target.value);
                setShowClientSuggestions(true);
              }}
              onFocus={() => setShowClientSuggestions(true)}
              onBlur={() => setTimeout(() => setShowClientSuggestions(false), 150)}
              placeholder="Who is this query for? (Surname, First Name)"
              className="w-full px-3 py-1.5 pr-8 rounded-lg border border-border bg-white text-sm focus:outline-none focus:border-blue-400"
            />
            {clientName && (
              <button
                type="button"
                onClick={() => setClientName("")}
                className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-700"
                title="Clear"
              >
                <X className="w-3.5 h-3.5" />
              </button>
            )}
            {showClientSuggestions && clientSuggestions.length > 0 && (
              <div className="absolute z-10 left-0 right-0 mt-1 bg-white border border-border rounded-lg shadow-lg max-h-56 overflow-y-auto">
                {clientSuggestions.map(name => (
                  <button
                    type="button"
                    key={name}
                    onMouseDown={e => {
                      // onMouseDown so it fires before the input's onBlur.
                      e.preventDefault();
                      setClientName(name);
                      setShowClientSuggestions(false);
                    }}
                    className="block w-full text-left px-3 py-1.5 text-sm hover:bg-blue-50"
                  >
                    {name}
                  </button>
                ))}
              </div>
            )}
          </div>

          {clientName.trim() && (
            <span className="inline-flex items-center gap-1 text-xs text-emerald-700">
              <Check className="w-3.5 h-3.5" />
              {clientName.trim()}
            </span>
          )}

          <span className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
            Entity
          </span>
          <input
            type="text"
            value={entityName}
            onChange={e => setEntityName(e.target.value)}
            placeholder="(optional)"
            className="flex-1 min-w-[180px] max-w-xs px-3 py-1.5 rounded-lg border border-border bg-white text-sm focus:outline-none focus:border-blue-400"
          />
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

      {/* Attached-file chips (uploaded + in-progress) */}
      {(attachedFiles.length > 0 || uploadingFiles.length > 0) && (
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
            {uploadingFiles.map(f => (
              <span
                key={`up-${f.name}`}
                className="inline-flex items-center gap-1.5 px-2 py-1 rounded-lg text-xs bg-slate-50 border border-slate-200 text-slate-600"
              >
                <span
                  className="w-3 h-3 rounded-full border-2 border-slate-300 border-t-blue-500 animate-spin"
                  aria-hidden
                />
                {f.name}
                <span className="text-slate-400">· {formatBytes(f.size)}</span>
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
          {selectedAgent?.supports_files && acceptAttr && (
            <>
              {" · "}Accepts {acceptAttr}{" · "}drop files here or use the paperclip
            </>
          )}
        </div>
      </div>
    </div>
  );
}
