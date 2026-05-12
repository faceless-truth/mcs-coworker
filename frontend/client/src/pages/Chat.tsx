// Design: Refined Dark Professional — AI Chat page (specialist agents)

import { useState, useRef, useEffect, useMemo } from "react";
import { Bot, Send, User, Paperclip, X, FileText, Download, Copy, Check, Mail } from "lucide-react";
import { toast } from "sonner";
import {
  sendChatMessage,
  fetchAgents,
  fetchClientNames,
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
  senior_tax_advisor: [
    "Corporate beneficiary UPE from FY24 trust resolution, $250k unpaid — Div 7A exposure post-Bendel?",
    "Family trust holding shares for 15 years — small business CGT concession eligibility on sale",
    "s 100A reimbursement agreement risk on adult-child distribution funding parents' lifestyle",
    "Pre-CGT business asset sale into a new corporate structure — Pt IVA exposure?",
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
  ato_correspondence: [
    "Process today's ATO correspondence batch",
    "I've uploaded the daily ATO folder — please classify and draft emails",
    "Sort these ATO documents by client and draft responses",
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

const MAX_FILES_PER_MESSAGE = 50;

interface EmailBlock {
  blockKey: string;       // stable key per (msgId, blockIndex)
  blockIndex: number;     // ordinal within the message
  rangeStart: number;     // char offset of block start in the source content
  rangeEnd: number;       // char offset of block end (exclusive)
  title: string;          // e.g. "Korkie, Gordon"
  to: string;             // raw To: value (may be a placeholder)
  subject: string;
  attachments: string[];  // filenames the assistant referenced
  body: string;           // cleaned body text
  flagged: boolean;       // true if this block is a "do not send" flag
}

const EMAIL_HEADER_RE = /^###\s*Email(?:\s+\d+(?:\s+of\s+\d+)?)?\s*:?\s*(.+)$/m;

function parseEmailBlocks(content: string, msgId: number | string): EmailBlock[] {
  if (!content || typeof content !== "string") return [];
  const blocks: EmailBlock[] = [];
  // Split by Email N headings while keeping their offsets.
  const lines = content.split("\n");
  const headers: { idx: number; title: string; charOffset: number }[] = [];
  let charOffset = 0;
  for (let i = 0; i < lines.length; i++) {
    const m = lines[i].match(EMAIL_HEADER_RE);
    if (m) {
      headers.push({ idx: i, title: m[1].trim(), charOffset });
    }
    charOffset += lines[i].length + 1; // +1 for the \n
  }
  if (headers.length === 0) return [];

  for (let h = 0; h < headers.length; h++) {
    const start = headers[h].idx;
    const end = h + 1 < headers.length ? headers[h + 1].idx : lines.length;
    const segLines = lines.slice(start, end);
    const seg = segLines.join("\n");
    const segStart = headers[h].charOffset;
    const segEnd = segStart + seg.length;

    const flagged = /flagged for review/i.test(seg) || /do not send/i.test(seg);

    const toMatch = seg.match(/\*\*To:\*\*\s*([^\n]+)/i);
    const subjectMatch = seg.match(/\*\*Subject:\*\*\s*([^\n]+)/i);
    const attMatch = seg.match(/\*\*Attachments?:\*\*\s*([^\n]+)/i);

    let body = seg;
    body = body.replace(/^###[^\n]*\n?/, "");
    body = body.replace(/\*\*To:\*\*[^\n]*\n?/i, "");
    body = body.replace(/\*\*Subject:\*\*[^\n]*\n?/i, "");
    body = body.replace(/\*\*Attachments?:\*\*[^\n]*\n?/i, "");
    body = body.replace(/\[No sign-off[^\]]*\]/i, "");
    body = body.replace(/^---+\s*$/gm, "");
    body = body.trim();

    let attachments: string[] = [];
    if (attMatch) {
      attachments = attMatch[1]
        .replace(/^\[/, "")
        .replace(/\]$/, "")
        .split(/[,;]/)
        .map(s => s.trim())
        .filter(Boolean);
    }

    blocks.push({
      blockKey: `${msgId}::${h}`,
      blockIndex: h,
      rangeStart: segStart,
      rangeEnd: segEnd,
      title: headers[h].title,
      to: toMatch ? toMatch[1].trim() : "",
      subject: subjectMatch ? subjectMatch[1].trim() : "",
      attachments,
      body,
      flagged,
    });
  }
  return blocks;
}

function isPlaceholderEmail(value: string): boolean {
  if (!value) return true;
  const v = value.trim();
  if (!v) return true;
  // Bracketed placeholders like "[need client email - please provide or check XPM]"
  if (/^\[.*\]$/.test(v)) return true;
  // Loose check: if it doesn't contain @ it's not a real address yet.
  return !/@/.test(v) || /need client email|please provide/i.test(v);
}

function bodyToHtml(body: string): string {
  return body
    .split(/\n{2,}/)
    .map(p =>
      p
        .split("\n")
        .map(line => line.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"))
        .join("<br>")
    )
    .map(p => `<p>${p}</p>`)
    .join("");
}

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
  const [uploadProgress, setUploadProgress] = useState<{ completed: number; total: number } | null>(null);
  const [dragActive, setDragActive] = useState(false);
  const [createdDrafts, setCreatedDrafts] = useState<Record<string, true>>({});
  const [draftModal, setDraftModal] = useState<{
    blockKey: string;
    title: string;
    to: string;
    subject: string;
    body: string;
    selectedAttachmentIds: Record<string, boolean>;
    availableFiles: ChatFileRef[];
  } | null>(null);
  const [creatingDraft, setCreatingDraft] = useState(false);
  const [exportMenuOpen, setExportMenuOpen] = useState(false);
  const [exportingType, setExportingType] = useState<"transcript" | "summary" | "recommendation" | null>(null);
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

  // Save the conversation when the window is about to unload (app close,
  // refresh). sendBeacon is fire-and-forget and survives page teardown.
  useEffect(() => {
    const onBeforeUnload = () => {
      const agent = agents.find(a => a.id === selectedAgentId);
      persistConversation(
        messages,
        selectedAgentId,
        agent?.name || "General Chat",
        clientName,
        entityName,
        true,
      );
    };
    window.addEventListener("beforeunload", onBeforeUnload);
    return () => window.removeEventListener("beforeunload", onBeforeUnload);
  }, [messages, selectedAgentId, agents, clientName, entityName]);

  const persistConversation = (
    msgs: Message[],
    agentId: string,
    agentName: string,
    client: string,
    entity: string,
    useBeacon = false,
  ) => {
    if (msgs.length <= 1) return;
    const payload = {
      messages: msgs.map(m => ({ role: m.role, content: m.content })),
      agent_id: agentId,
      agent_name: agentName,
      client_name: client.trim(),
      entity_name: entity.trim(),
    };
    const url = "http://127.0.0.1:7842/api/chat/save-conversation";
    if (useBeacon && navigator.sendBeacon) {
      navigator.sendBeacon(
        url,
        new Blob([JSON.stringify(payload)], { type: "application/json" }),
      );
      return;
    }
    fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
      keepalive: true,
    }).catch(() => {});
  };

  const handleAgentChange = (newId: string) => {
    if (newId === selectedAgentId) return;
    const outgoing = agents.find(a => a.id === selectedAgentId);
    persistConversation(
      messages,
      selectedAgentId,
      outgoing?.name || "General Chat",
      clientName,
      entityName,
    );
    const agent = agents.find(a => a.id === newId);
    setSelectedAgentId(newId);
    if (agent) toast(`Switched to ${agent.name} — conversation cleared.`);
  };

  const uploadFiles = async (files: File[]) => {
    if (files.length === 0) return;
    if (attachedFiles.length + files.length > MAX_FILES_PER_MESSAGE) {
      toast.error(`Max ${MAX_FILES_PER_MESSAGE} files per message`);
      return;
    }
    setUploading(true);
    setUploadingFiles(files.map(f => ({ name: f.name, size: f.size })));
    setUploadProgress({ completed: 0, total: files.length });
    try {
      let completed = 0;
      for (const f of files) {
        try {
          const ref = await uploadChatFile(f);
          setAttachedFiles(prev => [...prev, ref]);
        } catch (e: any) {
          toast.error(`Upload failed: ${f.name}`, { description: e?.message ?? "" });
        } finally {
          completed += 1;
          setUploadProgress({ completed, total: files.length });
          setUploadingFiles(prev => prev.filter(p => p.name !== f.name));
        }
      }
    } finally {
      setUploading(false);
      setUploadingFiles([]);
      setUploadProgress(null);
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

  // For an assistant message, find the most recent prior user message's
  // attached files — those are the candidate attachments for any draft
  // emails the assistant wrote.
  const filesAvailableForAssistant = (assistantMsgId: number): ChatFileRef[] => {
    const idx = messages.findIndex(m => m.id === assistantMsgId);
    if (idx < 0) return [];
    for (let i = idx - 1; i >= 0; i--) {
      if (messages[i].role === "user" && messages[i].files?.length) {
        return messages[i].files || [];
      }
    }
    return [];
  };

  const openDraftModal = (block: EmailBlock, msgId: number) => {
    const available = filesAvailableForAssistant(msgId);
    // Pre-check files whose name matches one the assistant referenced.
    const refNames = new Set(block.attachments.map(s => s.toLowerCase()));
    const selected: Record<string, boolean> = {};
    for (const f of available) {
      selected[f.id] = refNames.has(f.name.toLowerCase());
    }
    setDraftModal({
      blockKey: block.blockKey,
      title: block.title,
      to: isPlaceholderEmail(block.to) ? "" : block.to,
      subject: block.subject,
      body: block.body,
      selectedAttachmentIds: selected,
      availableFiles: available,
    });
  };

  const submitDraft = async () => {
    if (!draftModal) return;
    if (!draftModal.to.trim() || !/@/.test(draftModal.to)) {
      toast.error("Enter a valid recipient email address");
      return;
    }
    if (!draftModal.subject.trim()) {
      toast.error("Subject is required");
      return;
    }
    setCreatingDraft(true);
    try {
      const attachment_ids = Object.entries(draftModal.selectedAttachmentIds)
        .filter(([, on]) => on)
        .map(([id]) => id);
      const res = await fetch("http://127.0.0.1:7842/api/chat/create-draft", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          to: draftModal.to.trim(),
          subject: draftModal.subject.trim(),
          body: bodyToHtml(draftModal.body),
          attachment_ids,
        }),
      });
      const body = await res.json().catch(() => ({}));
      if (!res.ok || !body.ok) {
        throw new Error(body.error || `HTTP ${res.status}`);
      }
      setCreatedDrafts(prev => ({ ...prev, [draftModal.blockKey]: true }));
      toast.success(`Draft created in Outlook for ${draftModal.title}`);
      setDraftModal(null);
    } catch (e: any) {
      toast.error("Could not create draft", { description: e?.message ?? "" });
    } finally {
      setCreatingDraft(false);
    }
  };

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
              PDFs, Excel, Word, CSV, images — up to {MAX_FILES_PER_MESSAGE} per message
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
                    className="w-full text-left px-3 py-2 text-xs hover:bg-blue-50 disabled:opacity-50 disabled:hover:bg-white disabled:cursor-not-allowed"
                  >
                    <div className="font-semibold text-slate-700">Download Recommendation (.docx)</div>
                    <div className="text-slate-500">
                      {hasStructuredRecommendation
                        ? "Specialist's structured output as a formatted document"
                        : "Available once the specialist produces a structured doc"}
                    </div>
                  </button>
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
      <div className="chat-messages flex-1 overflow-y-auto p-6 space-y-4">
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
              {msg.role === "assistant" && (() => {
                const blocks = parseEmailBlocks(msg.content, msg.id);
                if (blocks.length === 0) return null;
                return (
                  <div className="mt-3 pt-3 border-t border-slate-200 flex flex-col gap-1.5">
                    {blocks.map(b => {
                      if (b.flagged) {
                        return (
                          <div
                            key={b.blockKey}
                            className="text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded px-2 py-1"
                          >
                            ⚠️ {b.title} — flagged for review, no draft button.
                          </div>
                        );
                      }
                      const done = !!createdDrafts[b.blockKey];
                      return (
                        <button
                          key={b.blockKey}
                          type="button"
                          disabled={done}
                          onClick={() => openDraftModal(b, msg.id)}
                          className={`inline-flex items-center gap-1.5 self-start px-2.5 py-1 rounded-md text-xs font-medium border transition-all ${
                            done
                              ? "border-slate-200 bg-slate-50 text-slate-400 cursor-default"
                              : "border-blue-200 bg-blue-50 text-blue-700 hover:border-blue-400 hover:bg-blue-100"
                          }`}
                          title={done ? "Already created" : `Create Outlook draft for ${b.title}`}
                        >
                          {done ? <Check className="w-3 h-3" /> : <Mail className="w-3 h-3" />}
                          {done ? "Draft Created" : `Create Draft — ${b.title}`}
                        </button>
                      );
                    })}
                  </div>
                );
              })()}
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
          {uploadProgress && uploadProgress.total > 1 && (
            <div className="text-xs text-slate-500 mb-1.5">
              Uploading {uploadProgress.completed} of {uploadProgress.total} files…
            </div>
          )}
          <div className="flex flex-wrap gap-1.5 max-h-32 overflow-y-auto">
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
                disabled={uploading || attachedFiles.length >= MAX_FILES_PER_MESSAGE}
                className="flex-shrink-0 w-9 h-9 rounded-lg flex items-center justify-center text-slate-500 hover:text-blue-600 hover:bg-slate-100 transition-all disabled:opacity-40"
                title={uploading ? "Uploading..." : `Attach files (max ${MAX_FILES_PER_MESSAGE})`}
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

      {/* Create Draft modal */}
      {draftModal && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40"
          onClick={() => !creatingDraft && setDraftModal(null)}
        >
          <div
            className="bg-white rounded-xl shadow-xl w-[520px] max-w-[92vw] max-h-[90vh] overflow-y-auto"
            onClick={e => e.stopPropagation()}
          >
            <div className="px-5 py-4 border-b border-border flex items-center justify-between">
              <div className="flex items-center gap-2 text-sm font-semibold text-slate-800">
                <Mail className="w-4 h-4 text-blue-600" />
                Create Draft in Outlook — {draftModal.title}
              </div>
              <button
                type="button"
                onClick={() => !creatingDraft && setDraftModal(null)}
                className="text-slate-400 hover:text-slate-600"
                aria-label="Close"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <div className="px-5 py-4 space-y-3 text-sm">
              <label className="block">
                <span className="text-xs font-medium text-slate-600">To</span>
                <input
                  type="email"
                  value={draftModal.to}
                  onChange={e => setDraftModal(m => m && { ...m, to: e.target.value })}
                  placeholder="client@example.com"
                  className="mt-1 w-full px-3 py-2 rounded-md border border-border focus:border-blue-400 focus:outline-none"
                />
              </label>
              <label className="block">
                <span className="text-xs font-medium text-slate-600">Subject</span>
                <input
                  type="text"
                  value={draftModal.subject}
                  onChange={e => setDraftModal(m => m && { ...m, subject: e.target.value })}
                  className="mt-1 w-full px-3 py-2 rounded-md border border-border focus:border-blue-400 focus:outline-none"
                />
              </label>
              {draftModal.availableFiles.length > 0 && (
                <div>
                  <div className="text-xs font-medium text-slate-600 mb-1">Attach</div>
                  <div className="border border-border rounded-md max-h-40 overflow-y-auto divide-y divide-border">
                    {draftModal.availableFiles.map(f => (
                      <label
                        key={f.id}
                        className="flex items-center gap-2 px-3 py-2 hover:bg-slate-50 cursor-pointer"
                      >
                        <input
                          type="checkbox"
                          checked={!!draftModal.selectedAttachmentIds[f.id]}
                          onChange={e =>
                            setDraftModal(m =>
                              m && {
                                ...m,
                                selectedAttachmentIds: {
                                  ...m.selectedAttachmentIds,
                                  [f.id]: e.target.checked,
                                },
                              },
                            )
                          }
                        />
                        <FileText className="w-3.5 h-3.5 text-slate-500" />
                        <span className="text-xs text-slate-700 flex-1 truncate">{f.name}</span>
                        <span className="text-xs text-slate-400">{formatBytes(f.size)}</span>
                      </label>
                    ))}
                  </div>
                </div>
              )}
              <div>
                <span className="text-xs font-medium text-slate-600">Body preview</span>
                <div className="mt-1 px-3 py-2 rounded-md border border-border bg-slate-50 text-xs text-slate-700 whitespace-pre-wrap max-h-48 overflow-y-auto">
                  {draftModal.body}
                </div>
              </div>
            </div>
            <div className="px-5 py-3 border-t border-border flex items-center justify-end gap-2">
              <button
                type="button"
                onClick={() => setDraftModal(null)}
                disabled={creatingDraft}
                className="px-3 py-1.5 rounded-md text-xs font-medium border border-border bg-white text-slate-700 hover:bg-slate-50 disabled:opacity-60"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={submitDraft}
                disabled={creatingDraft}
                className="px-3 py-1.5 rounded-md text-xs font-medium text-white disabled:opacity-60"
                style={{ background: "oklch(0.5 0.2 250)" }}
              >
                {creatingDraft ? "Creating…" : "Create Draft"}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
