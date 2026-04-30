import { useEffect, useState } from "react";
import { Folder, ChevronRight, FolderPlus, X } from "lucide-react";
import { toast } from "sonner";

interface FolderEntry {
  name: string;
  is_folder: boolean;
}

export type SharePointExportType = "transcript" | "summary" | "recommendation";

interface Props {
  isOpen: boolean;
  clientName: string;
  defaultExportType?: SharePointExportType;
  hasStructuredRecommendation?: boolean;
  onCancel: () => void;
  onSelect: (path: string, exportType: SharePointExportType) => void;
}

const API_BASE = "http://127.0.0.1:7842";

export default function SharePointFolderPicker({
  isOpen,
  clientName,
  defaultExportType = "transcript",
  hasStructuredRecommendation = false,
  onCancel,
  onSelect,
}: Props) {
  const [path, setPath] = useState<string>("");
  const [folders, setFolders] = useState<FolderEntry[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [creatingFolder, setCreatingFolder] = useState(false);
  const [newFolderName, setNewFolderName] = useState("");
  const [exportType, setExportType] = useState<SharePointExportType>(defaultExportType);

  useEffect(() => {
    if (!isOpen) return;
    setPath("");
    setExportType(defaultExportType);
    setCreatingFolder(false);
    setNewFolderName("");
  }, [isOpen, defaultExportType]);

  useEffect(() => {
    if (!isOpen || !clientName) return;
    let cancelled = false;
    setLoading(true);
    setError(null);
    const url = `${API_BASE}/api/sharepoint/browse?client=${encodeURIComponent(
      clientName,
    )}&path=${encodeURIComponent(path)}`;
    fetch(url)
      .then(r => r.json())
      .then(data => {
        if (cancelled) return;
        if (!data.ok) {
          setError(data.error || "Failed to load folders");
          setFolders([]);
          return;
        }
        setFolders(data.folders || []);
      })
      .catch(e => {
        if (cancelled) return;
        setError(e?.message || "Failed to load folders");
        setFolders([]);
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [isOpen, clientName, path]);

  const segments = path ? path.split("/").filter(Boolean) : [];

  const navigateInto = (folderName: string) => {
    setPath(p => (p ? `${p}/${folderName}` : folderName));
  };

  const navigateToSegment = (idx: number) => {
    if (idx < 0) {
      setPath("");
      return;
    }
    setPath(segments.slice(0, idx + 1).join("/"));
  };

  const handleCreateFolder = async () => {
    const name = newFolderName.trim();
    if (!name) return;
    try {
      const res = await fetch(`${API_BASE}/api/sharepoint/create-folder`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ client_name: clientName, path, folder_name: name }),
      });
      const body = await res.json().catch(() => ({}));
      if (!res.ok || !body.ok) {
        throw new Error(body.error || `HTTP ${res.status}`);
      }
      toast.success(`Created folder: ${name}`);
      setCreatingFolder(false);
      setNewFolderName("");
      // Re-trigger fetch by tweaking path key — re-set to same value forces refresh.
      const current = path;
      setPath("");
      setTimeout(() => setPath(current), 0);
    } catch (e: any) {
      toast.error("Could not create folder", { description: e?.message ?? "" });
    }
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40">
      <div className="bg-white rounded-xl shadow-xl w-full max-w-lg max-h-[85vh] flex flex-col">
        <div className="flex items-center justify-between px-5 py-4 border-b border-border">
          <div className="font-semibold text-base">
            Save to SharePoint — {clientName}
          </div>
          <button
            onClick={onCancel}
            className="text-muted-foreground hover:text-foreground"
            aria-label="Close"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <div className="px-5 py-3 border-b border-border text-sm text-muted-foreground flex items-center gap-1 flex-wrap">
          <button
            onClick={() => navigateToSegment(-1)}
            className="hover:text-foreground hover:underline"
          >
            {clientName}
          </button>
          {segments.map((seg, idx) => (
            <span key={idx} className="flex items-center gap-1">
              <ChevronRight className="w-3 h-3" />
              <button
                onClick={() => navigateToSegment(idx)}
                className="hover:text-foreground hover:underline"
              >
                {seg}
              </button>
            </span>
          ))}
        </div>

        <div className="flex-1 overflow-y-auto p-2 min-h-[240px]">
          {loading && (
            <div className="p-4 text-sm text-muted-foreground">Loading…</div>
          )}
          {!loading && error && (
            <div className="p-4 text-sm text-red-600">{error}</div>
          )}
          {!loading && !error && folders.length === 0 && (
            <div className="p-4 text-sm text-muted-foreground">
              No subfolders here. You can save in this folder or create a new one.
            </div>
          )}
          {!loading &&
            !error &&
            folders.map(f => (
              <button
                key={f.name}
                onDoubleClick={() => navigateInto(f.name)}
                onClick={() => navigateInto(f.name)}
                className="w-full flex items-center gap-2 px-3 py-2 text-sm rounded-lg hover:bg-slate-100 text-left"
              >
                <Folder className="w-4 h-4 text-blue-500" />
                <span className="flex-1 truncate">{f.name}</span>
              </button>
            ))}
        </div>

        <div className="px-5 py-3 border-t border-border space-y-3">
          <div className="text-xs text-muted-foreground">
            Saving to: <span className="font-mono text-foreground">/{segments.join("/") || ""}</span>
          </div>

          <div className="flex items-center gap-4 text-sm">
            <span className="text-muted-foreground">Export as:</span>
            {(["transcript", "summary", "recommendation"] as SharePointExportType[]).map(t => {
              const disabled = t === "recommendation" && !hasStructuredRecommendation;
              return (
                <label
                  key={t}
                  className={`flex items-center gap-1.5 ${
                    disabled ? "opacity-40 cursor-not-allowed" : "cursor-pointer"
                  }`}
                >
                  <input
                    type="radio"
                    checked={exportType === t}
                    disabled={disabled}
                    onChange={() => setExportType(t)}
                  />
                  <span className="capitalize">{t}</span>
                </label>
              );
            })}
          </div>

          {creatingFolder ? (
            <div className="flex items-center gap-2">
              <input
                autoFocus
                value={newFolderName}
                onChange={e => setNewFolderName(e.target.value)}
                onKeyDown={e => {
                  if (e.key === "Enter") handleCreateFolder();
                  if (e.key === "Escape") {
                    setCreatingFolder(false);
                    setNewFolderName("");
                  }
                }}
                placeholder="New folder name"
                className="flex-1 px-3 py-1.5 rounded-lg border border-border text-sm focus:outline-none focus:border-blue-400"
              />
              <button
                onClick={handleCreateFolder}
                className="px-3 py-1.5 rounded-lg bg-blue-500 text-white text-sm hover:bg-blue-600"
              >
                Create
              </button>
              <button
                onClick={() => {
                  setCreatingFolder(false);
                  setNewFolderName("");
                }}
                className="px-3 py-1.5 rounded-lg border border-border text-sm"
              >
                Cancel
              </button>
            </div>
          ) : (
            <div className="flex items-center justify-between">
              <button
                onClick={() => setCreatingFolder(true)}
                className="flex items-center gap-1.5 text-sm text-muted-foreground hover:text-foreground"
              >
                <FolderPlus className="w-4 h-4" />
                New Folder
              </button>
              <div className="flex items-center gap-2">
                <button
                  onClick={onCancel}
                  className="px-4 py-1.5 rounded-lg border border-border text-sm hover:bg-slate-50"
                >
                  Cancel
                </button>
                <button
                  onClick={() => onSelect(path, exportType)}
                  className="px-4 py-1.5 rounded-lg bg-blue-500 text-white text-sm hover:bg-blue-600"
                >
                  Save Here
                </button>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
