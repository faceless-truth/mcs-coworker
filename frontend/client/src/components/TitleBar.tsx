/**
 * TitleBar — Custom title bar for the Electron frameless window.
 * Only renders when running inside Electron (window.__ELECTRON__ is set).
 * Provides window drag region, minimize, maximize, and close buttons.
 */

import React from "react";
import { Minus, Square, X } from "lucide-react";

declare global {
  interface Window {
    electronAPI?: {
      minimize: () => void;
      maximize: () => void;
      close: () => void;
      quit: () => void;
      onNavigate: (callback: (page: string) => void) => void;
    };
    __ELECTRON__?: boolean;
  }
}

export function TitleBar() {
  // Only render inside Electron
  if (!window.__ELECTRON__) return null;

  return (
    <div
      className="flex items-center justify-between h-9 px-4 select-none shrink-0"
      style={{
        background: "#0a1628",
        // @ts-ignore — Electron-specific CSS property
        WebkitAppRegion: "drag",
        borderBottom: "1px solid rgba(255,255,255,0.06)",
      } as React.CSSProperties}
    >
      {/* App name */}
      <span
        className="text-xs font-medium tracking-widest uppercase"
        style={{ color: "rgba(255,255,255,0.4)", letterSpacing: "0.15em" }}
      >
        MC&amp;S CoWorker
      </span>

      {/* Window controls */}
      <div
        className="flex items-center gap-1"
        style={{ WebkitAppRegion: "no-drag" } as React.CSSProperties}
      >
        <button
          onClick={() => window.electronAPI?.minimize()}
          className="w-7 h-7 flex items-center justify-center rounded hover:bg-white/10 transition-colors"
          title="Minimise"
        >
          <Minus size={12} className="text-white/50" />
        </button>
        <button
          onClick={() => window.electronAPI?.maximize()}
          className="w-7 h-7 flex items-center justify-center rounded hover:bg-white/10 transition-colors"
          title="Maximise"
        >
          <Square size={11} className="text-white/50" />
        </button>
        <button
          onClick={() => window.electronAPI?.close()}
          className="w-7 h-7 flex items-center justify-center rounded hover:bg-red-500/80 transition-colors group"
          title="Close to tray"
        >
          <X size={13} className="text-white/50 group-hover:text-white" />
        </button>
      </div>
    </div>
  );
}
