import { createRoot } from "react-dom/client";
import App from "./App";
import "./index.css";

// ── Per-install API token injection ────────────────────────────────────────────
// The desktop shell (pywebview) sets window.__API_TOKEN__ after the page loads.
// Wrap fetch() and EventSource to attach the token to every request to our
// Flask API, so browser tabs or other apps on the machine can't call the API
// without the token.
(function installApiAuth() {
  if (typeof window === "undefined") return;

  const isOurApi = (url: string): boolean =>
    url.includes("127.0.0.1:7842") ||
    url.includes("localhost:7842") ||
    url.startsWith("/api/");

  // Poll briefly for the token — it's injected via on_loaded after scripts run,
  // so very early fetches may race it.
  const waitForToken = (timeoutMs = 5000): Promise<string> => {
    return new Promise(resolve => {
      const start = Date.now();
      const tick = () => {
        const t = (window as any).__API_TOKEN__;
        if (t) return resolve(t);
        if (Date.now() - start >= timeoutMs) return resolve("");
        setTimeout(tick, 50);
      };
      tick();
    });
  };

  const origFetch = window.fetch.bind(window);
  window.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
    const url =
      typeof input === "string"
        ? input
        : input instanceof URL
        ? input.toString()
        : (input as Request).url;
    if (!isOurApi(url)) return origFetch(input as any, init);
    const token = (window as any).__API_TOKEN__ || (await waitForToken());
    if (!token) return origFetch(input as any, init);
    const headers = new Headers(init?.headers);
    if (!headers.has("Authorization")) {
      headers.set("Authorization", `Bearer ${token}`);
    }
    return origFetch(input as any, { ...(init || {}), headers });
  }) as typeof fetch;

  const OrigES = (window as any).EventSource;
  if (OrigES) {
    const Wrapped: any = function (url: string | URL, init?: EventSourceInit) {
      let urlStr = typeof url === "string" ? url : url.toString();
      const token = (window as any).__API_TOKEN__ || "";
      if (isOurApi(urlStr) && token) {
        const sep = urlStr.includes("?") ? "&" : "?";
        urlStr = `${urlStr}${sep}token=${encodeURIComponent(token)}`;
      }
      return new OrigES(urlStr, init);
    };
    Wrapped.prototype = OrigES.prototype;
    Wrapped.CONNECTING = OrigES.CONNECTING;
    Wrapped.OPEN = OrigES.OPEN;
    Wrapped.CLOSED = OrigES.CLOSED;
    (window as any).EventSource = Wrapped;
  }
})();

createRoot(document.getElementById("root")!).render(<App />);
