import { Toaster } from "@/components/ui/sonner";
import { TooltipProvider } from "@/components/ui/tooltip";
import { ThemeProvider } from "./contexts/ThemeContext";
import { useState, useEffect } from "react";
import Sidebar from "./components/Sidebar";
import { TitleBar } from "./components/TitleBar";
import Dashboard from "./pages/Dashboard";
import Plugins from "./pages/Plugins";
import Approvals from "./pages/Approvals";
import Activity from "./pages/Activity";
import Memory from "./pages/Memory";
import KpiMonitor from "./pages/KpiMonitor";
import Usage from "./pages/Usage";
import Settings from "./pages/Settings";
import Chat from "./pages/Chat";
import EmailRules from "./pages/EmailRules";
import StaffNotify from "./pages/StaffNotify";
import { mockApprovals } from "./lib/mockData";

function App() {
  const [activePage, setActivePage] = useState("dashboard");
  const [pendingApprovals] = useState(mockApprovals.length);

  // Wire Electron tray navigation events
  useEffect(() => {
    if (window.__ELECTRON__ && window.electronAPI?.onNavigate) {
      window.electronAPI.onNavigate((page) => setActivePage(page));
    }
  }, []);

  const renderPage = () => {
    switch (activePage) {
      case "dashboard": return <Dashboard />;
      case "plugins": return <Plugins />;
      case "approvals": return <Approvals />;
      case "activity": return <Activity />;
      case "memory": return <Memory />;
      case "kpi": return <KpiMonitor />;
      case "usage": return <Usage />;
      case "settings": return <Settings />;
      case "chat": return <Chat />;
      case "rules": return <EmailRules />;
      case "staff": return <StaffNotify />;
      default: return <Dashboard />;
    }
  };

  return (
    <ThemeProvider defaultTheme="light">
      <TooltipProvider>
        <Toaster position="top-right" />
        <div className="flex flex-col h-screen overflow-hidden bg-background">
          <TitleBar />
          <div className="flex flex-1 overflow-hidden">
            <Sidebar
              activePage={activePage}
              onNavigate={setActivePage}
              pendingApprovals={pendingApprovals}
            />
            <main className="flex-1 overflow-y-auto">
              {renderPage()}
            </main>
          </div>
        </div>
      </TooltipProvider>
    </ThemeProvider>
  );
}

export default App;
