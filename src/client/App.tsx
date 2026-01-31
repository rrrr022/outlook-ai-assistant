import React, { useState, useEffect } from 'react';
import ChatPanel from './components/ChatPanel';
import EmailPanel from './components/EmailPanel';
import CalendarPanel from './components/CalendarPanel';
import TasksPanel from './components/TasksPanel';
import SettingsPanel from './components/SettingsPanel';
import AnalyticsPanel from './components/AnalyticsPanel';
import ApprovalModal from './components/ApprovalModal';
import OnboardingPanel from './components/OnboardingPanel';
import { Spinner } from './components/ui/NativeComponents';
import { useAppStore } from './store/appStore';
import { useAPIKeyStore } from './store/apiKeyStore';
import { approvalService } from './services/approvalService';

type TabType = 'chat' | 'email' | 'calendar' | 'tasks' | 'analytics' | 'settings';

// Check if onboarding has been completed
const ONBOARDING_KEY = 'outlook-ai-onboarding-complete';

const App: React.FC = () => {
  const [activeTab, setActiveTab] = useState<TabType>('chat');
  const [showOnboarding, setShowOnboarding] = useState(false);
  const { isLoading, currentApproval, pendingApprovals } = useAppStore();
  const { apiKeys } = useAPIKeyStore();

  // Check if user needs onboarding
  useEffect(() => {
    const hasCompletedOnboarding = localStorage.getItem(ONBOARDING_KEY);
    const hasAnyKey = Object.values(apiKeys).some(key => key?.trim());
    
    // Show onboarding if never completed AND no keys set
    if (!hasCompletedOnboarding && !hasAnyKey) {
      setShowOnboarding(true);
    }
  }, [apiKeys]);

  const handleOnboardingComplete = () => {
    localStorage.setItem(ONBOARDING_KEY, 'true');
    setShowOnboarding(false);
  };

  // If showing onboarding, render only that
  if (showOnboarding) {
    return <OnboardingPanel onComplete={handleOnboardingComplete} />;
  }

  const tabs: { id: TabType; label: string; icon: string }[] = [
    { id: 'chat', label: 'Chat', icon: '⌨️' },
    { id: 'email', label: 'Mail', icon: '📨' },
    { id: 'calendar', label: 'Cal', icon: '📅' },
    { id: 'tasks', label: 'Tasks', icon: '☑️' },
    { id: 'analytics', label: 'Stats', icon: '📊' },
    { id: 'settings', label: 'Config', icon: '⚙️' },
  ];

  const renderContent = () => {
    switch (activeTab) {
      case 'chat':
        return <ChatPanel />;
      case 'email':
        return <EmailPanel />;
      case 'calendar':
        return <CalendarPanel />;
      case 'tasks':
        return <TasksPanel />;
      case 'analytics':
        return <AnalyticsPanel />;
      case 'settings':
        return <SettingsPanel />;
      default:
        return <ChatPanel />;
    }
  };

  const handleApprove = async (request: typeof currentApproval) => {
    if (request) {
      await approvalService.handleApproval(request);
    }
  };

  const handleReject = (request: typeof currentApproval) => {
    if (request) {
      approvalService.handleRejection(request);
    }
  };

  return (
    <div className="app-container">
      <header className="app-header">
        <span className="header-icon">🤖</span>
        <h1>Outlook AI</h1>
        <span className="header-version">v1.5.0.0</span>
        {pendingApprovals.length > 0 && (
          <div className="pending-badge" title={`${pendingApprovals.length} pending approval(s)`}>
            <span className="pending-count">{pendingApprovals.length}</span>
          </div>
        )}
      </header>

      <nav className="tab-container">
        {tabs.map((tab) => (
          <button
            key={tab.id}
            className={`tab-button ${activeTab === tab.id ? 'active' : ''}`}
            onClick={() => setActiveTab(tab.id)}
          >
            <span className="tab-icon">{tab.icon}</span>
            <span>{tab.label}</span>
          </button>
        ))}
      </nav>

      <main className="content-area">
        {isLoading ? (
          <div className="loading-spinner">
            <Spinner size="medium" label="Loading..." />
          </div>
        ) : (
          renderContent()
        )}
      </main>

      {/* Approval Modal */}
      <ApprovalModal 
        request={currentApproval}
        onApprove={handleApprove}
        onReject={handleReject}
      />

      {/* Footer with credits */}
      <footer className="app-footer">
        <div className="credits">
          Built by <span className="brand-name">FreedomForged_AI</span> • 
          <span className="status-indicator">● ONLINE</span>
        </div>
      </footer>
    </div>
  );
};

export default App;
