import React, { useState, useRef, useEffect, useCallback } from 'react';
import { Button, Input, SendIcon } from './ui/NativeComponents';
import { useAppStore } from '../store/appStore';
import { agentRuntime, AgentState, PendingApproval } from '../services/agentRuntime';
import { brandingService, UserBranding } from '../services/brandingService';
import BrandingPanel from './BrandingPanel';
import { v4 as uuidv4 } from 'uuid';

// Lazy load graph service for sign-in
const loadGraphService = () => import(/* webpackChunkName: "graph-service" */ '../services/graphService');

// Types for document export
type DocumentType = 'word' | 'pdf' | 'excel' | 'powerpoint';
const loadDocumentService = () => import(/* webpackChunkName: "document-service" */ '../services/documentService');

const AgentChatPanel: React.FC = () => {
  const [input, setInput] = useState('');
  const [isSignedIn, setIsSignedIn] = useState(false);
  const [userName, setUserName] = useState<string | null>(null);
  const [showBranding, setShowBranding] = useState(false);
  const [userBranding, setUserBranding] = useState<UserBranding>(brandingService.load());
  const [agentState, setAgentState] = useState<AgentState>(agentRuntime.getState());
  const [intermediateMessages, setIntermediateMessages] = useState<string[]>([]);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const { messages, addMessage, currentEmail, tasks } = useAppStore();

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  };

  useEffect(() => {
    scrollToBottom();
  }, [messages, intermediateMessages]);

  // Subscribe to agent state updates
  useEffect(() => {
    const unsubState = agentRuntime.onStateUpdate(setAgentState);
    const unsubMessage = agentRuntime.onMessage((msg, isIntermediate) => {
      if (isIntermediate) {
        setIntermediateMessages(prev => [...prev, msg]);
      }
    });

    // Initialize agent
    agentRuntime.initialize();

    return () => {
      unsubState();
      unsubMessage();
    };
  }, []);

  // Check sign-in status on mount
  useEffect(() => {
    const checkSignIn = async () => {
      const { graphService } = await loadGraphService();
      setIsSignedIn(graphService.isSignedIn);
      if (graphService.isSignedIn) {
        const userInfo = graphService.getUserInfo();
        setUserName(userInfo?.name || null);
      }
    };
    checkSignIn();
  }, []);

  // Handle Microsoft sign-in
  const handleMicrosoftSignIn = async () => {
    try {
      const { graphService } = await loadGraphService();
      const success = await graphService.signIn();
      if (success) {
        setIsSignedIn(true);
        const userInfo = graphService.getUserInfo();
        setUserName(userInfo?.name || null);
        addMessage({
          id: uuidv4(),
          role: 'assistant',
          content: `✅ **Signed in as ${userInfo?.name || 'User'}!**\n\nI now have full access to your inbox, calendar, and can take actions on your behalf. Just tell me what you need!`,
          timestamp: new Date(),
        });
        // Reinitialize agent with user context
        agentRuntime.initialize();
      }
    } catch (error: any) {
      addMessage({
        id: uuidv4(),
        role: 'assistant',
        content: `❌ **Sign-in failed**\n\n${error.message || 'Please try again.'}`,
        timestamp: new Date(),
      });
    }
  };

  // Handle sign out
  const handleSignOut = async () => {
    const { graphService } = await loadGraphService();
    await graphService.signOut();
    setIsSignedIn(false);
    setUserName(null);
    agentRuntime.clearConversation();
    addMessage({
      id: uuidv4(),
      role: 'assistant',
      content: '👋 Signed out. Some features will be limited.',
      timestamp: new Date(),
    });
  };

  // Quick actions
  const quickActions = [
    { icon: '📥', label: 'Review Inbox', prompt: 'Review my inbox and tell me what needs my attention' },
    { icon: '📅', label: 'My Calendar', prompt: 'Show me my calendar for today and help me plan my day' },
    { icon: '📋', label: 'My Tasks', prompt: 'What tasks do I have pending and how should I prioritize them?' },
    { icon: '🔍', label: 'Search', prompt: 'search_prompt' },
    { icon: '✉️', label: 'Compose', prompt: 'Help me compose a new email' },
    { icon: '🎨', label: 'Branding', prompt: 'branding' },
  ];

  const handleQuickAction = async (action: typeof quickActions[0]) => {
    if (action.prompt === 'branding') {
      setShowBranding(true);
      return;
    }
    if (action.prompt === 'search_prompt') {
      addMessage({
        id: uuidv4(),
        role: 'assistant',
        content: '🔍 What would you like to search for? Tell me the topic, person, or keywords.',
        timestamp: new Date(),
      });
      return;
    }
    await handleSend(action.prompt);
  };

  // Main send handler - delegates to agent runtime
  const handleSend = async (messageText?: string) => {
    const text = messageText || input.trim();
    if (!text || agentState.isProcessing) return;

    // Clear intermediate messages
    setIntermediateMessages([]);

    // Add user message to UI
    addMessage({
      id: uuidv4(),
      role: 'user',
      content: text,
      timestamp: new Date(),
    });

    setInput('');

    // Let the agent runtime handle everything
    const response = await agentRuntime.processUserMessage(text);

    // Add final response to UI
    addMessage({
      id: uuidv4(),
      role: 'assistant',
      content: response,
      timestamp: new Date(),
    });

    // Clear intermediate messages now that we have final response
    setIntermediateMessages([]);
  };

  // Handle approval/rejection of pending actions
  const handleApproval = async (approval: PendingApproval, approved: boolean) => {
    if (approved) {
      const response = await agentRuntime.approveAction(approval.id);
      addMessage({
        id: uuidv4(),
        role: 'assistant',
        content: response,
        timestamp: new Date(),
      });
    } else {
      const response = await agentRuntime.rejectAction(approval.id, 'User declined');
      addMessage({
        id: uuidv4(),
        role: 'assistant',
        content: response,
        timestamp: new Date(),
      });
    }
  };

  const handleKeyPress = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  const handleBrandingSave = (branding: UserBranding) => {
    brandingService.save(branding);
    setUserBranding(branding);
    setShowBranding(false);
    addMessage({
      id: uuidv4(),
      role: 'assistant',
      content: '✅ **Branding saved!** I\'ll use your preferences for all future communications.',
      timestamp: new Date(),
    });
  };

  // Handle document export
  const handleExport = async (type: DocumentType, content: string) => {
    try {
      const { documentService } = await loadDocumentService();
      const detectedTemplate = documentService.detectTemplate(content);
      const templateInfo = documentService.templates.find(t => t.id === detectedTemplate);
      const title = `${templateInfo?.name || 'AI Response'} - ${new Date().toLocaleDateString()}`;
      
      addMessage({
        id: uuidv4(),
        role: 'assistant',
        content: `📄 Exporting as ${type.toUpperCase()}...`,
        timestamp: new Date(),
      });
      
      // Use the appropriate create method based on type
      if (type === 'word') {
        await documentService.createWord(title, content, { template: detectedTemplate });
      } else if (type === 'pdf') {
        await documentService.createPDF(title, content, { template: detectedTemplate });
      } else if (type === 'excel') {
        await documentService.createExcel(title, content, { template: detectedTemplate });
      } else if (type === 'powerpoint') {
        await documentService.createPowerPoint(title, content, { template: detectedTemplate });
      }
      
      addMessage({
        id: uuidv4(),
        role: 'assistant',
        content: `✅ **Exported successfully!**`,
        timestamp: new Date(),
      });
    } catch (error: any) {
      addMessage({
        id: uuidv4(),
        role: 'assistant',
        content: `❌ Export failed: ${error.message}`,
        timestamp: new Date(),
      });
    }
  };

  // Render pending approvals
  const renderPendingApprovals = () => {
    if (agentState.pendingApprovals.length === 0) return null;

    return (
      <div className="pending-approvals">
        {agentState.pendingApprovals.map(approval => (
          <div key={approval.id} className="approval-card">
            <div className="approval-header">
              <span className="approval-icon">⚠️</span>
              <span className="approval-title">Action Requires Approval</span>
            </div>
            <div className="approval-content">
              <p><strong>{approval.tool}</strong></p>
              <p>{approval.description}</p>
              <pre className="approval-details">
                {JSON.stringify(approval.params, null, 2)}
              </pre>
            </div>
            <div className="approval-actions">
              <Button
                onClick={() => handleApproval(approval, true)}
                className="approve-btn"
              >
                ✅ Approve
              </Button>
              <Button
                onClick={() => handleApproval(approval, false)}
                className="reject-btn"
              >
                ❌ Reject
              </Button>
            </div>
          </div>
        ))}
      </div>
    );
  };

  // Branding panel
  if (showBranding) {
    return (
      <BrandingPanel
        onSave={handleBrandingSave}
        onClose={() => setShowBranding(false)}
      />
    );
  }

  return (
    <div className="chat-panel">
      {/* Header */}
      <div className="chat-header">
        <div className="header-left">
          <h2>🤖 AI Agent</h2>
          <span className="status-indicator">
            {agentState.isProcessing ? '⏳ Working...' : '🟢 Ready'}
          </span>
        </div>
        <div className="header-right">
          {isSignedIn ? (
            <div className="user-status">
              <span className="user-name">👤 {userName}</span>
              <button onClick={handleSignOut} className="sign-out-btn">Sign Out</button>
            </div>
          ) : (
            <button onClick={handleMicrosoftSignIn} className="sign-in-btn">
              🔐 Sign in with Microsoft
            </button>
          )}
        </div>
      </div>

      {/* Quick Actions */}
      <div className="quick-actions">
        {quickActions.map((action, i) => (
          <button
            key={i}
            onClick={() => handleQuickAction(action)}
            className="quick-action-btn"
            disabled={agentState.isProcessing}
          >
            <span className="action-icon">{action.icon}</span>
            <span className="action-label">{action.label}</span>
          </button>
        ))}
      </div>

      {/* Messages */}
      <div className="messages-container">
        {messages.map((message) => (
          <div key={message.id} className={`message ${message.role}`}>
            <div className="message-avatar">
              {message.role === 'user' ? '👤' : '🤖'}
            </div>
            <div className="message-content">
              <div 
                className="message-text"
                dangerouslySetInnerHTML={{ 
                  __html: formatMessage(message.content) 
                }}
              />
              {message.role === 'assistant' && message.content.length > 100 && (
                <div className="export-actions">
                  <button onClick={() => handleExport('word', message.content)}>📄 Word</button>
                  <button onClick={() => handleExport('pdf', message.content)}>📑 PDF</button>
                </div>
              )}
            </div>
          </div>
        ))}

        {/* Intermediate messages (agent working) */}
        {intermediateMessages.map((msg, i) => (
          <div key={`intermediate-${i}`} className="message assistant intermediate">
            <div className="message-avatar">🔧</div>
            <div className="message-content">
              <div className="message-text">{msg}</div>
            </div>
          </div>
        ))}

        {/* Processing indicator */}
        {agentState.isProcessing && (
          <div className="message assistant processing">
            <div className="message-avatar">🤖</div>
            <div className="message-content">
              <div className="typing-indicator">
                <span></span><span></span><span></span>
              </div>
              {agentState.currentTask && (
                <div className="current-task">Working on: {agentState.currentTask}</div>
              )}
            </div>
          </div>
        )}

        {/* Pending approvals */}
        {renderPendingApprovals()}

        <div ref={messagesEndRef} />
      </div>

      {/* Input */}
      <div className="input-container">
        <div className="input-wrapper">
          <textarea
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyPress={handleKeyPress}
            placeholder="Ask me anything or tell me what to do..."
            disabled={agentState.isProcessing}
            rows={1}
          />
          <button
            onClick={() => handleSend()}
            disabled={!input.trim() || agentState.isProcessing}
            className="send-btn"
          >
            <SendIcon />
          </button>
        </div>
        <div className="input-hint">
          {isSignedIn 
            ? "I can search, send emails, manage calendar, and more. Just ask!"
            : "Sign in to enable full inbox and calendar access"
          }
        </div>
      </div>

      <style>{`
        .chat-panel {
          display: flex;
          flex-direction: column;
          height: 100%;
          background: var(--bg-primary, #1e1e1e);
          color: var(--text-primary, #fff);
        }

        .chat-header {
          display: flex;
          justify-content: space-between;
          align-items: center;
          padding: 12px 16px;
          border-bottom: 1px solid var(--border-color, #333);
          background: var(--bg-secondary, #252526);
        }

        .header-left {
          display: flex;
          align-items: center;
          gap: 12px;
        }

        .header-left h2 {
          margin: 0;
          font-size: 16px;
        }

        .status-indicator {
          font-size: 12px;
          color: var(--text-secondary, #888);
        }

        .user-status {
          display: flex;
          align-items: center;
          gap: 8px;
          font-size: 12px;
        }

        .sign-in-btn, .sign-out-btn {
          padding: 6px 12px;
          border-radius: 4px;
          border: none;
          cursor: pointer;
          font-size: 12px;
        }

        .sign-in-btn {
          background: #0078d4;
          color: white;
        }

        .sign-out-btn {
          background: transparent;
          color: var(--text-secondary, #888);
          border: 1px solid var(--border-color, #333);
        }

        .quick-actions {
          display: flex;
          gap: 8px;
          padding: 12px 16px;
          overflow-x: auto;
          border-bottom: 1px solid var(--border-color, #333);
        }

        .quick-action-btn {
          display: flex;
          flex-direction: column;
          align-items: center;
          gap: 4px;
          padding: 8px 12px;
          border-radius: 8px;
          border: 1px solid var(--border-color, #333);
          background: var(--bg-secondary, #252526);
          color: var(--text-primary, #fff);
          cursor: pointer;
          min-width: 70px;
          transition: all 0.2s;
        }

        .quick-action-btn:hover:not(:disabled) {
          background: var(--bg-hover, #2d2d2d);
          border-color: #0078d4;
        }

        .quick-action-btn:disabled {
          opacity: 0.5;
          cursor: not-allowed;
        }

        .action-icon {
          font-size: 20px;
        }

        .action-label {
          font-size: 10px;
        }

        .messages-container {
          flex: 1;
          overflow-y: auto;
          padding: 16px;
          display: flex;
          flex-direction: column;
          gap: 16px;
        }

        .message {
          display: flex;
          gap: 12px;
          animation: fadeIn 0.3s ease;
        }

        @keyframes fadeIn {
          from { opacity: 0; transform: translateY(10px); }
          to { opacity: 1; transform: translateY(0); }
        }

        .message.user {
          flex-direction: row-reverse;
        }

        .message-avatar {
          font-size: 24px;
          flex-shrink: 0;
        }

        .message-content {
          max-width: 80%;
          background: var(--bg-secondary, #252526);
          border-radius: 12px;
          padding: 12px 16px;
        }

        .message.user .message-content {
          background: #0078d4;
        }

        .message.intermediate .message-content {
          background: var(--bg-tertiary, #1a1a1a);
          border: 1px dashed var(--border-color, #333);
          font-style: italic;
          opacity: 0.8;
        }

        .message-text {
          line-height: 1.5;
          white-space: pre-wrap;
          word-break: break-word;
        }

        .message-text strong {
          color: #4fc3f7;
        }

        .message-text code {
          background: rgba(0,0,0,0.3);
          padding: 2px 6px;
          border-radius: 4px;
          font-family: monospace;
        }

        .export-actions {
          display: flex;
          gap: 8px;
          margin-top: 12px;
          padding-top: 12px;
          border-top: 1px solid var(--border-color, #333);
        }

        .export-actions button {
          padding: 4px 8px;
          border-radius: 4px;
          border: 1px solid var(--border-color, #333);
          background: transparent;
          color: var(--text-secondary, #888);
          cursor: pointer;
          font-size: 11px;
        }

        .export-actions button:hover {
          background: var(--bg-hover, #2d2d2d);
        }

        .typing-indicator {
          display: flex;
          gap: 4px;
        }

        .typing-indicator span {
          width: 8px;
          height: 8px;
          background: #0078d4;
          border-radius: 50%;
          animation: bounce 1.4s infinite ease-in-out both;
        }

        .typing-indicator span:nth-child(1) { animation-delay: -0.32s; }
        .typing-indicator span:nth-child(2) { animation-delay: -0.16s; }

        @keyframes bounce {
          0%, 80%, 100% { transform: scale(0); }
          40% { transform: scale(1); }
        }

        .current-task {
          margin-top: 8px;
          font-size: 12px;
          color: var(--text-secondary, #888);
        }

        .pending-approvals {
          display: flex;
          flex-direction: column;
          gap: 12px;
        }

        .approval-card {
          background: #2d1f00;
          border: 1px solid #f5a623;
          border-radius: 12px;
          padding: 16px;
          animation: slideIn 0.3s ease;
        }

        @keyframes slideIn {
          from { opacity: 0; transform: translateX(-20px); }
          to { opacity: 1; transform: translateX(0); }
        }

        .approval-header {
          display: flex;
          align-items: center;
          gap: 8px;
          margin-bottom: 12px;
          font-weight: bold;
          color: #f5a623;
        }

        .approval-content {
          margin-bottom: 12px;
        }

        .approval-details {
          background: rgba(0,0,0,0.3);
          padding: 8px;
          border-radius: 4px;
          font-size: 11px;
          overflow-x: auto;
          max-height: 150px;
          overflow-y: auto;
        }

        .approval-actions {
          display: flex;
          gap: 8px;
        }

        .approve-btn {
          background: #107c10 !important;
          color: white !important;
          border: none !important;
        }

        .reject-btn {
          background: #c42b1c !important;
          color: white !important;
          border: none !important;
        }

        .input-container {
          padding: 16px;
          border-top: 1px solid var(--border-color, #333);
          background: var(--bg-secondary, #252526);
        }

        .input-wrapper {
          display: flex;
          gap: 8px;
          align-items: flex-end;
        }

        .input-wrapper textarea {
          flex: 1;
          padding: 12px 16px;
          border-radius: 24px;
          border: 1px solid var(--border-color, #333);
          background: var(--bg-primary, #1e1e1e);
          color: var(--text-primary, #fff);
          resize: none;
          min-height: 44px;
          max-height: 120px;
          font-family: inherit;
          font-size: 14px;
        }

        .input-wrapper textarea:focus {
          outline: none;
          border-color: #0078d4;
        }

        .send-btn {
          width: 44px;
          height: 44px;
          border-radius: 50%;
          border: none;
          background: #0078d4;
          color: white;
          cursor: pointer;
          display: flex;
          align-items: center;
          justify-content: center;
          transition: all 0.2s;
        }

        .send-btn:hover:not(:disabled) {
          background: #106ebe;
        }

        .send-btn:disabled {
          opacity: 0.5;
          cursor: not-allowed;
        }

        .input-hint {
          margin-top: 8px;
          font-size: 11px;
          color: var(--text-secondary, #888);
          text-align: center;
        }
      `}</style>
    </div>
  );
};

// Helper function to format message content with markdown-like syntax
function formatMessage(content: string): string {
  return content
    .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
    .replace(/\*(.*?)\*/g, '<em>$1</em>')
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    .replace(/\n/g, '<br />');
}

export default AgentChatPanel;
