import React, { useState, useRef, useEffect } from 'react';
import { Button, Input } from '@fluentui/react-components';
import { Send24Regular } from '@fluentui/react-icons';
import { useAppStore } from '../store/appStore';
import { aiService } from '../services/aiService';
import { outlookService } from '../services/outlookService';
import { graphService } from '../services/graphService';
import { documentService, DocumentType, TemplateType } from '../services/documentService';
import { v4 as uuidv4 } from 'uuid';

const ChatPanel: React.FC = () => {
  const [input, setInput] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [isRecording, setIsRecording] = useState(false);
  const [showSuggestions, setShowSuggestions] = useState(false);
  const [isSignedIn, setIsSignedIn] = useState(graphService.isSignedIn);
  const [userName, setUserName] = useState<string | null>(null);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const { messages, addMessage, currentEmail, tasks } = useAppStore();

  // Smart suggestions based on context
  const suggestions = [
    currentEmail ? `Summarize email from ${currentEmail.sender}` : null,
    tasks.filter(t => t.status !== 'completed').length > 0 ? 'What are my top priority tasks?' : null,
    'What meetings do I have today?',
    'Help me write a follow-up email',
    'Create a task from my last email',
  ].filter(Boolean) as string[];

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  };

  useEffect(() => {
    scrollToBottom();
  }, [messages]);

  // Check sign-in status on mount
  useEffect(() => {
    setIsSignedIn(graphService.isSignedIn);
    if (graphService.isSignedIn) {
      const userInfo = graphService.getUserInfo();
      setUserName(userInfo?.name || null);
    }
  }, []);

  // Handle Microsoft sign-in
  const handleMicrosoftSignIn = async () => {
    try {
      const success = await graphService.signIn();
      if (success) {
        setIsSignedIn(true);
        const userInfo = graphService.getUserInfo();
        setUserName(userInfo?.name || null);
        addMessage({
          id: uuidv4(),
          role: 'assistant',
          content: `✅ **Signed in successfully!**\n\nWelcome ${userInfo?.name || 'there'}! You now have full inbox access.\n\nTry:\n• "Review my entire inbox"\n• "Show me my unread emails"\n• "Search emails from [person]"`,
          timestamp: new Date(),
        });
      }
    } catch (error: any) {
      addMessage({
        id: uuidv4(),
        role: 'assistant',
        content: `❌ **Sign-in failed**\n\n${error.message || 'Please try again.'}\n\nMake sure you've approved the permissions in the popup.`,
        timestamp: new Date(),
      });
    }
  };

  // Handle sign out
  const handleSignOut = async () => {
    await graphService.signOut();
    setIsSignedIn(false);
    setUserName(null);
    addMessage({
      id: uuidv4(),
      role: 'assistant',
      content: '👋 Signed out. You can still work with the currently selected email.',
      timestamp: new Date(),
    });
  };

  const quickActions = [
    { icon: '�', label: 'Review Inbox', action: 'review-inbox' },
    { icon: '📅', label: 'My Calendar', action: 'plan-day' },
    { icon: '📋', label: 'My Tasks', action: 'show-tasks' },
    { icon: '🔍', label: 'Search Emails', action: 'search-emails' },
    { icon: '✉️', label: 'Compose Email', action: 'compose' },
    { icon: '📝', label: 'Summarize Email', action: 'summarize' },
  ];

  const handleQuickAction = async (action: string) => {
    let prompt = '';
    
    switch (action) {
      case 'summarize':
        if (currentEmail) {
          prompt = `Please summarize this email from ${currentEmail.sender}: "${currentEmail.subject}"`;
        } else {
          prompt = 'Review my recent emails and summarize the most important ones';
        }
        break;
      case 'reply':
        prompt = 'Help me draft a professional reply to this email';
        break;
      case 'extract-tasks':
        prompt = 'Extract any action items or tasks from my recent emails';
        break;
      case 'plan-day':
        prompt = 'Show me my calendar for today and help me plan my day based on meetings and pending tasks';
        break;
      case 'show-tasks':
        prompt = 'Show me all my current tasks and help me prioritize them';
        break;
      case 'compose':
        prompt = 'Help me compose a new email. What would you like to write about?';
        addMessage({
          id: uuidv4(),
          role: 'assistant',
          content: '✉️ I can help you compose an email! Just tell me:\n\n• **Who** do you want to email?\n• **What** is it about?\n• **What tone** (formal, casual, friendly)?\n\nOr just describe what you need and I\'ll draft something for you.',
          timestamp: new Date(),
        });
        return;
      case 'review-inbox':
        prompt = 'Review my entire inbox and give me a summary of important emails, unread messages, and action items';
        break;
      case 'search-emails':
        prompt = 'What would you like to search for in your inbox?';
        addMessage({
          id: uuidv4(),
          role: 'assistant',
          content: prompt,
          timestamp: new Date(),
        });
        return;
    }
    
    if (prompt) {
      setInput(prompt);
      await handleSend(prompt);
    }
  };

  const handleSend = async (messageText?: string) => {
    const text = messageText || input.trim();
    if (!text || isProcessing) return;

    // Add user message
    addMessage({
      id: uuidv4(),
      role: 'user',
      content: text,
      timestamp: new Date(),
    });

    setInput('');
    setIsProcessing(true);

    try {
      // Detect action requests - be specific to avoid false positives
      const wantsMarkAsRead = /mark.*read|mark.*as read|mark all.*read|mark them.*read/i.test(text);
      const wantsReviewUnread = /review.*unread|unread email|unopen email/i.test(text);
      // Only trigger inbox mode for explicit inbox scanning requests, not general questions
      const isInboxScanRequest = /^(review|scan|check|show|list).*(inbox|all.*email|entire|unread)/i.test(text) || 
                                  /^(what|how many).*(unread|inbox)/i.test(text);
      const isSearchRequest = /search|find|look for|emails from|emails about/i.test(text);
      const isSummarizeRequest = /summarize|summary|tldr|what.*about|key points/i.test(text);
      const isReplyRequest = /reply|respond|draft|write back|answer/i.test(text);
      const isTaskRequest = /action item|task|todo|to-do|extract.*task|create.*task/i.test(text);
      
      console.log('Request detection:', { wantsMarkAsRead, isInboxScanRequest, isSummarizeRequest, isReplyRequest });
      
      // Get current email context - THIS WORKS without Graph API
      const emailContext = await outlookService.getCurrentEmailContext();
      const calendarContext = await outlookService.getUpcomingEvents(7);
      
      console.log('Email context:', emailContext ? `Got email: ${emailContext.subject}` : 'No email selected');
      console.log('Graph signed in:', graphService.isSignedIn);

      // Handle explicit inbox-wide scanning requests
      if (isInboxScanRequest || wantsReviewUnread) {
        // If signed in to Microsoft, use Graph API for real inbox access
        if (graphService.isSignedIn) {
          try {
            addMessage({
              id: uuidv4(),
              role: 'assistant',
              content: '🔍 Scanning your inbox...',
              timestamp: new Date(),
            });
            
            const inboxSummary = await graphService.getInboxSummary();
            
            if (inboxSummary.isRealData) {
              const topSendersList = inboxSummary.topSenders.slice(0, 5)
                .map(s => `• ${s.name} (${s.count} emails)`)
                .join('\n');
              
              const recentEmailsList = inboxSummary.recentEmails.slice(0, 5)
                .map(e => `• ${e.isRead ? '📖' : '📬'} **${e.subject}** - from ${e.sender}`)
                .join('\n');
              
              addMessage({
                id: uuidv4(),
                role: 'assistant',
                content: `📊 **Inbox Summary**\n\n` +
                  `📬 **${inboxSummary.unreadCount}** unread emails\n` +
                  `📧 **${inboxSummary.totalEmails}** total (last 100)\n\n` +
                  `**Top Senders:**\n${topSendersList}\n\n` +
                  `**Recent Emails:**\n${recentEmailsList}\n\n` +
                  `Would you like me to summarize specific emails or filter by sender?`,
                timestamp: new Date(),
              });
              setIsProcessing(false);
              return;
            }
          } catch (error) {
            console.error('Graph API error:', error);
          }
        }
        
        // Fallback to current email context
        if (emailContext) {
          addMessage({
            id: uuidv4(),
            role: 'assistant',
            content: `📧 **I can help with your current email!**\n\n*Sign in with Microsoft (button above) for full inbox access.*\n\n**Current Email:**\n- **From:** ${emailContext.sender}\n- **Subject:** ${emailContext.subject || '(no subject)'}\n- **Status:** ${emailContext.isRead ? 'Read' : '📬 Unread'}\n\n**What would you like me to do?**\n• "Summarize this email"\n• "Draft a reply"\n• "Extract action items"`,
            timestamp: new Date(),
          });
        } else {
          addMessage({
            id: uuidv4(),
            role: 'assistant',
            content: `📭 **Full Inbox Access**\n\nTo scan your entire inbox:\n1. Click **"Sign in for full inbox access"** above\n2. Approve the permissions\n3. Ask me to "review my inbox" again\n\n*Or select an email to work with it directly.*`,
            timestamp: new Date(),
          });
        }
        setIsProcessing(false);
        return;
      }

      // Handle mark as read request
      if (wantsMarkAsRead) {
        if (emailContext) {
          addMessage({
            id: uuidv4(),
            role: 'assistant',
            content: `📧 **Current email:** "${emailContext.subject}"\n\n${emailContext.isRead ? '✅ This email is already marked as read.' : '📬 This email is unread.'}\n\n**To mark as read:** Simply click on the email in Outlook, or right-click and select "Mark as Read".\n\n**To mark ALL as read:** Right-click on your Inbox folder → "Mark all as read"`,
            timestamp: new Date(),
          });
        } else {
          addMessage({
            id: uuidv4(),
            role: 'assistant',
            content: `📭 **No email selected**\n\nTo mark emails as read:\n1. **Single email:** Click on it or right-click → "Mark as Read"\n2. **All emails:** Right-click on Inbox folder → "Mark all as read"`,
            timestamp: new Date(),
          });
        }
        setIsProcessing(false);
        return;
      }
      
      // Build enhanced context for AI - always include current email if available
      let enhancedPrompt = text;
      
      if (emailContext) {
        // We have current email context - include it for AI to work with
        enhancedPrompt = `USER REQUEST: ${text}

CURRENT EMAIL CONTEXT:
- From: ${emailContext.sender} (${emailContext.senderEmail || 'no email'})
- Subject: ${emailContext.subject}
- Received: ${emailContext.receivedDateTime}
- Status: ${emailContext.isRead ? 'Read' : 'Unread'}
- Has Attachments: ${emailContext.hasAttachments ? 'Yes' : 'No'}
- Content: ${emailContext.preview || 'No content preview available'}

Please help the user with their request. If they want a summary, provide a clear summary. If they want action items, extract them. If they want to reply, draft a professional response.`;
      }
      
      // Send to AI for analysis/response
      const response = await aiService.chat({
        prompt: enhancedPrompt,
        context: {
          currentEmail: emailContext || undefined,
          upcomingEvents: calendarContext,
        },
      });

      // Add assistant response
      addMessage({
        id: uuidv4(),
        role: 'assistant',
        content: response.content,
        timestamp: new Date(),
        metadata: {
          action: response.suggestedActions?.[0]?.type,
        },
      });

      // Handle any suggested actions
      if (response.extractedTasks && response.extractedTasks.length > 0) {
        // Could auto-create tasks here based on user settings
      }

    } catch (error) {
      console.error('ChatPanel error:', error);
      addMessage({
        id: uuidv4(),
        role: 'assistant',
        content: 'Sorry, I encountered an error processing your request. Please try again.',
        timestamp: new Date(),
      });
    } finally {
      setIsProcessing(false);
    }
  };

  const handleKeyPress = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  const handleVoiceInput = () => {
    if (!('webkitSpeechRecognition' in window) && !('SpeechRecognition' in window)) {
      alert('Voice input is not supported in this browser.');
      return;
    }

    const SpeechRecognition = (window as any).SpeechRecognition || (window as any).webkitSpeechRecognition;
    const recognition = new SpeechRecognition();
    
    recognition.continuous = false;
    recognition.interimResults = false;
    recognition.lang = 'en-US';

    recognition.onstart = () => {
      setIsRecording(true);
    };

    recognition.onresult = (event: any) => {
      const transcript = event.results[0][0].transcript;
      setInput(transcript);
      setIsRecording(false);
    };

    recognition.onerror = () => {
      setIsRecording(false);
    };

    recognition.onend = () => {
      setIsRecording(false);
    };

    if (isRecording) {
      recognition.stop();
    } else {
      recognition.start();
    }
  };

  const handleSuggestionClick = (suggestion: string) => {
    setInput(suggestion);
    setShowSuggestions(false);
    handleSend(suggestion);
  };

  // Handle document export with auto-template detection
  const handleExport = async (type: DocumentType, content: string) => {
    try {
      const detectedTemplate = documentService.detectTemplate(content);
      const templateInfo = documentService.templates.find(t => t.id === detectedTemplate);
      const title = `${templateInfo?.name || 'AI Response'} - ${new Date().toLocaleDateString()}`;
      
      // Show loading state
      addMessage({
        id: uuidv4(),
        role: 'assistant',
        content: `⏳ Generating ${type.toUpperCase()} document...`,
        timestamp: new Date(),
      });
      
      const options = { template: detectedTemplate, includeCharts: true };
      
      switch (type) {
        case 'word':
          await documentService.createWord(title, content, options);
          break;
        case 'pdf':
          await documentService.createPDF(title, content, options);
          break;
        case 'excel':
          await documentService.createExcel(title, content, options);
          break;
        case 'powerpoint':
          await documentService.createPowerPoint(title, content, options);
          break;
      }
      
      // Success message with template info
      addMessage({
        id: uuidv4(),
        role: 'assistant',
        content: `✅ **${type.toUpperCase()} exported!** ${templateInfo?.icon || '📄'}\n\nUsed template: **${templateInfo?.name || 'Custom'}**\n\n${type === 'pdf' || type === 'powerpoint' ? '💡 Tip: Use Ctrl+P to save as PDF or F11 for fullscreen presentation.' : 'Check your downloads folder.'}`,
        timestamp: new Date(),
      });
    } catch (error: any) {
      console.error('Export error:', error);
      addMessage({
        id: uuidv4(),
        role: 'assistant',
        content: `❌ Export failed: ${error.message}`,
        timestamp: new Date(),
      });
    }
  };

  return (
    <div className="chat-container">
      {/* Microsoft Sign-in for Full Inbox Access */}
      <div className="microsoft-auth-bar">
        {!isSignedIn ? (
          <button 
            className="microsoft-signin-btn"
            onClick={handleMicrosoftSignIn}
          >
            <svg width="16" height="16" viewBox="0 0 21 21" xmlns="http://www.w3.org/2000/svg" style={{ marginRight: '8px' }}>
              <rect x="1" y="1" width="9" height="9" fill="#f25022"/>
              <rect x="1" y="11" width="9" height="9" fill="#00a4ef"/>
              <rect x="11" y="1" width="9" height="9" fill="#7fba00"/>
              <rect x="11" y="11" width="9" height="9" fill="#ffb900"/>
            </svg>
            Sign in for full inbox access
          </button>
        ) : (
          <div className="signed-in-info">
            <span className="user-badge">✓ {userName || 'Signed in'}</span>
            <button className="signout-btn" onClick={handleSignOut}>Sign out</button>
          </div>
        )}
      </div>

      {/* Quick Actions */}
      <div className="quick-actions">
        {quickActions.map((action) => (
          <button
            key={action.action}
            className="quick-action-btn"
            onClick={() => handleQuickAction(action.action)}
            disabled={isProcessing}
          >
            <div className="quick-action-icon">{action.icon}</div>
            <div className="quick-action-label">{action.label}</div>
          </button>
        ))}
      </div>

      {/* Chat Messages */}
      <div className="chat-messages">
        {messages.length === 0 ? (
          <div className="empty-state">
            <div className="empty-state-icon">🤖</div>
            <p className="empty-state-text">
              <strong>Hi! I'm your AI Assistant.</strong>
            </p>
            <p className="empty-state-text" style={{ fontSize: '13px', marginTop: '8px' }}>
              I'm always here to help you with:
            </p>
            <ul style={{ fontSize: '12px', textAlign: 'left', margin: '8px 0', paddingLeft: '20px', color: '#555' }}>
              <li>📥 Review your entire inbox</li>
              <li>📅 Manage your calendar & meetings</li>
              <li>📋 Organize tasks & to-dos</li>
              <li>✉️ Draft & send emails</li>
              <li>🔍 Search through your mailbox</li>
            </ul>
            <p className="empty-state-text" style={{ fontSize: '12px', marginTop: '8px' }}>
              Use the quick actions above or just ask me anything!
            </p>
          </div>
        ) : (
          messages.map((message) => (
            <div
              key={message.id}
              className={`chat-message ${message.role}`}
            >
              <div className="message-content">{message.content}</div>
              {/* Export buttons for AI responses */}
              {message.role === 'assistant' && message.content.length > 50 && (
                <div className="message-export-buttons">
                  <button
                    className="export-btn"
                    onClick={() => handleExport('word', message.content)}
                    title="Export to Word"
                  >
                    📄 Word
                  </button>
                  <button
                    className="export-btn"
                    onClick={() => handleExport('pdf', message.content)}
                    title="Export to PDF"
                  >
                    📕 PDF
                  </button>
                  <button
                    className="export-btn"
                    onClick={() => handleExport('excel', message.content)}
                    title="Export to Excel"
                  >
                    📊 Excel
                  </button>
                  <button
                    className="export-btn"
                    onClick={() => handleExport('powerpoint', message.content)}
                    title="Export to PowerPoint"
                  >
                    📽️ PPT
                  </button>
                </div>
              )}
            </div>
          ))
        )}
        {isProcessing && (
          <div className="chat-message assistant">
            <div className="spinner spinner--sm" />
          </div>
        )}
        <div ref={messagesEndRef} />
      </div>

      {/* Chat Input */}
      <div className="chat-input-container" style={{ position: 'relative' }}>
        {/* Suggestions Dropdown */}
        {showSuggestions && suggestions.length > 0 && (
          <div className="compose-suggestions">
            {suggestions.map((suggestion, index) => (
              <div 
                key={index}
                className="suggestion-item"
                onClick={() => handleSuggestionClick(suggestion)}
              >
                💡 {suggestion}
              </div>
            ))}
          </div>
        )}
        
        <input
          className="chat-input"
          type="text"
          placeholder="Ask me anything..."
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyPress={handleKeyPress}
          disabled={isProcessing}
        />
        <button
          className={`voice-btn ${isRecording ? 'recording' : ''}`}
          onClick={handleVoiceInput}
          title={isRecording ? 'Stop recording' : 'Voice input'}
        >
          {isRecording ? '⏹️' : '🎤'}
        </button>
        <button
          className="send-button"
          onClick={() => handleSend()}
          disabled={!input.trim() || isProcessing}
        >
          ➤
        </button>
      </div>
    </div>
  );
};

export default ChatPanel;
