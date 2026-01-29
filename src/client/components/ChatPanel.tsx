import React, { useState, useRef, useEffect } from 'react';
import { Button, Input, SendIcon } from './ui/NativeComponents';
import { useAppStore } from '../store/appStore';
import { aiService } from '../services/aiService';
import { outlookService } from '../services/outlookService';
import { brandingService, UserBranding } from '../services/brandingService';
import { agentService } from '../services/agentService';
import BrandingPanel from './BrandingPanel';
import { v4 as uuidv4 } from 'uuid';

// Lazy load graph service to avoid bundling MSAL
const loadGraphService = () => import(/* webpackChunkName: "graph-service" */ '../services/graphService');

// Types defined inline to avoid importing documentService at build time
type DocumentType = 'word' | 'pdf' | 'excel' | 'powerpoint';
type TemplateType = 'professional-report' | 'meeting-summary' | 'project-status' | 'data-analysis' | 'sales-pitch' | 'email-summary' | 'action-items' | 'custom';

// Lazy load document service (reduces initial bundle by ~2MB)
const loadDocumentService = () => import(/* webpackChunkName: "document-service" */ '../services/documentService');

const ChatPanel: React.FC = () => {
  const [input, setInput] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [isRecording, setIsRecording] = useState(false);
  const [showSuggestions, setShowSuggestions] = useState(false);
  const [isSignedIn, setIsSignedIn] = useState(false); // Default false, will check on mount
  const [userName, setUserName] = useState<string | null>(null);
  const [showBranding, setShowBranding] = useState(false);
  const [userBranding, setUserBranding] = useState<UserBranding>(brandingService.load());
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

  // Check sign-in status on mount (lazy load graphService)
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
    const { graphService } = await loadGraphService();
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
    { icon: '📥', label: 'Review Inbox', action: 'review-inbox' },
    { icon: '📅', label: 'My Calendar', action: 'plan-day' },
    { icon: '📋', label: 'My Tasks', action: 'show-tasks' },
    { icon: '🔍', label: 'Search Emails', action: 'search-emails' },
    { icon: '✉️', label: 'Compose Email', action: 'compose' },
    { icon: '🎨', label: 'My Branding', action: 'branding' },
  ];

  const handleQuickAction = async (action: string) => {
    let prompt = '';
    
    switch (action) {
      case 'branding':
        setShowBranding(true);
        return;
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

    // Also add to agent context
    agentService.addToHistory({
      role: 'user',
      content: text,
      timestamp: new Date(),
    });

    setInput('');
    setIsProcessing(true);

    try {
      const { graphService } = await loadGraphService();
      
      // Detect user intent using the intelligent agent
      const intent = agentService.detectIntent(text);
      console.log('🧠 Detected intent:', intent);

      // Get current email context
      const emailContext = await outlookService.getCurrentEmailContext();
      
      // ========================================
      // INTELLIGENT ACTION EXECUTION
      // ========================================
      
      // Handle follow-up (user said "please" or "yes" after being asked)
      if (intent.type === 'follow_up') {
        const context = agentService.getContext();
        if (context.pendingAction === 'search' && context.lastSearchQuery) {
          // User confirmed they want to search
          intent.type = 'search';
          intent.entities.searchTerms = [context.lastSearchQuery];
          console.log('🔄 Following up on search:', context.lastSearchQuery);
        }
      }

      // SEARCH: Actually search the inbox
      if (intent.type === 'search' && graphService.isSignedIn) {
        const searchQueries = agentService.extractSearchQuery(text, intent.entities);
        console.log('🔍 Search queries extracted:', searchQueries);

        if (searchQueries.length > 0) {
          addMessage({
            id: uuidv4(),
            role: 'assistant',
            content: `🔍 Searching your inbox for: **${searchQueries.join(', ')}**...`,
            timestamp: new Date(),
          });

          // Actually search the inbox!
          let allResults: any[] = [];
          for (const query of searchQueries) {
            const results = await graphService.searchEmails(query, 25);
            allResults = [...allResults, ...results];
          }
          
          // Deduplicate by email ID
          const uniqueResults = Array.from(
            new Map(allResults.map(r => [r.id, r])).values()
          );
          
          // Store results in agent context
          agentService.setSearchResults(uniqueResults, searchQueries.join(', '));
          agentService.setPendingAction(null);

          if (uniqueResults.length === 0) {
            const responseContent = `📭 **No emails found** matching "${searchQueries.join(', ')}"\n\n` +
              `I searched your inbox but didn't find any matching emails. Try:\n` +
              `• Different search terms\n` +
              `• Company name instead of product\n` +
              `• Sender's name or email address`;
            
            addMessage({
              id: uuidv4(),
              role: 'assistant',
              content: responseContent,
              timestamp: new Date(),
            });
            
            agentService.addToHistory({
              role: 'assistant',
              content: responseContent,
              timestamp: new Date(),
              metadata: { searchQuery: searchQueries.join(', '), searchResults: [] }
            });
          } else {
            // Format results for display
            const emailList = uniqueResults.slice(0, 10).map((e: any, i: number) => 
              `${i + 1}. **${e.sender}** <${e.senderEmail}>\n   📧 ${e.subject}\n   📅 ${new Date(e.receivedDateTime).toLocaleDateString()}`
            ).join('\n\n');
            
            // Extract unique senders for compose help
            const uniqueSenders = Array.from(
              new Map(uniqueResults.map((r: any) => [r.senderEmail, { name: r.sender, email: r.senderEmail }])).values()
            );
            
            const responseContent = `📬 **Found ${uniqueResults.length} emails** matching "${searchQueries.join(', ')}"\n\n` +
              `${emailList}\n\n` +
              `---\n` +
              `**What would you like to do?**\n` +
              `• "Summarize these emails"\n` +
              `• "Draft an email to [sender name]"\n` +
              `• "What action items are in these emails?"`;
            
            addMessage({
              id: uuidv4(),
              role: 'assistant',
              content: responseContent,
              timestamp: new Date(),
            });
            
            agentService.addToHistory({
              role: 'assistant',
              content: responseContent,
              timestamp: new Date(),
              metadata: { searchQuery: searchQueries.join(', '), searchResults: uniqueResults }
            });
          }
          
          setIsProcessing(false);
          return;
        }
      }

      // COMPOSE: Help compose emails using inbox context
      if (intent.type === 'compose') {
        const context = agentService.getContext();
        
        // Check if we have contacts from previous searches
        if (context.lastSearchResults.length > 0) {
          const uniqueSenders = Array.from(
            new Map(context.lastSearchResults.map((r: any) => [r.senderEmail, { name: r.sender, email: r.senderEmail }])).values()
          );
          
          // Build AI prompt with context
          const contextStr = agentService.buildAIContext(emailContext);
          const enhancedPrompt = `USER REQUEST: ${text}

${contextStr}

The user wants to compose an email. Based on the conversation context and search results above:
1. If they mentioned specific contacts/companies, use those email addresses
2. If they want pricing information, draft a professional pricing request
3. Include all relevant contacts from the search results
4. Be specific and professional

AVAILABLE CONTACTS FROM THEIR INBOX:
${uniqueSenders.slice(0, 10).map((s: any) => `- ${s.name}: ${s.email}`).join('\n')}

Draft the email(s) they requested. For each email, provide:
- TO: [email address]
- SUBJECT: [subject line]
- BODY: [full email body]`;

          const response = await aiService.chat({
            prompt: enhancedPrompt,
            context: { currentEmail: emailContext || undefined },
          });

          addMessage({
            id: uuidv4(),
            role: 'assistant',
            content: response.content,
            timestamp: new Date(),
          });
          
          agentService.addToHistory({
            role: 'assistant',
            content: response.content,
            timestamp: new Date(),
          });
          
          setIsProcessing(false);
          return;
        }
        
        // No context - need to search first
        if (graphService.isSignedIn) {
          // Try to extract what they want to compose about
          const searchTerms = agentService.extractSearchQuery(text, intent.entities);
          
          if (searchTerms.length > 0) {
            addMessage({
              id: uuidv4(),
              role: 'assistant',
              content: `🔍 Let me find relevant contacts in your inbox first...`,
              timestamp: new Date(),
            });
            
            let allResults: any[] = [];
            for (const query of searchTerms) {
              const results = await graphService.searchEmails(query, 20);
              allResults = [...allResults, ...results];
            }
            
            if (allResults.length > 0) {
              agentService.setSearchResults(allResults, searchTerms.join(', '));
              
              const uniqueSenders = Array.from(
                new Map(allResults.map((r: any) => [r.senderEmail, { name: r.sender, email: r.senderEmail }])).values()
              );
              
              const senderList = uniqueSenders.slice(0, 10).map((s: any, i: number) => 
                `${i + 1}. **${s.name}** - ${s.email}`
              ).join('\n');
              
              const responseContent = `📬 **Found ${uniqueSenders.length} contacts** related to "${searchTerms.join(', ')}":\n\n` +
                `${senderList}\n\n` +
                `Would you like me to draft emails to all of them, or specific ones?`;
              
              addMessage({
                id: uuidv4(),
                role: 'assistant',
                content: responseContent,
                timestamp: new Date(),
              });
              
              agentService.addToHistory({
                role: 'assistant',
                content: responseContent,
                timestamp: new Date(),
                metadata: { searchResults: allResults }
              });
              
              agentService.setPendingAction('compose');
              setIsProcessing(false);
              return;
            }
          }
        }
      }

      // LIST/REVIEW: Show inbox summary
      if (intent.type === 'list' && graphService.isSignedIn) {
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
          
          const responseContent = `📊 **Inbox Summary**\n\n` +
            `📬 **${inboxSummary.unreadCount}** unread emails\n` +
            `📧 **${inboxSummary.totalEmails}** total (last 100)\n\n` +
            `**Top Senders:**\n${topSendersList}\n\n` +
            `**Recent Emails:**\n${recentEmailsList}\n\n` +
            `What would you like to do? You can:\n` +
            `• Search for specific emails\n` +
            `• Summarize emails from a sender\n` +
            `• Find action items`;
          
          addMessage({
            id: uuidv4(),
            role: 'assistant',
            content: responseContent,
            timestamp: new Date(),
          });
          
          agentService.addToHistory({
            role: 'assistant',
            content: responseContent,
            timestamp: new Date(),
          });
          
          setIsProcessing(false);
          return;
        }
      }

      // ========================================
      // FALLBACK: Send to AI with full context
      // ========================================
      
      // Build comprehensive context for AI
      const agentContext = agentService.buildAIContext(emailContext);
      const hasSearchResults = agentService.getContext().lastSearchResults.length > 0;
      
      const systemInstructions = hasSearchResults 
        ? `You have access to the user's inbox search results. Use them to answer questions.
If they ask about emails, contacts, or want to compose - USE THE SEARCH RESULTS PROVIDED.
Don't ask them to search - YOU have the results already.
Be proactive and helpful.`
        : `You are a helpful email assistant. If the user asks about finding or searching emails, and you don't have search results, tell them you'll search for it and then ACTUALLY describe what you found (the system will search automatically).`;

      const enhancedPrompt = `${systemInstructions}

USER REQUEST: ${text}

${agentContext}

${emailContext ? `
CURRENTLY SELECTED EMAIL:
- From: ${emailContext.sender} (${emailContext.senderEmail || 'no email'})
- Subject: ${emailContext.subject}
- Preview: ${emailContext.preview || 'No preview'}
` : ''}

Respond helpfully. If you have search results above, USE THEM to answer. Don't ask the user to search - you already have the data.`;

      const response = await aiService.chat({
        prompt: enhancedPrompt,
        context: { currentEmail: emailContext || undefined },
      });

      addMessage({
        id: uuidv4(),
        role: 'assistant',
        content: response.content,
        timestamp: new Date(),
      });
      
      agentService.addToHistory({
        role: 'assistant',
        content: response.content,
        timestamp: new Date(),
      });

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
      // Lazy load document service when needed
      const { documentService } = await loadDocumentService();
      
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

  // Handle branding save
  const handleBrandingSave = (branding: UserBranding) => {
    setUserBranding(branding);
    setShowBranding(false);
    addMessage({
      id: uuidv4(),
      role: 'assistant',
      content: `✅ **Branding saved!** 🎨\n\nYour custom branding will now be used in all document exports:\n• Company: **${branding.companyName || 'Not set'}**\n• Colors: Primary ${branding.primaryColor}\n• Logo: ${branding.logoUrl ? '✓ Uploaded' : 'Not set'}\n\nTry exporting a document to see your branding in action!`,
      timestamp: new Date(),
    });
  };

  return (
    <div className="chat-container">
      {/* Branding Panel Modal */}
      {showBranding && (
        <BrandingPanel
          onClose={() => setShowBranding(false)}
          onSave={handleBrandingSave}
        />
      )}
      
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
