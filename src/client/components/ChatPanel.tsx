import React, { useState, useRef, useEffect } from 'react';
import { Button, Input, SendIcon } from './ui/NativeComponents';
import { useAppStore } from '../store/appStore';
import { aiService } from '../services/aiService';
import { outlookService } from '../services/outlookService';
import { brandingService, UserBranding } from '../services/brandingService';
import { autonomousAgent } from '../services/autonomousAgent';
import BrandingPanel from './BrandingPanel';
import { v4 as uuidv4 } from 'uuid';

// Lazy load graph service to avoid bundling MSAL
const loadGraphService = () => import(/* webpackChunkName: "graph-service" */ '../services/graphService');

// Lazy load action executor
const loadActionExecutor = () => import(/* webpackChunkName: "action-executor" */ '../services/actionExecutor');

// Types defined inline to avoid importing documentService at build time
type DocumentType = 'word' | 'pdf' | 'excel' | 'powerpoint';
type TemplateType = 'professional-report' | 'meeting-summary' | 'project-status' | 'data-analysis' | 'sales-pitch' | 'email-summary' | 'action-items' | 'custom';

// Lazy load document service (reduces initial bundle by ~2MB)
const loadDocumentService = () => import(/* webpackChunkName: "document-service" */ '../services/documentService');

// Pending approval state
interface PendingApproval {
  id: string;
  type: string;
  description: string;
  details: any;
}

const ChatPanel: React.FC = () => {
  const [input, setInput] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [isRecording, setIsRecording] = useState(false);
  const [showSuggestions, setShowSuggestions] = useState(false);
  const [isSignedIn, setIsSignedIn] = useState(false);
  const [userName, setUserName] = useState<string | null>(null);
  const [showBranding, setShowBranding] = useState(false);
  const [userBranding, setUserBranding] = useState<UserBranding>(brandingService.load());
  const [pendingApprovals, setPendingApprovals] = useState<PendingApproval[]>([]);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const { messages, addMessage, currentEmail, tasks } = useAppStore();
  const showFrameworkMessages = false;

  const addFrameworkMessage = (content: string) => {
    if (!showFrameworkMessages) return;
    addMessage({
      id: uuidv4(),
      role: 'assistant',
      content,
      timestamp: new Date(),
    });
  };

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
      await graphService.initializeAuth();
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
        addFrameworkMessage(
          `✅ **Signed in successfully!**\n\nWelcome ${userInfo?.name || 'there'}! You now have full inbox access.\n\nTry:\n• "Review my entire inbox"\n• "Show me my unread emails"\n• "Search emails from [person]"`
        );
      }
    } catch (error: any) {
      addFrameworkMessage(
        `❌ **Sign-in failed**\n\n${error.message || 'Please try again.'}\n\nMake sure you've approved the permissions in the popup.`
      );
    }
  };

  // Handle sign out
  const handleSignOut = async () => {
    const { graphService } = await loadGraphService();
    await graphService.signOut();
    setIsSignedIn(false);
    setUserName(null);
    addFrameworkMessage('👋 Signed out. You can still work with the currently selected email.');
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
        addFrameworkMessage(
          '✉️ I can help you compose an email! Just tell me:\n\n• **Who** do you want to email?\n• **What** is it about?\n• **What tone** (formal, casual, friendly)?\n\nOr just describe what you need and I\'ll draft something for you.'
        );
        return;
      case 'review-inbox':
        prompt = 'Review my entire inbox and give me a summary of important emails, unread messages, and action items';
        break;
      case 'search-emails':
        prompt = 'What would you like to search for in your inbox?';
        addFrameworkMessage(prompt);
        return;
    }
    
    if (prompt) {
      setInput(prompt);
      await handleSend(prompt);
    }
  };

  // Handle approval of pending actions
  const handleApproval = async (approvalId: string, approved: boolean) => {
    const approval = pendingApprovals.find(a => a.id === approvalId);
    if (!approval) return;

    // Remove from pending
    setPendingApprovals(prev => prev.filter(a => a.id !== approvalId));

    if (!approved) {
      addFrameworkMessage(`❌ Cancelled: ${approval.description}`);
      return;
    }

    // Execute the approved action
    setIsProcessing(true);
    try {
      const { graphService } = await loadGraphService();
      
      if (approval.type === 'send_email') {
        // Actually send the email via Graph API
        const success = await graphService.sendEmail(
          approval.details.to,
          approval.details.subject,
          approval.details.body
        );
        
        if (success) {
          addFrameworkMessage(
            `✅ **Email sent successfully!**\n\n📤 To: ${approval.details.to}\n📋 Subject: ${approval.details.subject}`
          );
        } else {
          addFrameworkMessage('❌ Failed to send email. Please try again.');
        }
      } else if (approval.type === 'create_calendar_event') {
        addFrameworkMessage(`✅ **Calendar event created!**\n\n📅 ${approval.details.title}`);
      } else if (approval.type === 'delete_email') {
        addFrameworkMessage('✅ **Email deleted.**');
      }
    } catch (error) {
      console.error('Error executing approved action:', error);
      addFrameworkMessage(`❌ Error: Could not complete the action. ${error}`);
    } finally {
      setIsProcessing(false);
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

    // Add to agent conversation
    autonomousAgent.addTurn({
      role: 'user',
      content: text,
      timestamp: new Date(),
    });

    setInput('');
    setIsProcessing(true);

    try {
      const { graphService } = await loadGraphService();
      
      // Get current email context
      const emailContext = await outlookService.getCurrentEmailContext();
      
      // ========================================
      // PHASE 1: AUTO-EXECUTE SAFE ACTIONS
      // ========================================
      
      // Detect if we need to search (the agent will decide, but we can pre-fetch)
      const lowerText = text.toLowerCase();
      
      // Check for specific search targets in the message
      const hasSearchTarget = /activated carbon|patent|distributor|supplier|vendor|pricing|quote/i.test(lowerText);
      const needsSearch = hasSearchTarget || /search|find|look for|any.*from|emails? (from|about)|do i have/i.test(lowerText);
      const needsInboxSummary = /review.*inbox|inbox.*summary|show.*inbox|scan.*inbox/i.test(lowerText) && !hasSearchTarget;
      const needsUnreadCount = /unread|how many.*emails?|count.*emails?/i.test(lowerText);
      
      let searchResults: any[] = [];
      let inboxSummary: any = null;
      
      // Auto-fetch unread emails if asking about unread count
      if (needsUnreadCount && graphService.isSignedIn) {
        addFrameworkMessage('📊 Checking your unread emails...');
        
        try {
          // Get unread emails using the dedicated method
          const unreadEmails = await graphService.getUnreadEmails(100);
          console.log(`📬 Found ${unreadEmails.length} unread emails`);
          
          // Set context for AI
          autonomousAgent.setSearchResults(unreadEmails);
          searchResults = unreadEmails;
          
          // Provide immediate answer
          const unreadSummary = unreadEmails.slice(0, 5).map((e: any, i: number) => 
            `${i + 1}. **${e.sender}**: ${e.subject}`
          ).join('\n');
          
          addFrameworkMessage(
            `📬 **You have ${unreadEmails.length} unread emails**\n\n**Most recent:**\n${unreadSummary}${unreadEmails.length > 5 ? `\n\n...and ${unreadEmails.length - 5} more` : ''}`
          );
          
          // Still continue to AI for additional analysis
        } catch (error) {
          console.error('Error fetching unread emails:', error);
        }
      }
      
      // Auto-execute search if needed (this takes priority)
      if (needsSearch && graphService.isSignedIn) {
        // Extract search terms
        const searchTerms = extractSearchTerms(text);
        console.log('🔍 Extracted search terms:', searchTerms);
        
        if (searchTerms.length > 0) {
          addFrameworkMessage(`🔍 Searching inbox for: **${searchTerms.join(', ')}**...`);
          
          for (const term of searchTerms) {
            console.log(`🔍 Searching for: "${term}"`);
            const results = await graphService.searchEmails(term, 500);
            console.log(`📬 Found ${results.length} results for "${term}"`);
            searchResults = [...searchResults, ...results];
          }
          
          // Deduplicate
          searchResults = Array.from(new Map(searchResults.map(r => [r.id, r])).values());
          autonomousAgent.setSearchResults(searchResults);
          
          // Show search results immediately
          if (searchResults.length > 0) {
            const uniqueSenders = Array.from(
              new Map(searchResults.map(r => [r.senderEmail, { name: r.sender, email: r.senderEmail }])).values()
            );
            
            const resultsSummary = searchResults.slice(0, 8).map((e: any, i: number) => 
              `${i + 1}. **${e.sender}** <${e.senderEmail}>\n   📧 ${e.subject}\n   📅 ${new Date(e.receivedDateTime).toLocaleDateString()}`
            ).join('\n\n');
            
            addFrameworkMessage(
              `📬 **Found ${searchResults.length} emails** (${uniqueSenders.length} unique senders)\n\n${resultsSummary}`
            );
          } else {
            addFrameworkMessage(
              `📭 **No emails found** matching "${searchTerms.join(', ')}". Try different search terms.`
            );
          }
        }
      }
      
      // Auto-execute inbox summary if needed (only if not searching)
      else if (needsInboxSummary && graphService.isSignedIn) {
        addFrameworkMessage('🔍 Scanning your inbox...');
        inboxSummary = await graphService.getInboxSummary();
        autonomousAgent.setSearchResults(inboxSummary.recentEmails);
      }
      
      // ========================================
      // PHASE 2: LET THE AI AGENT THINK & PLAN
      // ========================================
      
      // Build the autonomous agent prompt with all context
      const agentPrompt = autonomousAgent.buildAgentPrompt(text, emailContext);
      
      // Send to AI model
      const response = await aiService.chat({
        prompt: agentPrompt,
        context: { currentEmail: emailContext || undefined },
      });
      
      // Parse the agent's response
      const parsed = autonomousAgent.parseAgentResponse(response.content);
      
      // ========================================
      // PHASE 3: EXECUTE AI ACTIONS AUTOMATICALLY
      // ========================================
      
      // Check for action commands in the AI response
      const { parseActionsFromResponse, executeAction } = await loadActionExecutor();
      let aiActions = parseActionsFromResponse(response.content);
      
      // ========================================
      // SMART ACTION CONSOLIDATION
      // If AI outputs multiple move_email actions, convert to bulk move_emails_from_sender
      // ========================================
      const moveEmailActions = aiActions.filter(a => a.type === 'move_email' || a.type === 'move');
      if (moveEmailActions.length > 2) {
        console.log(`🔄 Converting ${moveEmailActions.length} move_email actions to bulk operation`);
        
        // Get the target folder from the first move action
        const targetFolder = moveEmailActions[0]?.params?.folderName || 'Archive';
        
        // Try to determine the sender domain from search context
        const searchContext = autonomousAgent.getSearchContext();
        let senderDomain = '';
        
        // Look for common domains in the search results
        if (searchContext && searchContext.length > 0) {
          const domains: Record<string, number> = {};
          searchContext.forEach((email: any) => {
            if (email.senderEmail) {
              const domain = '@' + email.senderEmail.split('@')[1];
              domains[domain] = (domains[domain] || 0) + 1;
            }
          });
          // Find the most common domain
          const sortedDomains = Object.entries(domains).sort((a, b) => b[1] - a[1]);
          if (sortedDomains.length > 0) {
            senderDomain = sortedDomains[0][0];
          }
        }
        
        // Remove the individual move_email actions
        aiActions = aiActions.filter(a => a.type !== 'move_email' && a.type !== 'move');
        
        // Add a single bulk move action
        if (senderDomain) {
          aiActions.push({
            type: 'move_emails_from_sender',
            params: { senderDomain, folderName: targetFolder }
          });
          console.log(`🔄 Converted to: move_emails_from_sender { senderDomain: "${senderDomain}", folderName: "${targetFolder}" }`);
        } else {
          // Fallback: move by email IDs if we can't determine domain
          const emailIds = moveEmailActions.map(a => a.params?.emailId).filter(Boolean);
          for (const emailId of emailIds) {
            aiActions.push({ type: 'move_email', params: { emailId, folderName: targetFolder } });
          }
        }
      }
      
      // Execute safe actions automatically - these are either read-only or low-risk operations
      const safeActions = [
        // Search/Query (read-only)
        'search', 'search_emails', 'advanced_search', 'get_unread', 'unread', 
        'get_email_details', 'details', 'get_emails_from_sender', 'from_sender',
        'get_calendar', 'calendar', 'get_tasks', 'tasks', 'get_contacts', 'contacts', 
        'get_folders', 'folders', 'search_contacts', 'get_attachments', 'get_conversation',
        'get_sent_items', 'get_drafts', 'get_free_busy', 'get_mail_rules', 'get_auto_reply', 'get_categories',
        
        // Email management (non-destructive)
        'mark_read', 'mark_unread', 'flag', 'flag_email', 'unflag', 'unflag_email',
        'mark_all_unread_as_read', 'archive_email', 'archive', 'move_email', 'move',
        'set_categories', 'set_importance',
        
        // Email management (destructive but internal)
        'delete_email', 'delete', 'delete_all_from_sender',
        
        // Task management (internal)
        'create_task', 'task', 'todo', 'complete_task', 'update_task', 'delete_task',
        
        // Draft creation (doesn't send)
        'create_draft', 'draft',
        
        // Folder management (internal)
        'create_folder', 'rename_folder', 'delete_folder',
        
        // Contact management (internal)
        'create_contact', 'update_contact', 'delete_contact',
        
        // Rules (internal)
        'create_mail_rule', 'delete_mail_rule',
        
        // Calendar management (internal unless attendees present)
        'create_event', 'schedule', 'meeting', 'create_recurring_event', 'update_event', 'delete_event',
        
        // Analytics and summaries (read-only)
        'get_email_stats', 'get_top_senders', 'summarize_email', 'summarize_thread', 
        'extract_action_items', 'draft_reply',
        
        // Bulk archiving/moving (non-destructive, can be undone)
        'archive_all_from_sender', 'archive_older_than', 'move_emails_from_sender',
      ];
      
      // Actions that always need user approval before executing (outbound)
      const needsApprovalActions = [
        // Sending messages
        'send_email', 'send', 'reply_email', 'reply', 'reply_all', 'forward_email', 'forward', 'send_draft',

        // Meeting responses (notify others)
        'accept_meeting', 'decline_meeting', 'tentative_meeting',

        // Auto-reply (sends out of office responses)
        'set_auto_reply',
      ];

      const requiresApproval = (action: { type: string; params: Record<string, any> }) => {
        if (needsApprovalActions.includes(action.type)) return true;

        // Calendar actions only require approval if attendees are included (outbound invite/updates)
        const isCalendarAction = [
          'create_event', 'schedule', 'meeting',
          'create_recurring_event', 'update_event', 'delete_event',
        ].includes(action.type);

        if (isCalendarAction) {
          const attendees = action.params?.attendees;
          return Array.isArray(attendees) ? attendees.length > 0 : !!attendees;
        }

        return false;
      };
      
      // Collect action results for context feedback
      const actionResults: Array<{ action: string; success: boolean; message: string; data?: any }> = [];
      
      for (const action of aiActions) {
        if (safeActions.includes(action.type)) {
          // Execute read-only actions immediately
          console.log('🤖 Auto-executing safe action:', action.type, action.params);
          const result = await executeAction(action);
          
          // Store result for AI context feedback
          actionResults.push({
            action: action.type,
            success: result.success,
            message: result.message,
            data: result.data,
          });
          
          // Always show feedback for executed actions
          if (result.success) {
            let resultSummary = '';
            
            if (result.data && Array.isArray(result.data)) {
              if (result.data.length > 0) {
                // Format based on data type
                if (result.data[0].subject !== undefined) {
                  // Email results - store in agent context for future queries
                  autonomousAgent.setSearchResults(result.data);
                  resultSummary = result.data.slice(0, 10).map((e: any, i: number) => 
                    `${i + 1}. **${e.sender || e.senderEmail}** <${e.senderEmail}>\n   📧 ${e.subject}\n   📅 ${new Date(e.receivedDateTime).toLocaleDateString()}`
                  ).join('\n\n');
                } else if (result.data[0].title !== undefined) {
                  // Task results
                  resultSummary = result.data.slice(0, 10).map((t: any, i: number) => 
                    `${i + 1}. ${t.title}${t.dueDate ? ` (Due: ${new Date(t.dueDate).toLocaleDateString()})` : ''}`
                  ).join('\n');
                } else if (result.data[0].displayName !== undefined && result.data[0].unreadItemCount !== undefined) {
                  // Folder results
                  resultSummary = result.data.map((f: any) => 
                    `📁 ${f.displayName}${f.unreadItemCount > 0 ? ` (${f.unreadItemCount} unread)` : ''}`
                  ).join('\n');
                } else if (result.data[0].displayName !== undefined) {
                  // Contact results
                  autonomousAgent.setContactsContext(result.data);
                  resultSummary = result.data.slice(0, 10).map((c: any) => 
                    `👤 ${c.displayName}${c.emailAddresses?.length ? ` - ${c.emailAddresses[0]}` : ''}`
                  ).join('\n');
                }
              }
            }
            
            // Add action result to agent's conversation context for continuity
            autonomousAgent.addTurn({
              role: 'system',
              content: `[ACTION COMPLETED: ${result.message}]${resultSummary ? '\n' + resultSummary : ''}`,
              timestamp: new Date(),
            });
            
            // Always show a message for successful actions
            addFrameworkMessage(
              resultSummary
                ? `📊 **${result.message}**\n\n${resultSummary}`
                : `✅ **${result.message}**`
            );
          } else {
            // Show failed actions
            console.warn('⚠️ Action failed:', action.type, result.message);
            addFrameworkMessage(`❌ **Action failed:** ${result.message}`);
          }
        } else if (requiresApproval(action)) {
          // Queue for approval
          const approval = {
            id: uuidv4(),
            type: action.type,
            description: getActionDescription(action),
            details: action.params,
          };
          setPendingApprovals(prev => [...prev, approval]);
          
          // Show that approval is pending
          addFrameworkMessage(
            `⏳ **Awaiting approval:** ${getActionDescription(action)}\n\nPlease approve or reject this action above.`
          );
        } else {
          // Unknown action - log it for debugging
          console.warn('⚠️ Unknown action type:', action.type, '- not in safe or approval lists');
        }
      }
      
      // Log action summary for debugging
      if (actionResults.length > 0) {
        console.log('📋 Action results summary:', actionResults.map(r => `${r.action}: ${r.success ? '✅' : '❌'}`).join(', '));
      }
      
      // ========================================
      // PHASE 4: DISPLAY RESULTS & HANDLE APPROVALS
      // ========================================
      
      // Store any pending approvals from parsing
      if (parsed.pendingApprovals.length > 0) {
        const newApprovals = parsed.pendingApprovals.map(a => ({
          id: uuidv4(),
          type: a.type,
          description: a.description,
          details: a.details,
        }));
        setPendingApprovals(prev => [...prev, ...newApprovals]);
        autonomousAgent.setPendingApprovals(parsed.pendingApprovals);
      }
      
      // Build the display message
      let displayContent = '';
      
      // Show thinking if present (collapsed/subtle)
      if (parsed.thinking) {
        displayContent += `💭 *${parsed.thinking}*\n\n`;
      }
      
      // Show actions taken
      if (parsed.actions) {
        displayContent += `${parsed.actions}\n\n`;
      }
      
      // Show main result
      displayContent += parsed.result;
      
      // Add the response
      addMessage({
        id: uuidv4(),
        role: 'assistant',
        content: displayContent,
        timestamp: new Date(),
      });
      
      // Add to agent conversation
      autonomousAgent.addTurn({
        role: 'agent',
        content: displayContent,
        timestamp: new Date(),
        searchResults: searchResults.length > 0 ? searchResults : undefined,
      });

    } catch (error) {
      console.error('Agent error:', error);
      addFrameworkMessage('Sorry, I encountered an error. Please try again.');
    } finally {
      setIsProcessing(false);
    }
  };
  
  // Helper to get human-readable description of an action
  const getActionDescription = (action: { type: string; params: Record<string, any> }): string => {
    const { type, params } = action;
    switch (type) {
      case 'send_email':
      case 'send':
        return `Send email to ${params.to} - Subject: "${params.subject}"`;
      case 'reply_email':
      case 'reply':
        return `Reply to email`;
      case 'forward_email':
      case 'forward':
        return `Forward email to ${params.to}`;
      case 'delete_email':
      case 'delete':
        return `Delete email`;
      case 'archive_email':
      case 'archive':
        return `Archive email`;
      case 'create_event':
      case 'schedule':
      case 'meeting':
        return `Create calendar event: "${params.subject}" on ${params.start}`;
      case 'create_recurring_event':
        return `Create recurring event: "${params.subject}"`;
      case 'update_event':
        return `Update calendar event`;
      case 'delete_event':
        return `Delete calendar event`;
      case 'create_task':
      case 'task':
      case 'todo':
        return `Create task: "${params.title}"`;
      case 'delete_task':
        return `Delete task`;
      case 'delete_contact':
        return `Delete contact`;
      case 'delete_folder':
        return `Delete folder: "${params.folderId || params.folderName}"`;
      case 'delete_mail_rule':
        return `Delete mail rule`;
      case 'send_draft':
        return `Send draft email`;
      case 'reply_all':
        return `Reply all to email`;
      default:
        return `Execute: ${type}`;
    }
  };
  
  // Helper to extract search terms from user message
  const extractSearchTerms = (message: string): string[] => {
    const terms: string[] = [];
    const lowerMsg = message.toLowerCase();
    
    // Direct keyword extraction - high priority terms
    const directKeywords = [
      'activated carbon',
      'patent office',
      'uspto',
    ];
    
    for (const keyword of directKeywords) {
      if (lowerMsg.includes(keyword) && !terms.includes(keyword)) {
        terms.push(keyword);
      }
    }
    
    // Common patterns
    const patterns = [
      /(?:from|by)\s+(?:the\s+)?([a-z0-9\s\-\.@]+?)(?:\s+(?:in|about|and|asking|then)|$)/gi,
      /(?:about|regarding)\s+([a-z0-9\s\-]+?)(?:\s+(?:in|from|and|then)|$)/gi,
      /(?:any\s+)?([a-z0-9\s\-]+?)\s+(?:distributors?|suppliers?|vendors?|companies?)/gi,
      /(?:for\s+(?:any\s+)?)?([a-z0-9\s\-]+?)\s+(?:distribut|suppli|vendor)/gi,
    ];
    
    for (const pattern of patterns) {
      let match;
      pattern.lastIndex = 0; // Reset regex
      while ((match = pattern.exec(message)) !== null) {
        const term = match[1]?.trim();
        if (term && term.length > 2 && !terms.some(t => t.toLowerCase() === term.toLowerCase())) {
          // Don't add generic words
          const skipWords = ['any', 'the', 'all', 'my', 'inbox', 'email', 'emails', 'review', 'then', 'lets', 'make'];
          if (!skipWords.includes(term.toLowerCase())) {
            terms.push(term);
          }
        }
      }
    }
    
    // If no patterns matched and we have specific industry terms, use those
    if (terms.length === 0) {
      const industryTerms = ['carbon', 'chemical', 'steel', 'plastic', 'metal', 'supply'];
      for (const term of industryTerms) {
        if (lowerMsg.includes(term)) {
          terms.push(term);
        }
      }
    }
    
    console.log('📝 Extracted terms from message:', terms);
    return terms.slice(0, 3); // Limit to 3 search terms
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
      addFrameworkMessage(`⏳ Generating ${type.toUpperCase()} document...`);
      
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
      addFrameworkMessage(
        `✅ **${type.toUpperCase()} exported!** ${templateInfo?.icon || '📄'}\n\nUsed template: **${templateInfo?.name || 'Custom'}**\n\n${type === 'pdf' || type === 'powerpoint' ? '💡 Tip: Use Ctrl+P to save as PDF or F11 for fullscreen presentation.' : 'Check your downloads folder.'}`
      );
    } catch (error: any) {
      console.error('Export error:', error);
      addFrameworkMessage(`❌ Export failed: ${error.message}`);
    }
  };

  // Handle branding save
  const handleBrandingSave = (branding: UserBranding) => {
    setUserBranding(branding);
    setShowBranding(false);
    addFrameworkMessage(
      `✅ **Branding saved!** 🎨\n\nYour custom branding will now be used in all document exports:\n• Company: **${branding.companyName || 'Not set'}**\n• Colors: Primary ${branding.primaryColor}\n• Logo: ${branding.logoUrl ? '✓ Uploaded' : 'Not set'}\n\nTry exporting a document to see your branding in action!`
    );
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
              I can autonomously help you with:
            </p>
            <ul style={{ fontSize: '12px', textAlign: 'left', margin: '8px 0', paddingLeft: '20px', color: '#555' }}>
              <li>📥 Search & analyze your entire inbox</li>
              <li>📅 Manage your calendar & meetings</li>
              <li>✉️ Draft & send emails (with your approval)</li>
              <li>🔍 Find contacts and compose messages</li>
              <li>📋 Extract action items & create tasks</li>
            </ul>
            <p className="empty-state-text" style={{ fontSize: '12px', marginTop: '8px' }}>
              Just tell me what you need - I'll handle the rest!
            </p>
          </div>
        ) : (
          <>
            {messages.map((message) => (
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
            ))}
            
            {/* Pending Approvals */}
            {pendingApprovals.length > 0 && (
              <div className="pending-approvals">
                <div className="approval-header">
                  ⚠️ <strong>Actions awaiting your approval:</strong>
                </div>
                {pendingApprovals.map((approval) => (
                  <div key={approval.id} className="approval-card">
                    <div className="approval-type">
                      {approval.type === 'send_email' && '📤 Send Email'}
                      {approval.type === 'delete_email' && '🗑️ Delete Email'}
                      {approval.type === 'create_calendar_event' && '📅 Create Event'}
                      {approval.type === 'reply_to_email' && '↩️ Reply'}
                      {approval.type === 'forward_email' && '↪️ Forward'}
                    </div>
                    <div className="approval-details">
                      {approval.type === 'send_email' && (
                        <>
                          <div><strong>To:</strong> {approval.details.to}</div>
                          <div><strong>Subject:</strong> {approval.details.subject}</div>
                          <div className="approval-body"><strong>Body:</strong><br/>{approval.details.body}</div>
                        </>
                      )}
                      {approval.type === 'create_calendar_event' && (
                        <>
                          <div><strong>Event:</strong> {approval.details.title}</div>
                          <div><strong>Time:</strong> {approval.details.time}</div>
                        </>
                      )}
                      {approval.type === 'delete_email' && (
                        <div><strong>Email:</strong> {approval.details.subject}</div>
                      )}
                    </div>
                    <div className="approval-buttons">
                      <button 
                        className="approve-btn"
                        onClick={() => handleApproval(approval.id, true)}
                        disabled={isProcessing}
                      >
                        ✅ Approve
                      </button>
                      <button 
                        className="reject-btn"
                        onClick={() => handleApproval(approval.id, false)}
                        disabled={isProcessing}
                      >
                        ❌ Cancel
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </>
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
