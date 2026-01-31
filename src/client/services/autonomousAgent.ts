import { EmailSummary } from '../../shared/types';

/**
 * Autonomous Agent Service
 * 
 * This agent has full control and can:
 * - Plan actions based on user requests
 * - Execute safe actions automatically (search, read, summarize)
 * - Request permission for sensitive actions (send, delete, calendar)
 * - Maintain full conversation context
 * 
 * Think of it like how a human assistant works:
 * 1. Listen to what you need
 * 2. Research and gather information
 * 3. Draft a plan
 * 4. Ask "Should I go ahead with this?" for important actions
 * 5. Execute when approved
 */

// Actions the agent can take
export type AgentAction = 
  | { type: 'search_inbox'; query: string }
  | { type: 'get_email_details'; emailId: string }
  | { type: 'get_inbox_summary' }
  | { type: 'get_calendar'; days: number }
  | { type: 'summarize_emails'; emails: EmailSummary[] }
  | { type: 'draft_email'; to: string; subject: string; body: string }
  | { type: 'send_email'; to: string; subject: string; body: string }
  | { type: 'delete_email'; emailId: string }
  | { type: 'create_calendar_event'; title: string; start: Date; end: Date; attendees?: string[] }
  | { type: 'reply_to_email'; emailId: string; body: string }
  | { type: 'forward_email'; emailId: string; to: string; note?: string }
  | { type: 'mark_as_read'; emailId: string }
  | { type: 'think'; thought: string }
  | { type: 'respond'; message: string };

// Actions that require user permission
const PERMISSION_REQUIRED_ACTIONS = [
  'send_email',
  'delete_email', 
  'create_calendar_event',
  'reply_to_email',
  'forward_email',
];

// Agent's execution plan
export interface AgentPlan {
  thinking: string;           // Agent's reasoning
  actions: PlannedAction[];   // What it plans to do
  requiresPermission: boolean; // Does any action need approval?
  pendingActions: PlannedAction[]; // Actions waiting for approval
}

export interface PlannedAction {
  action: AgentAction;
  description: string;        // Human-readable description
  status: 'pending' | 'executing' | 'completed' | 'failed' | 'awaiting_permission';
  result?: any;
  error?: string;
}

export interface ConversationTurn {
  role: 'user' | 'agent' | 'system';
  content: string;
  timestamp: Date;
  plan?: AgentPlan;
  searchResults?: EmailSummary[];
  executedActions?: PlannedAction[];
}

export interface AgentState {
  conversation: ConversationTurn[];
  currentPlan: AgentPlan | null;
  lastSearchResults: EmailSummary[];
  knownContacts: Map<string, string>;
  pendingApprovals: PlannedAction[];
  isExecuting: boolean;
}

class AutonomousAgent {
  private state: AgentState = {
    conversation: [],
    currentPlan: null,
    lastSearchResults: [],
    knownContacts: new Map(),
    pendingApprovals: [],
    isExecuting: false,
  };

  /**
   * Check if an action requires permission
   */
  requiresPermission(action: AgentAction): boolean {
    return PERMISSION_REQUIRED_ACTIONS.includes(action.type);
  }

  /**
   * Get current state
   */
  getState(): AgentState {
    return this.state;
  }

  /**
   * Add a conversation turn
   */
  addTurn(turn: ConversationTurn) {
    this.state.conversation.push(turn);
    // Keep last 50 turns
    if (this.state.conversation.length > 50) {
      this.state.conversation = this.state.conversation.slice(-50);
    }
  }

  /**
   * Store search results for context
   */
  setSearchResults(results: EmailSummary[]) {
    this.state.lastSearchResults = results;
    // Extract contacts
    for (const email of results) {
      if (email.sender && email.senderEmail) {
        this.state.knownContacts.set(email.sender.toLowerCase(), email.senderEmail);
      }
    }
  }

  /**
   * Get the last search results (for smart action consolidation)
   */
  getSearchContext(): EmailSummary[] {
    return this.state.lastSearchResults;
  }

  /**
   * Store contacts for context
   */
  setContactsContext(contacts: any[]) {
    for (const contact of contacts) {
      if (contact.displayName && contact.emailAddresses?.length > 0) {
        this.state.knownContacts.set(contact.displayName.toLowerCase(), contact.emailAddresses[0]);
      }
    }
  }

  /**
   * Build the full context for the AI model
   * This gives the model everything it needs to make decisions
   */
  buildAgentPrompt(userMessage: string, currentEmail?: EmailSummary | null): string {
    const conversationContext = this.buildConversationContext();
    const searchContext = this.buildSearchContext();
    const contactsContext = this.buildContactsContext();
    
    return `You are a helpful Outlook email assistant with FULL control over the user's Outlook. You can manage emails, calendar, tasks, contacts, folders, and settings.

## CAPABILITIES

### Reading (automatic)
- Search/view emails, calendar, tasks, contacts, folders
- Get attachments, conversation threads, sent items, drafts
- Check free/busy schedules, mail rules, auto-reply settings

### Writing (some require confirmation)
- Send/reply/forward/delete emails
- Create/update/delete calendar events
- Accept/decline/tentative meeting responses
- Create/complete/delete tasks
- Create/update/delete contacts
- Create/rename/delete folders
- Set categories, importance, flags
- Configure mail rules and auto-reply (OOF)

## HOW TO EXECUTE ACTIONS

Include actions in your response:
[ACTION:action_name]{"param": "value"}[/ACTION]

### Email Actions:
- send_email: {"to": "email", "subject": "...", "body": "...", "cc": "...", "bcc": "..."}
- reply_email: {"emailId": "id", "body": "..."}
- reply_all: {"emailId": "id", "body": "..."}
- forward_email: {"emailId": "id", "to": "email", "comment": "..."}
- delete_email: {"emailId": "id"}
- archive_email: {"emailId": "id"}
- flag_email / unflag_email: {"emailId": "id"}
- mark_read: {"emailId": "id"} - mark single email as read
- mark_unread: {"emailId": "id"} - mark single email as unread
- mark_all_unread_as_read: {"count": 500} - **BULK OPERATION** marks ALL unread emails as read (use this when user asks to mark all as read!)
- move_email: {"emailId": "id", "folderName": "Folder"}
- create_draft: {"to": [...], "subject": "...", "body": "..."}
- send_draft: {"draftId": "id"}
- set_categories: {"emailId": "id", "categories": ["Red", "Blue"]}
- set_importance: {"emailId": "id", "importance": "high|normal|low"}
- get_attachments: {"emailId": "id"}
- get_conversation: {"conversationId": "id"}
- get_sent_items / get_drafts: {"count": 20}

### Calendar Actions:
- create_event: {"subject": "...", "start": "ISO date", "end": "ISO date", "attendees": [...], "location": "..."}
- create_recurring_event: {"subject": "...", "start": "...", "end": "...", "recurrence": {"pattern": "daily|weekly|monthly", "interval": 1, "daysOfWeek": ["monday"], "endDate": "..."}}
- update_event: {"eventId": "id", "subject": "...", "start": "...", "end": "...", "location": "..."}
- delete_event: {"eventId": "id"}
- accept_meeting: {"eventId": "id", "comment": "..."}
- decline_meeting: {"eventId": "id", "comment": "..."}
- tentative_meeting: {"eventId": "id", "comment": "..."}
- get_calendar: {"days": 7}
- get_free_busy: {"emails": [...], "start": "...", "end": "..."}

### Task Actions:
- create_task: {"title": "...", "dueDate": "...", "body": "..."}
- complete_task: {"taskId": "id"}
- delete_task: {"taskId": "id"}
- update_task: {"taskId": "id", "title": "...", "dueDate": "...", "importance": "high"}
- get_tasks: {"count": 20}

### Contact Actions:
- create_contact: {"givenName": "...", "surname": "...", "emailAddresses": [...], "companyName": "..."}
- update_contact: {"contactId": "id", "jobTitle": "..."}
- delete_contact: {"contactId": "id"}
- get_contacts / search_contacts: {"query": "..."}

### Folder Actions:
- create_folder: {"displayName": "..."}
- rename_folder: {"folderId": "id", "newName": "..."}
- delete_folder: {"folderId": "id"}
- get_folders: {}

### Rules & Settings:
- get_mail_rules / create_mail_rule / delete_mail_rule
- get_auto_reply: {} (check OOF status)
- set_auto_reply: {"status": "disabled|alwaysEnabled|scheduled", "internalMessage": "...", "externalMessage": "...", "scheduledStartDateTime": "...", "scheduledEndDateTime": "..."}
- get_categories: {}

### Search Actions (use these to find emails):
- search_emails: {"query": "search text", "count": 250} - Full text search across all folders
- advanced_search: {"searchText": "...", "senderEmail": "...", "senderDomain": "@company.com", "subjectContains": "...", "hasAttachments": true/false, "isRead": true/false, "importance": "high|normal|low", "startDate": "2024-01-01", "endDate": "2024-12-31", "count": 500} - Powerful filtered search
- get_unread: {"count": 250} - Get all unread emails
- get_emails_from_sender: {"senderEmail": "...", "count": 100} - Get emails from specific sender
- get_email_details: {"emailId": "id"} - Get full email content with body

### Bulk Operations (use for mass actions):
- mark_all_unread_as_read: {"count": 500} - Mark all unread emails as read at once
- archive_all_from_sender: {"senderEmail": "...", "count": 100} - Archive all emails from a specific sender
- archive_older_than: {"days": 30, "count": 200} - Archive emails older than X days
- delete_all_from_sender: {"senderEmail": "...", "count": 100} - Delete all emails from sender (requires approval)
- move_emails_from_sender: {"senderEmail": "...", "folderName": "Folder Name", "count": 100} - Move all emails from sender to a folder
- move_emails_from_sender: {"senderDomain": "@domain.com", "folderName": "Folder Name"} - Move all emails from a domain to a folder

### Analytics & Insights:
- get_email_stats: {"days": 30} - Get email statistics (totals, averages, top senders)
- get_top_senders: {"count": 10, "days": 30} - Get top N email senders

### AI-Powered Features:
- summarize_email: {"emailId": "id"} - Get a summary of a specific email
- summarize_thread: {"emailId": "id"} OR {"conversationId": "id"} - Summarize an email conversation thread
- extract_action_items: {"emailId": "id"} - Extract action items and to-dos from an email
- draft_reply: {"emailId": "id", "tone": "professional|casual|formal"} - Generate a reply template

## ⚠️ CRITICAL RULES - YOU MUST FOLLOW THESE EXACTLY ⚠️

1. **ABSOLUTELY NO HEADERS** - NEVER use "##", "###", "Step 1", "Creating folder", etc.
2. **NEVER NEVER NEVER LIST EMAIL IDs** - Even if search results show email IDs, DO NOT USE THEM!
3. **ALWAYS USE BULK ACTIONS** for multiple emails:
   - USPTO emails → move_emails_from_sender with senderDomain: "@uspto.gov"
   - Newsletter emails → move_emails_from_sender with senderEmail
   - Mark all read → mark_all_unread_as_read
   - Archive/Delete from sender → archive_all_from_sender / delete_all_from_sender
4. **RESPONSE FORMAT**: One concise sentence, then action blocks. No status words like "Done!".
5. **IGNORE the email IDs in search results** - they are for display only, not for you to use!
3. **BE BRIEF** - One sentence max, then actions. No explanations.
4. **EXECUTE IMMEDIATELY** - Don't describe, just DO IT.
5. For SEND/DELETE, user confirmation is automatic.

## EXAMPLES:

User: "Create a Patent folder and move USPTO emails there"
CORRECT RESPONSE:
Creating the folder and moving those emails now.
[ACTION:create_folder]{"displayName": "Patent"}[/ACTION]
[ACTION:move_emails_from_sender]{"senderDomain": "@uspto.gov", "folderName": "Patent"}[/ACTION]

WRONG RESPONSE (DO NOT DO THIS):
## Creating Patent Folder
[ACTION:create_folder]...
## Moving Emails
The following email IDs will be moved...
[ACTION:move_email]{"emailId": "AAMk..."}

User: "Mark all unread as read"
CORRECT:
Marking all unread emails as read.
[ACTION:mark_all_unread_as_read]{}[/ACTION]

---

## CURRENT CONTEXT

${conversationContext}

${searchContext}

${contactsContext}

${currentEmail ? `
### Currently Selected Email
- From: ${currentEmail.sender} <${currentEmail.senderEmail}>
- Subject: ${currentEmail.subject}
- Preview: ${currentEmail.preview}
` : ''}

---

## USER REQUEST

${userMessage}

---

Remember: One sentence then actions only. No headers, no email IDs, no explanations.`;
  }

  /**
   * Build conversation history context
   */
  private buildConversationContext(): string {
    if (this.state.conversation.length === 0) {
      return '### Conversation History\nThis is the start of the conversation.';
    }

    const recent = this.state.conversation.slice(-10);
    let context = '### Conversation History\n';
    
    for (const turn of recent) {
      const role = turn.role === 'user' ? 'User' : 'Agent';
      context += `\n**${role}**: ${turn.content.substring(0, 500)}${turn.content.length > 500 ? '...' : ''}`;
      
      if (turn.searchResults && turn.searchResults.length > 0) {
        context += `\n  [Found ${turn.searchResults.length} emails]`;
      }
      if (turn.executedActions && turn.executedActions.length > 0) {
        context += `\n  [Executed: ${turn.executedActions.map(a => a.action.type).join(', ')}]`;
      }
    }
    
    return context;
  }

  /**
   * Build search results context
   */
  private buildSearchContext(): string {
    if (this.state.lastSearchResults.length === 0) {
      return '### Previous Search Results\nNo recent searches.';
    }

    let context = `### Previous Search Results (${this.state.lastSearchResults.length} emails)\n`;
    
    for (const email of this.state.lastSearchResults.slice(0, 15)) {
      context += `\n- **${email.sender}** <${email.senderEmail}>`;
      context += `\n  Subject: ${email.subject}`;
      context += `\n  Date: ${new Date(email.receivedDateTime).toLocaleDateString()}`;
      context += `\n  Preview: ${email.preview?.substring(0, 100) || 'No preview'}...`;
      context += `\n  ID: ${email.id}\n`;
    }
    
    return context;
  }

  /**
   * Build known contacts context
   */
  private buildContactsContext(): string {
    if (this.state.knownContacts.size === 0) {
      return '### Known Contacts\nNo contacts discovered yet.';
    }

    let context = '### Known Contacts (from inbox)\n';
    
    for (const [name, email] of Array.from(this.state.knownContacts.entries()).slice(0, 20)) {
      context += `- ${name}: ${email}\n`;
    }
    
    return context;
  }

  /**
   * Parse the AI response to extract planned actions
   */
  parseAgentResponse(response: string): {
    thinking: string;
    actions: string;
    result: string;
    pendingApprovals: Array<{
      type: 'send_email' | 'delete_email' | 'create_calendar_event' | 'reply_to_email' | 'forward_email';
      details: any;
      description: string;
    }>;
  } {
    // Extract sections from response
    const thinkingMatch = response.match(/\*\*THINKING\*\*:?\s*([\s\S]*?)(?=\*\*ACTIONS\*\*|\*\*RESULT\*\*|$)/i);
    const actionsMatch = response.match(/\*\*ACTIONS\*\*:?\s*([\s\S]*?)(?=\*\*RESULT\*\*|\*\*NEEDS APPROVAL\*\*|$)/i);
    const resultMatch = response.match(/\*\*RESULT\*\*:?\s*([\s\S]*?)(?=\*\*NEEDS APPROVAL\*\*|$)/i);
    const approvalMatch = response.match(/\*\*NEEDS APPROVAL\*\*:?\s*([\s\S]*?)$/i);

    const pendingApprovals: Array<{
      type: 'send_email' | 'delete_email' | 'create_calendar_event' | 'reply_to_email' | 'forward_email';
      details: any;
      description: string;
    }> = [];

    // Parse email drafts from approval section
    if (approvalMatch) {
      const approvalText = approvalMatch[1];
      
      // Look for email patterns
      const emailPattern = /Send email to:?\s*([^\n]+)\s*Subject:?\s*"?([^"\n]+)"?\s*Body:?\s*"?([\s\S]*?)"?(?=\n\n|✅|$)/gi;
      let match;
      while ((match = emailPattern.exec(approvalText)) !== null) {
        pendingApprovals.push({
          type: 'send_email',
          details: {
            to: match[1].trim(),
            subject: match[2].trim(),
            body: match[3].trim(),
          },
          description: `Send email to ${match[1].trim()} about "${match[2].trim()}"`,
        });
      }

      // Look for calendar event patterns
      const calendarPattern = /Create (?:calendar )?event:?\s*"?([^"\n]+)"?\s*(?:on|at|from):?\s*([^\n]+)/gi;
      while ((match = calendarPattern.exec(approvalText)) !== null) {
        pendingApprovals.push({
          type: 'create_calendar_event',
          details: {
            title: match[1].trim(),
            time: match[2].trim(),
          },
          description: `Create calendar event: ${match[1].trim()}`,
        });
      }

      // Look for delete patterns
      const deletePattern = /Delete email:?\s*"?([^"\n]+)"?/gi;
      while ((match = deletePattern.exec(approvalText)) !== null) {
        pendingApprovals.push({
          type: 'delete_email',
          details: {
            subject: match[1].trim(),
          },
          description: `Delete email: ${match[1].trim()}`,
        });
      }
    }

    return {
      thinking: thinkingMatch?.[1]?.trim() || '',
      actions: actionsMatch?.[1]?.trim() || '',
      result: resultMatch?.[1]?.trim() || response,
      pendingApprovals,
    };
  }

  /**
   * Set pending approvals
   */
  setPendingApprovals(approvals: any[]) {
    this.state.pendingApprovals = approvals.map(a => ({
      action: { type: a.type, ...a.details } as AgentAction,
      description: a.description,
      status: 'awaiting_permission' as const,
    }));
  }

  /**
   * Get pending approvals
   */
  getPendingApprovals(): PlannedAction[] {
    return this.state.pendingApprovals;
  }

  /**
   * Clear pending approvals
   */
  clearPendingApprovals() {
    this.state.pendingApprovals = [];
  }

  /**
   * Approve a specific action
   */
  approveAction(index: number): PlannedAction | null {
    if (index >= 0 && index < this.state.pendingApprovals.length) {
      const action = this.state.pendingApprovals[index];
      action.status = 'pending';
      return action;
    }
    return null;
  }

  /**
   * Reset agent state
   */
  reset() {
    this.state = {
      conversation: [],
      currentPlan: null,
      lastSearchResults: [],
      knownContacts: new Map(),
      pendingApprovals: [],
      isExecuting: false,
    };
  }
}

export const autonomousAgent = new AutonomousAgent();
