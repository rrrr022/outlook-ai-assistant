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
   * Build the full context for the AI model
   * This gives the model everything it needs to make decisions
   */
  buildAgentPrompt(userMessage: string, currentEmail?: EmailSummary | null): string {
    const conversationContext = this.buildConversationContext();
    const searchContext = this.buildSearchContext();
    const contactsContext = this.buildContactsContext();
    
    return `You are a helpful Outlook email assistant. You help the user manage their inbox, calendar, and compose emails.

## WHAT YOU CAN DO

### Reading & Searching (do these directly)
- Search emails by keyword or sender
- Get email details and summaries
- View calendar events
- Analyze and summarize email threads

### Composing & Sending (ask user first)
- Draft new emails
- Draft replies
- Suggest calendar events

## GUIDELINES

1. When asked to find emails, use the search results provided below.
2. When asked to summarize, provide specific details from the emails.
3. When drafting emails, always show the draft and ask if the user wants to send it.
4. Use the context and data provided - don't say you lack access to information that's given below.
5. Be concise and helpful.

## RESPONSE FORMAT

**Summary**: Brief overview of what you found or did

**Details**: Specific information, email summaries, or drafted content

**Next Steps**: If you need user approval to send/create something, clearly state what action you're proposing

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

Now process this request. Remember:
- Search/read actions: Just do them and report results
- Send/delete/calendar actions: Draft them and ask for permission
- Use the contacts and search results I've provided
- Be helpful and proactive`;
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
