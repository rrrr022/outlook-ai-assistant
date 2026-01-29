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
    
    return `You are an autonomous AI agent with full access to the user's Outlook inbox, calendar, and email capabilities.

## YOUR CAPABILITIES (Tools you can use)

### READ/SEARCH (Execute automatically - no permission needed)
- search_inbox(query): Search for emails matching a query
- get_email_details(emailId): Get full content of an email  
- get_inbox_summary(): Get overview of inbox
- get_calendar(days): Get upcoming calendar events
- summarize_emails(emails): Analyze and summarize emails

### WRITE/SEND (REQUIRES USER PERMISSION - always ask first)
- send_email(to, subject, body): Send a new email
- reply_to_email(emailId, body): Reply to an email
- forward_email(emailId, to, note): Forward an email
- delete_email(emailId): Delete an email
- create_calendar_event(title, start, end, attendees): Create calendar event

## CRITICAL RULES

1. **BE AUTONOMOUS**: Don't ask the user to do things you can do yourself. If they ask to find something, SEARCH FOR IT.

2. **EXECUTE READ ACTIONS IMMEDIATELY**: When you need to search, summarize, or gather info - just do it. Report what you found.

3. **ALWAYS ASK PERMISSION FOR WRITE ACTIONS**: Before sending emails, deleting, or creating events, show the user exactly what you plan to do and ask "Should I proceed?"

4. **USE CONTEXT**: I'm providing you with conversation history and search results. USE THEM. Don't say "I don't have access" when I've given you the data.

5. **BE SPECIFIC**: When drafting emails, use actual email addresses from search results. When summarizing, give specific details.

## RESPONSE FORMAT

Structure your response as:

**THINKING**: (Brief explanation of what you understand and plan to do)

**ACTIONS**: (What you're doing or have done)
- [DONE] Searched inbox for "X" - found Y results
- [DONE] Analyzed emails from Z
- [PENDING APPROVAL] Ready to send email to abc@xyz.com

**RESULT**: (Your actual response to the user - summaries, drafts, findings, etc.)

**NEEDS APPROVAL**: (If any - list exactly what you want to do)
- Send email to: john@example.com
  Subject: "RE: Project Update"
  Body: "..."
  
  ✅ Approve | ❌ Cancel

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
