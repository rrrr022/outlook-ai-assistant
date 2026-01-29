import { EmailSummary } from '../../shared/types';

/**
 * Intelligent Agent Service
 * 
 * This service acts like a smart assistant that:
 * 1. Understands user intents (search, compose, summarize, etc.)
 * 2. Maintains conversation context and memory
 * 3. Actually performs actions (searches inbox, gets emails, etc.)
 * 4. Provides rich context to the AI for intelligent responses
 */

interface ConversationMessage {
  role: 'user' | 'assistant' | 'system';
  content: string;
  timestamp: Date;
  metadata?: {
    action?: string;
    searchResults?: EmailSummary[];
    emailContext?: EmailSummary;
    searchQuery?: string;
  };
}

interface Intent {
  type: 'search' | 'compose' | 'reply' | 'summarize' | 'list' | 'action_items' | 'general' | 'follow_up';
  confidence: number;
  entities: {
    searchTerms?: string[];
    senderName?: string;
    senderEmail?: string;
    subject?: string;
    dateRange?: { start?: Date; end?: Date };
    topic?: string;
  };
}

interface AgentContext {
  conversationHistory: ConversationMessage[];
  lastSearchResults: EmailSummary[];
  lastSearchQuery: string | null;
  currentEmailContext: EmailSummary | null;
  mentionedContacts: Map<string, string>; // name -> email
  mentionedTopics: string[];
  pendingAction: string | null;
}

class AgentService {
  private context: AgentContext = {
    conversationHistory: [],
    lastSearchResults: [],
    lastSearchQuery: null,
    currentEmailContext: null,
    mentionedContacts: new Map(),
    mentionedTopics: [],
    pendingAction: null,
  };

  private maxHistoryLength = 20;

  /**
   * Analyze user message to detect intent
   */
  detectIntent(message: string): Intent {
    const lowerMessage = message.toLowerCase();
    
    // Search/find intent patterns
    const searchPatterns = [
      /(?:search|find|look for|show me|get|check for|any|do i have)\s+(?:emails?|messages?)?.*(?:from|about|regarding|related to|containing)/i,
      /(?:search|find|look for|show|get|check|any)\s+(?:emails?|messages?)?\s*(?:from|about|with|containing)/i,
      /(?:emails?|messages?)\s+(?:from|about|regarding)/i,
      /(?:from|about)\s+(?:the\s+)?(.+?)(?:\s+in\s+my\s+inbox)?$/i,
      /(?:patent office|activated carbon|distributor|supplier|vendor|client)/i,
      /do i have any.*(?:from|about|regarding)/i,
      /any\s+(?:emails?|messages?)?\s*(?:from|about)/i,
    ];
    
    // Compose/draft intent patterns
    const composePatterns = [
      /(?:compose|draft|write|send|create)\s+(?:an?\s+)?(?:email|message|response)/i,
      /(?:reach out|contact|email)\s+(?:to\s+)?/i,
      /(?:ask|request)\s+(?:for|about)\s+/i,
    ];
    
    // Reply intent
    const replyPatterns = [
      /(?:reply|respond|answer|write back)/i,
    ];
    
    // Summarize intent
    const summarizePatterns = [
      /(?:summarize|summary|tldr|key points|what.*about|what.*say)/i,
    ];
    
    // Action items intent
    const actionPatterns = [
      /(?:action items?|tasks?|todo|to-do|follow up|need to do)/i,
    ];
    
    // List/review intent
    const listPatterns = [
      /(?:list|show|review|display)\s+(?:all\s+)?(?:my\s+)?(?:emails?|inbox|messages?)/i,
      /(?:inbox|unread)\s+(?:summary|overview)/i,
    ];

    // Check for follow-up context (user said "please", "yes", "do it", etc. after a question)
    const followUpPatterns = [
      /^(?:please|yes|yeah|yep|sure|ok|okay|do it|go ahead|proceed)\.?$/i,
      /^(?:please|yes|can you)?\s*(?:do that|search|find|look)\s*(?:for me|please)?\.?$/i,
      /^can you do that/i,
    ];

    // Extract entities
    const entities: Intent['entities'] = {};
    
    // Extract search terms/topics
    const searchTermMatches = [
      // "from [name/company]"
      lowerMessage.match(/(?:from|by)\s+(?:the\s+)?([a-z0-9\s\-\.@]+?)(?:\s+(?:in|about|and|asking|requesting)|$)/i),
      // "about [topic]"
      lowerMessage.match(/about\s+([a-z0-9\s\-]+?)(?:\s+(?:in|from|and)|$)/i),
      // "regarding [topic]"
      lowerMessage.match(/regarding\s+([a-z0-9\s\-]+?)(?:\s+(?:in|from|and)|$)/i),
      // "[topic] distributors/suppliers/etc"
      lowerMessage.match(/([a-z0-9\s\-]+?)\s+(?:distributors?|suppliers?|vendors?|companies?)/i),
      // Specific entities
      lowerMessage.match(/(patent office|uspto|activated carbon|pricing)/i),
    ];

    for (const match of searchTermMatches) {
      if (match?.[1]) {
        entities.searchTerms = entities.searchTerms || [];
        entities.searchTerms.push(match[1].trim());
      }
    }

    // Check for follow-up first (when user says "please" or "yes" after a question)
    if (followUpPatterns.some(p => p.test(lowerMessage)) && this.context.pendingAction) {
      return {
        type: 'follow_up',
        confidence: 0.95,
        entities: this.getLastEntities(),
      };
    }

    // Detect primary intent
    if (searchPatterns.some(p => p.test(lowerMessage))) {
      return { type: 'search', confidence: 0.9, entities };
    }
    
    if (composePatterns.some(p => p.test(lowerMessage))) {
      return { type: 'compose', confidence: 0.9, entities };
    }
    
    if (replyPatterns.some(p => p.test(lowerMessage))) {
      return { type: 'reply', confidence: 0.9, entities };
    }
    
    if (summarizePatterns.some(p => p.test(lowerMessage))) {
      return { type: 'summarize', confidence: 0.8, entities };
    }
    
    if (actionPatterns.some(p => p.test(lowerMessage))) {
      return { type: 'action_items', confidence: 0.8, entities };
    }
    
    if (listPatterns.some(p => p.test(lowerMessage))) {
      return { type: 'list', confidence: 0.8, entities };
    }

    return { type: 'general', confidence: 0.5, entities };
  }

  /**
   * Get entities from the last pending action
   */
  private getLastEntities(): Intent['entities'] {
    // Look back through conversation for the last search terms mentioned
    const recentMessages = this.context.conversationHistory.slice(-5);
    for (const msg of recentMessages.reverse()) {
      if (msg.metadata?.searchQuery) {
        return { searchTerms: [msg.metadata.searchQuery] };
      }
    }
    return {};
  }

  /**
   * Extract search query from user message
   */
  extractSearchQuery(message: string, entities: Intent['entities']): string[] {
    const queries: string[] = [];
    
    // Use extracted entities
    if (entities.searchTerms) {
      queries.push(...entities.searchTerms);
    }
    
    // Additional keyword extraction
    const keywordPatterns = [
      /(?:patent\s*office|uspto)/i,
      /activated\s*carbon/i,
      /distributor/i,
      /pricing|price|quote/i,
      /2026|2025|2024/i,
    ];
    
    for (const pattern of keywordPatterns) {
      const match = message.match(pattern);
      if (match && !queries.some(q => q.toLowerCase().includes(match[0].toLowerCase()))) {
        queries.push(match[0]);
      }
    }
    
    // If no specific queries found, use the whole message as a search
    if (queries.length === 0) {
      // Clean up the message for search
      const cleanedMessage = message
        .replace(/^(?:please|can you|could you|i need to|i want to)\s*/i, '')
        .replace(/\s*(?:in my inbox|from my email|please)$/i, '')
        .replace(/(?:search|find|look for|show me|get)\s*/i, '')
        .trim();
      
      if (cleanedMessage.length > 2) {
        queries.push(cleanedMessage);
      }
    }
    
    return queries;
  }

  /**
   * Add message to conversation history
   */
  addToHistory(message: ConversationMessage) {
    this.context.conversationHistory.push(message);
    
    // Trim history if too long
    if (this.context.conversationHistory.length > this.maxHistoryLength) {
      this.context.conversationHistory = this.context.conversationHistory.slice(-this.maxHistoryLength);
    }
    
    // Update mentioned topics from search results
    if (message.metadata?.searchResults) {
      this.context.lastSearchResults = message.metadata.searchResults;
    }
    
    if (message.metadata?.searchQuery) {
      this.context.lastSearchQuery = message.metadata.searchQuery;
    }
  }

  /**
   * Store search results in context
   */
  setSearchResults(results: EmailSummary[], query: string) {
    this.context.lastSearchResults = results;
    this.context.lastSearchQuery = query;
    
    // Extract contacts from results
    for (const email of results) {
      if (email.sender && email.senderEmail) {
        this.context.mentionedContacts.set(email.sender.toLowerCase(), email.senderEmail);
      }
    }
  }

  /**
   * Set pending action (when AI asks user for confirmation)
   */
  setPendingAction(action: string | null) {
    this.context.pendingAction = action;
  }

  /**
   * Get conversation history for AI context
   */
  getConversationContext(): string {
    if (this.context.conversationHistory.length === 0) {
      return '';
    }

    const recentHistory = this.context.conversationHistory.slice(-10);
    let contextStr = '\n\n--- CONVERSATION HISTORY ---\n';
    
    for (const msg of recentHistory) {
      const role = msg.role === 'user' ? 'User' : 'Assistant';
      contextStr += `${role}: ${msg.content.substring(0, 500)}${msg.content.length > 500 ? '...' : ''}\n`;
      
      if (msg.metadata?.searchResults && msg.metadata.searchResults.length > 0) {
        contextStr += `[Found ${msg.metadata.searchResults.length} emails]\n`;
      }
    }
    
    return contextStr;
  }

  /**
   * Get last search results context for AI
   */
  getSearchResultsContext(): string {
    if (this.context.lastSearchResults.length === 0) {
      return '';
    }

    let contextStr = `\n\n--- PREVIOUS SEARCH RESULTS (query: "${this.context.lastSearchQuery}") ---\n`;
    contextStr += `Found ${this.context.lastSearchResults.length} emails:\n\n`;
    
    for (const email of this.context.lastSearchResults.slice(0, 10)) {
      contextStr += `• From: ${email.sender} <${email.senderEmail}>\n`;
      contextStr += `  Subject: ${email.subject}\n`;
      contextStr += `  Date: ${email.receivedDateTime}\n`;
      contextStr += `  Preview: ${email.preview?.substring(0, 150) || 'No preview'}...\n\n`;
    }
    
    return contextStr;
  }

  /**
   * Get known contacts from conversation
   */
  getKnownContacts(): Map<string, string> {
    return this.context.mentionedContacts;
  }

  /**
   * Build full context for AI request
   */
  buildAIContext(currentEmail: EmailSummary | null): string {
    let context = '';
    
    // Add conversation history
    context += this.getConversationContext();
    
    // Add search results if available
    context += this.getSearchResultsContext();
    
    // Add current email context if available
    if (currentEmail) {
      context += `\n\n--- CURRENTLY SELECTED EMAIL ---\n`;
      context += `From: ${currentEmail.sender} <${currentEmail.senderEmail}>\n`;
      context += `Subject: ${currentEmail.subject}\n`;
      context += `Preview: ${currentEmail.preview}\n`;
    }
    
    // Add known contacts
    if (this.context.mentionedContacts.size > 0) {
      context += `\n\n--- KNOWN CONTACTS FROM INBOX ---\n`;
      for (const [name, email] of this.context.mentionedContacts) {
        context += `• ${name}: ${email}\n`;
      }
    }
    
    return context;
  }

  /**
   * Clear context (for new conversation)
   */
  clearContext() {
    this.context = {
      conversationHistory: [],
      lastSearchResults: [],
      lastSearchQuery: null,
      currentEmailContext: null,
      mentionedContacts: new Map(),
      mentionedTopics: [],
      pendingAction: null,
    };
  }

  /**
   * Get current context state
   */
  getContext(): AgentContext {
    return this.context;
  }

  /**
   * Check if we have relevant search results for a query
   */
  hasRelevantResults(query: string): boolean {
    if (this.context.lastSearchResults.length === 0) return false;
    if (!this.context.lastSearchQuery) return false;
    
    // Check if the new query is related to the last search
    const lastTerms = this.context.lastSearchQuery.toLowerCase().split(/\s+/);
    const newTerms = query.toLowerCase().split(/\s+/);
    
    return lastTerms.some(term => newTerms.some(newTerm => 
      term.includes(newTerm) || newTerm.includes(term)
    ));
  }
}

export const agentService = new AgentService();
export type { Intent, ConversationMessage, AgentContext };
