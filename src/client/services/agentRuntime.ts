/**
 * Agent Runtime - Manages autonomous AI agent execution
 * The AI has full control and can call tools, maintain context, and work autonomously
 */

import { aiService } from './aiService';
import {
  generateToolsPrompt,
  parseToolCalls,
  executeTool,
  toolRequiresApproval,
  getTool,
  ToolResult,
} from './agentTools';
import { graphService } from './graphService';

export interface AgentMessage {
  role: 'user' | 'assistant' | 'system' | 'tool_result';
  content: string;
  toolCall?: {
    tool: string;
    params: Record<string, any>;
    result?: ToolResult;
  };
  timestamp: Date;
}

export interface PendingApproval {
  id: string;
  tool: string;
  params: Record<string, any>;
  description: string;
}

export interface AgentState {
  conversationHistory: AgentMessage[];
  isProcessing: boolean;
  pendingApprovals: PendingApproval[];
  currentTask?: string;
  memory: Record<string, any>; // Agent's working memory
}

type StateUpdateCallback = (state: AgentState) => void;
type MessageCallback = (message: string, isIntermediate?: boolean) => void;

class AgentRuntime {
  private state: AgentState = {
    conversationHistory: [],
    isProcessing: false,
    pendingApprovals: [],
    memory: {},
  };

  private stateUpdateCallbacks: StateUpdateCallback[] = [];
  private messageCallbacks: MessageCallback[] = [];
  private maxIterations = 10; // Prevent infinite loops

  /**
   * Subscribe to state updates
   */
  onStateUpdate(callback: StateUpdateCallback): () => void {
    this.stateUpdateCallbacks.push(callback);
    return () => {
      this.stateUpdateCallbacks = this.stateUpdateCallbacks.filter(cb => cb !== callback);
    };
  }

  /**
   * Subscribe to messages (for UI updates)
   */
  onMessage(callback: MessageCallback): () => void {
    this.messageCallbacks.push(callback);
    return () => {
      this.messageCallbacks = this.messageCallbacks.filter(cb => cb !== callback);
    };
  }

  private notifyStateUpdate(): void {
    for (const callback of this.stateUpdateCallbacks) {
      callback({ ...this.state });
    }
  }

  private notifyMessage(message: string, isIntermediate = false): void {
    for (const callback of this.messageCallbacks) {
      callback(message, isIntermediate);
    }
  }

  /**
   * Build the system prompt that gives the AI full context and capabilities
   */
  private buildSystemPrompt(): string {
    return `You are an intelligent AI assistant integrated into Microsoft Outlook. You have FULL AUTONOMOUS CONTROL over the user's email, calendar, and tasks through the tools provided below.

## YOUR CAPABILITIES
- Search and read emails
- Draft, send, reply to, and delete emails (with user approval for destructive actions)
- View and manage calendar events
- Create tasks and reminders
- Access user profile information

## YOUR BEHAVIOR
1. **Be Proactive**: Don't just answer questions - take action to help the user
2. **Work Autonomously**: Use tools in sequence to complete complex tasks without asking for permission on every step
3. **Maintain Context**: Remember what you've learned from previous tool calls and user messages
4. **Confirm Destructive Actions**: Always request approval before sending emails, deleting, or creating calendar events
5. **Be Efficient**: Batch related operations when possible
6. **Communicate Progress**: For multi-step tasks, briefly explain what you're doing

## CURRENT CONTEXT
- Date/Time: ${new Date().toLocaleString()}
- Day: ${new Date().toLocaleDateString('en-US', { weekday: 'long' })}
${this.state.memory.userProfile ? `- User: ${this.state.memory.userProfile.displayName} (${this.state.memory.userProfile.mail})` : ''}

## WORKING MEMORY
${Object.keys(this.state.memory).length > 0 ? JSON.stringify(this.state.memory, null, 2) : 'Empty - gather information as needed'}

${generateToolsPrompt()}

Remember: You are in control. Take initiative, use your tools, and help the user accomplish their goals efficiently.`;
  }

  /**
   * Format conversation history for the AI
   */
  private formatConversationForAI(): { role: string; content: string }[] {
    const messages: { role: string; content: string }[] = [];

    for (const msg of this.state.conversationHistory) {
      if (msg.role === 'tool_result') {
        // Include tool results as system/assistant context
        messages.push({
          role: 'assistant',
          content: `[Tool Result for ${msg.toolCall?.tool}]:\n${JSON.stringify(msg.toolCall?.result, null, 2)}`,
        });
      } else {
        messages.push({
          role: msg.role,
          content: msg.content,
        });
      }
    }

    return messages;
  }

  /**
   * Process a user message - this is the main entry point
   */
  async processUserMessage(userMessage: string): Promise<string> {
    if (this.state.isProcessing) {
      return "I'm still processing the previous request. Please wait...";
    }

    this.state.isProcessing = true;
    this.state.currentTask = userMessage;
    this.notifyStateUpdate();

    // Add user message to history
    this.state.conversationHistory.push({
      role: 'user',
      content: userMessage,
      timestamp: new Date(),
    });

    try {
      // Run the agent loop
      const response = await this.runAgentLoop();
      return response;
    } catch (error: any) {
      console.error('Agent error:', error);
      return `I encountered an error: ${error.message}. Please try again.`;
    } finally {
      this.state.isProcessing = false;
      this.state.currentTask = undefined;
      this.notifyStateUpdate();
    }
  }

  /**
   * The main agent loop - AI calls tools until task is complete
   */
  private async runAgentLoop(): Promise<string> {
    let iterations = 0;
    let finalResponse = '';

    while (iterations < this.maxIterations) {
      iterations++;

      // Build the full prompt
      const systemPrompt = this.buildSystemPrompt();
      const conversationMessages = this.formatConversationForAI();

      // Get AI response
      let aiResponse: string;
      try {
        // Build the full prompt with system context and conversation history
        const fullPrompt = `${systemPrompt}

---
CONVERSATION HISTORY:
${conversationMessages.map(m => `${m.role.toUpperCase()}: ${m.content}`).join('\n\n')}
---

Now respond to the user's latest message. Use tools if needed, or respond directly.`;

        const response = await aiService.chat({ prompt: fullPrompt });
        aiResponse = response.content;
      } catch (error: any) {
        console.error('AI call failed:', error);
        return `Failed to get AI response: ${error.message}`;
      }

      // Parse any tool calls from the response
      const toolCalls = parseToolCalls(aiResponse);

      // Extract the text response (everything except tool blocks)
      const textResponse = aiResponse
        .replace(/```tool[\s\S]*?```/g, '')
        .trim();

      // If there are no tool calls, this is the final response
      if (toolCalls.length === 0) {
        finalResponse = textResponse;

        // Add to history
        this.state.conversationHistory.push({
          role: 'assistant',
          content: aiResponse,
          timestamp: new Date(),
        });
        this.notifyStateUpdate();

        break;
      }

      // Process tool calls
      if (textResponse) {
        this.notifyMessage(textResponse, true);
      }

      // Add assistant response (with tool calls) to history
      this.state.conversationHistory.push({
        role: 'assistant',
        content: aiResponse,
        timestamp: new Date(),
      });

      // Execute each tool call
      for (const call of toolCalls) {
        // Check if tool requires approval
        if (toolRequiresApproval(call.tool)) {
          const tool = getTool(call.tool);
          const approval: PendingApproval = {
            id: `${Date.now()}-${call.tool}`,
            tool: call.tool,
            params: call.params,
            description: tool?.description || call.tool,
          };
          this.state.pendingApprovals.push(approval);
          this.notifyStateUpdate();

          // Add pending approval to history so AI knows
          this.state.conversationHistory.push({
            role: 'tool_result',
            content: `Awaiting user approval for: ${call.tool}`,
            toolCall: {
              tool: call.tool,
              params: call.params,
              result: {
                success: false,
                data: { status: 'pending_approval', approvalId: approval.id },
              },
            },
            timestamp: new Date(),
          });

          // Return and wait for approval
          finalResponse = textResponse || `I need your approval to ${call.tool}. Please review and approve or reject the action.`;

          // Store context for continuation
          this.state.memory.pendingToolCall = call;
          this.notifyStateUpdate();

          return finalResponse;
        }

        // Execute the tool
        this.notifyMessage(`🔧 Executing: ${call.tool}...`, true);
        const result = await executeTool(call.tool, call.params);

        // Add result to history
        this.state.conversationHistory.push({
          role: 'tool_result',
          content: JSON.stringify(result),
          toolCall: {
            tool: call.tool,
            params: call.params,
            result,
          },
          timestamp: new Date(),
        });

        // Update memory with useful data
        if (result.success && result.data) {
          this.updateMemory(call.tool, result.data);
        }

        this.notifyStateUpdate();
      }

      // Continue the loop - AI will see the tool results and decide next step
    }

    if (iterations >= this.maxIterations) {
      return finalResponse || "I've reached my iteration limit. Here's what I've accomplished so far. Please let me know if you need me to continue.";
    }

    return finalResponse;
  }

  /**
   * Update agent memory with useful information from tool results
   */
  private updateMemory(tool: string, data: any): void {
    switch (tool) {
      case 'get_user_info':
        this.state.memory.userProfile = data;
        break;
      case 'search_emails':
        this.state.memory.lastSearchResults = data;
        break;
      case 'get_calendar_events':
        this.state.memory.upcomingEvents = data;
        break;
      case 'get_recent_emails':
        this.state.memory.recentEmails = data;
        break;
    }
  }

  /**
   * Handle user approval of a pending action
   */
  async approveAction(approvalId: string): Promise<string> {
    const approval = this.state.pendingApprovals.find(a => a.id === approvalId);
    if (!approval) {
      return 'Approval not found or already processed.';
    }

    // Remove from pending
    this.state.pendingApprovals = this.state.pendingApprovals.filter(a => a.id !== approvalId);
    this.notifyStateUpdate();

    // Execute the approved action
    this.state.isProcessing = true;
    this.notifyStateUpdate();

    try {
      this.notifyMessage(`✅ Executing approved action: ${approval.tool}...`, true);
      const result = await executeTool(approval.tool, approval.params);

      // Add to history
      this.state.conversationHistory.push({
        role: 'tool_result',
        content: `[APPROVED AND EXECUTED] ${approval.tool}`,
        toolCall: {
          tool: approval.tool,
          params: approval.params,
          result,
        },
        timestamp: new Date(),
      });

      // Clear pending context
      delete this.state.memory.pendingToolCall;
      this.notifyStateUpdate();

      // Continue the agent loop to process the result
      return await this.runAgentLoop();
    } finally {
      this.state.isProcessing = false;
      this.notifyStateUpdate();
    }
  }

  /**
   * Handle user rejection of a pending action
   */
  async rejectAction(approvalId: string, reason?: string): Promise<string> {
    const approval = this.state.pendingApprovals.find(a => a.id === approvalId);
    if (!approval) {
      return 'Action not found or already processed.';
    }

    // Remove from pending
    this.state.pendingApprovals = this.state.pendingApprovals.filter(a => a.id !== approvalId);

    // Add rejection to history
    this.state.conversationHistory.push({
      role: 'tool_result',
      content: `[REJECTED BY USER] ${approval.tool}${reason ? `: ${reason}` : ''}`,
      toolCall: {
        tool: approval.tool,
        params: approval.params,
        result: { success: false, error: `User rejected: ${reason || 'No reason provided'}` },
      },
      timestamp: new Date(),
    });

    // Clear pending context
    delete this.state.memory.pendingToolCall;
    this.notifyStateUpdate();

    // Continue to let AI respond to rejection
    this.state.isProcessing = true;
    try {
      return await this.runAgentLoop();
    } finally {
      this.state.isProcessing = false;
      this.notifyStateUpdate();
    }
  }

  /**
   * Get current state
   */
  getState(): AgentState {
    return { ...this.state };
  }

  /**
   * Clear conversation and memory
   */
  clearConversation(): void {
    this.state.conversationHistory = [];
    this.state.memory = {};
    this.state.pendingApprovals = [];
    this.state.currentTask = undefined;
    this.notifyStateUpdate();
  }

  /**
   * Initialize agent with user context
   */
  async initialize(): Promise<void> {
    try {
      // Pre-fetch user info for context
      const userInfo = graphService.getUserInfo();
      if (userInfo) {
        this.state.memory.userProfile = userInfo;
      }
    } catch (error) {
      console.log('Could not pre-fetch user profile');
    }
  }
}

// Export singleton instance
export const agentRuntime = new AgentRuntime();
