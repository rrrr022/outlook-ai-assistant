import { AIRequest, AIResponse } from '../../../shared/types';

/**
 * GitHub Models Provider
 * Uses GitHub's AI Models API (GPT-4, Claude, etc.) with your GitHub token
 * 
 * Available models:
 * - gpt-4o
 * - gpt-4o-mini  
 * - o1-preview
 * - o1-mini
 * - claude-3-5-sonnet
 * - claude-3-5-haiku (coming soon)
 * - And more...
 * 
 * @see https://github.com/marketplace/models
 */
export class GitHubModelsProvider {
  private token: string | null = null;
  private model: string;
  private baseUrl = 'https://models.inference.ai.azure.com';

  constructor() {
    this.token = process.env.GITHUB_TOKEN || null;
    this.model = process.env.GITHUB_MODEL || 'gpt-4o';
    console.log(`🤖 GitHub Models Provider initialized. Token: ${this.token ? 'Found (' + this.token.substring(0, 10) + '...)' : 'NOT FOUND'}, Model: ${this.model}`);
  }

  async chat(request: AIRequest): Promise<AIResponse> {
    console.log(`📝 AI Request received. Token exists: ${!!this.token}`);
    
    if (!this.token) {
      console.warn('❌ GitHub token not configured. Using mock response.');
      return this.getMockResponse(request.prompt);
    }

    try {
      console.log(`🚀 Calling GitHub Models API with model: ${this.model}`);
      
      // Build the system message with context
      let systemMessage = `You are an intelligent assistant integrated into Microsoft Outlook. 
You help users manage their emails, calendar, and tasks efficiently.
You provide concise, actionable responses.
When asked to draft emails, use a professional tone unless otherwise specified.
When extracting tasks, be specific and actionable.
IMPORTANT: Answer the user's actual question. Do not provide generic email templates unless asked.`;

      if (request.context) {
        systemMessage += '\n\nContext:';
        if (request.context.currentEmail) {
          systemMessage += `\nCurrent Email: From ${request.context.currentEmail.sender}, Subject: "${request.context.currentEmail.subject}"`;
        }
        if (request.context.upcomingEvents?.length) {
          systemMessage += `\nUpcoming Events: ${request.context.upcomingEvents.length} events`;
        }
        if (request.context.pendingTasks?.length) {
          systemMessage += `\nPending Tasks: ${request.context.pendingTasks.length} tasks`;
        }
      }

      const response = await fetch(`${this.baseUrl}/chat/completions`, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${this.token}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          model: this.model,
          messages: [
            { role: 'system', content: systemMessage },
            { role: 'user', content: request.prompt },
          ],
          temperature: 0.7,
          max_tokens: 1000,
        }),
      });

      console.log(`📡 API Response status: ${response.status}`);

      if (!response.ok) {
        const error = await response.text();
        console.error(`❌ API Error: ${response.status} - ${error}`);
        // If API fails, return a helpful error message instead of mock
        return {
          content: `I apologize, but I'm having trouble connecting to the AI service right now. Error: ${response.status}. Please try again in a moment.`,
          suggestedActions: [],
        };
      }

      const data: any = await response.json();
      const content = data.choices?.[0]?.message?.content || '';
      console.log(`✅ Got AI response, length: ${content.length} chars`);

      // Parse response for suggested actions and extracted data
      const aiResponse: AIResponse = {
        content,
        suggestedActions: this.parseSuggestedActions(content, request),
        extractedTasks: this.parseExtractedTasks(content),
      };

      return aiResponse;
    } catch (error: any) {
      console.error('❌ GitHub Models API error:', error);
      // Return friendly error instead of throwing
      return {
        content: `I apologize, but I encountered an error: ${error.message}. Please check that your AI configuration is correct.`,
        suggestedActions: [],
      };
    }
  }

  /**
   * List available models from GitHub Models
   */
  async listModels(): Promise<string[]> {
    // These are the currently available models on GitHub Models
    return [
      'gpt-4o',
      'gpt-4o-mini',
      'o1-preview',
      'o1-mini',
      'claude-3-5-sonnet',
      'meta-llama-3.1-405b-instruct',
      'meta-llama-3.1-70b-instruct',
      'mistral-large',
      'mistral-small',
      'cohere-command-r-plus',
    ];
  }

  private parseSuggestedActions(content: string, request: AIRequest): AIResponse['suggestedActions'] {
    const actions: AIResponse['suggestedActions'] = [];

    if (content.toLowerCase().includes('dear') || 
        content.toLowerCase().includes('hi ') ||
        content.toLowerCase().includes('hello')) {
      actions.push({
        type: 'insertReply',
        label: 'Insert as Reply',
        description: 'Insert this text into your email reply',
        parameters: { text: content },
      });
    }

    if (content.toLowerCase().includes('schedule') || 
        content.toLowerCase().includes('meeting')) {
      actions.push({
        type: 'scheduleMeeting',
        label: 'Schedule Meeting',
        description: 'Create a calendar event',
        parameters: {},
      });
    }

    return actions;
  }

  private parseExtractedTasks(content: string): AIResponse['extractedTasks'] {
    const taskPatterns = [
      /[-•*]\s*(.+)/g,
      /\d+[.)]\s*(.+)/g,
    ];

    const tasks: AIResponse['extractedTasks'] = [];

    for (const pattern of taskPatterns) {
      const matches = content.matchAll(pattern);
      for (const match of matches) {
        if (match[1] && match[1].trim().length > 5) {
          tasks.push({
            title: match[1].trim(),
            priority: this.detectPriority(match[1]),
          });
        }
      }
    }

    return tasks.length > 0 ? tasks : undefined;
  }

  private detectPriority(text: string): 'low' | 'normal' | 'high' {
    const lowercaseText = text.toLowerCase();
    if (lowercaseText.includes('urgent') || 
        lowercaseText.includes('asap') || 
        lowercaseText.includes('immediately')) {
      return 'high';
    }
    if (lowercaseText.includes('when possible') || 
        lowercaseText.includes('low priority')) {
      return 'low';
    }
    return 'normal';
  }

  private getMockResponse(prompt: string): AIResponse {
    const lowercasePrompt = prompt.toLowerCase();

    if (lowercasePrompt.includes('summarize')) {
      return {
        content: `**Email Summary:**

This email discusses project updates and contains the following key points:
- Status update on current deliverables
- Request for a follow-up meeting
- Action items requiring your attention

**Key Dates:** No specific deadlines mentioned.

**Required Actions:** Review attached documents and respond with availability for meeting.`,
      };
    }

    if (lowercasePrompt.includes('reply') || lowercasePrompt.includes('draft')) {
      return {
        content: `Thank you for your email.

I've reviewed the information you shared and appreciate the detailed update. I'd like to schedule a brief call to discuss the next steps.

Would you be available for a 30-minute meeting this week? Please let me know your preferred times and I'll send a calendar invite.

Best regards`,
        suggestedActions: [{
          type: 'insertReply',
          label: 'Insert as Reply',
          description: 'Insert this draft into your reply',
          parameters: {},
        }],
      };
    }

    return {
      content: `I understand you're asking about: "${prompt.substring(0, 100)}..."

As your AI assistant, I can help you with:
- **Email Management**: Summarize emails, draft replies, extract action items
- **Calendar**: View upcoming events, plan your day, schedule meetings
- **Tasks**: Create tasks, prioritize work, track progress

How can I assist you today?`,
    };
  }
}
