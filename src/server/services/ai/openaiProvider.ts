import OpenAI from 'openai';
import { AIRequest, AIResponse } from '../../../shared/types';

export class OpenAIProvider {
  private client: OpenAI | null = null;
  private model: string;

  constructor() {
    this.model = process.env.OPENAI_MODEL || 'gpt-4-turbo-preview';
    
    if (process.env.OPENAI_API_KEY) {
      this.client = new OpenAI({
        apiKey: process.env.OPENAI_API_KEY,
      });
    }
  }

  async chat(request: AIRequest): Promise<AIResponse> {
    if (!this.client) {
      console.warn('OpenAI client not initialized. Using mock response.');
      return this.getMockResponse(request.prompt);
    }

    try {
      // Build the system message with context
      let systemMessage = `You are an intelligent assistant integrated into Microsoft Outlook. 
You help users manage their emails, calendar, and tasks efficiently.
You provide concise, actionable responses.
When asked to draft emails, use a professional tone unless otherwise specified.
When extracting tasks, be specific and actionable.`;

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

      const completion = await this.client.chat.completions.create({
        model: this.model,
        messages: [
          { role: 'system', content: systemMessage },
          { role: 'user', content: request.prompt },
        ],
        temperature: 0.7,
        max_tokens: 4096,
      });

      const content = completion.choices[0]?.message?.content || '';

      // Parse response for suggested actions and extracted data
      const response: AIResponse = {
        content,
        suggestedActions: this.parseSuggestedActions(content, request),
        extractedTasks: this.parseExtractedTasks(content),
      };

      return response;
    } catch (error: any) {
      console.error('OpenAI API error:', error);
      throw new Error(`OpenAI API error: ${error.message}`);
    }
  }

  private parseSuggestedActions(content: string, request: AIRequest): AIResponse['suggestedActions'] {
    const actions: AIResponse['suggestedActions'] = [];

    // Detect if response contains draft reply
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

    // Detect if response contains meeting suggestion
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
    // Simple heuristic to detect task-like content
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
    // Provide useful mock responses for development
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
