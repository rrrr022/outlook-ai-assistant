import { AIRequest, AIResponse } from '@shared/types';
import environment from '../config/environment';

/**
 * Service for communicating with the AI backend
 * Supports dynamic port discovery for local development
 */
class AIService {
  private apiBaseUrl: string | null = null;
  private portDiscoveryPromise: Promise<string> | null = null;

  /**
   * Check if we're running in localhost/development mode
   */
  private isLocalhost(): boolean {
    return typeof window !== 'undefined' && 
      (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1');
  }

  /**
   * Discover the backend API URL by trying common ports (localhost only)
   */
  private async discoverApiUrl(): Promise<string> {
    // In production, use the configured API URL directly
    if (!this.isLocalhost()) {
      console.log(`✅ Using production API: ${environment.apiUrl}`);
      return environment.apiUrl;
    }

    // For localhost, try to discover the port
    const portsToTry = [3001, 3002, 3003, 3004, 3005];
    const protocols = ['https', 'http'];

    for (const protocol of protocols) {
      for (const port of portsToTry) {
        try {
          const url = `${protocol}://localhost:${port}`;
          const response = await fetch(`${url}/health`, { 
            method: 'GET',
            signal: AbortSignal.timeout(2000),
          });
          if (response.ok) {
            console.log(`✅ Found API server at ${url}`);
            return url;
          }
        } catch (err) {
          // Try next combination
        }
      }
    }

    // Fallback - use same protocol as page
    const protocol = window.location.protocol.replace(':', '');
    console.warn(`⚠️ Could not discover API server, using ${protocol}://localhost:3001`);
    return `${protocol}://localhost:3001`;
  }

  /**
   * Get the API base URL (with caching)
   */
  private async getApiBaseUrl(): Promise<string> {
    if (this.apiBaseUrl) {
      return this.apiBaseUrl;
    }

    // Prevent multiple simultaneous discovery attempts
    if (!this.portDiscoveryPromise) {
      this.portDiscoveryPromise = this.discoverApiUrl();
    }

    this.apiBaseUrl = await this.portDiscoveryPromise;
    return this.apiBaseUrl;
  }

  private async request<T>(endpoint: string, options: RequestInit = {}): Promise<T> {
    const baseUrl = await this.getApiBaseUrl();
    const response = await fetch(`${baseUrl}/api${endpoint}`, {
      ...options,
      headers: {
        'Content-Type': 'application/json',
        ...options.headers,
      },
    });

    if (!response.ok) {
      throw new Error(`API error: ${response.statusText}`);
    }

    return response.json();
  }

  /**
   * Send a chat message to the AI
   */
  async chat(request: AIRequest): Promise<AIResponse> {
    try {
      console.log('📤 Sending AI request:', request.prompt.substring(0, 100) + '...');
      const response = await this.request<AIResponse>('/ai/chat', {
        method: 'POST',
        body: JSON.stringify(request),
      });
      console.log('📥 Got AI response:', response.content.substring(0, 100) + '...');
      return response;
    } catch (error) {
      console.error('❌ AI chat error:', error);
      // Return a fallback response for development
      return this.getFallbackResponse(request.prompt);
    }
  }

  /**
   * Summarize an email
   */
  async summarizeEmail(emailContent: string): Promise<string> {
    const response = await this.chat({
      prompt: `Please summarize this email concisely:\n\n${emailContent}`,
    });
    return response.content;
  }

  /**
   * Draft an email reply
   */
  async draftReply(originalEmail: string, instructions?: string): Promise<string> {
    const prompt = instructions
      ? `Draft a reply to this email following these instructions: "${instructions}"\n\nOriginal email:\n${originalEmail}`
      : `Draft a professional reply to this email:\n\n${originalEmail}`;

    const response = await this.chat({ prompt });
    return response.content;
  }

  /**
   * Extract tasks from email content
   */
  async extractTasks(emailContent: string): Promise<string[]> {
    const response = await this.chat({
      prompt: `Extract action items and tasks from this email. Return as a simple bullet list:\n\n${emailContent}`,
    });

    // Parse the response into task items
    const lines = response.content.split('\n');
    return lines
      .filter(line => line.trim())
      .map(line => line.replace(/^[-•*]\s*/, '').trim())
      .filter(line => line.length > 0);
  }

  /**
   * Generate a day plan based on calendar and tasks
   */
  async generateDayPlan(context: { events: any[]; tasks: any[] }): Promise<string> {
    const response = await this.chat({
      prompt: `Help me plan my day based on these events and tasks:\n\nEvents: ${JSON.stringify(context.events)}\n\nTasks: ${JSON.stringify(context.tasks)}`,
    });
    return response.content;
  }

  /**
   * Fallback response for when API is unavailable (development)
   */
  private getFallbackResponse(prompt: string): AIResponse {
    const lowercasePrompt = prompt.toLowerCase();

    if (lowercasePrompt.includes('summarize')) {
      return {
        content: 'This email discusses project updates and requests a meeting to review progress. Key points include: deadline adjustments, resource allocation, and next steps for the team.',
      };
    }

    if (lowercasePrompt.includes('reply') || lowercasePrompt.includes('draft')) {
      return {
        content: `Thank you for your email. I appreciate you reaching out about this matter.

I've reviewed the information you provided and would like to schedule a time to discuss this further. Would you be available for a brief call this week?

Please let me know your availability and I'll send a calendar invite.

Best regards`,
      };
    }

    if (lowercasePrompt.includes('task') || lowercasePrompt.includes('action')) {
      return {
        content: `Based on the email, here are the action items:
- Review the attached document by Friday
- Schedule a follow-up meeting
- Prepare status update for the team
- Send updated timeline to stakeholders`,
        extractedTasks: [
          { title: 'Review attached document', priority: 'high' },
          { title: 'Schedule follow-up meeting', priority: 'normal' },
          { title: 'Prepare status update', priority: 'normal' },
        ],
      };
    }

    if (lowercasePrompt.includes('plan') || lowercasePrompt.includes('day')) {
      return {
        content: `Here's your optimized day plan:

🌅 Morning (9:00 - 12:00)
- Start with your Team Standup at 9:00 AM
- Use the focus time until noon to tackle high-priority tasks
- Review and respond to urgent emails

🌤️ Afternoon (12:00 - 17:00)  
- Project Review meeting at 2:00 PM
- 1:1 with Manager at 4:00 PM
- Use remaining time for task completion and planning

💡 Recommendations:
- Block 30 mins for email processing
- Take a short break between meetings
- Prepare talking points before your 1:1`,
      };
    }

    return {
      content: 'I understand your request. How can I help you further with your email, calendar, or tasks? You can ask me to summarize emails, draft replies, extract action items, or help plan your day.',
    };
  }
}

export const aiService = new AIService();
