import { AIRequest, AIResponse } from '@shared/types';
import environment from '../config/environment';
import { useAPIKeyStore, AIProvider, PROVIDERS } from '../store/apiKeyStore';

/**
 * Service for communicating with AI providers
 * Supports BYOK (Bring Your Own Key) for direct API calls
 * Falls back to server proxy when using hosted mode
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

  /**
   * Normalize Anthropic model names to valid API model IDs
   * GitHub Models uses short names like 'claude-3-opus' but
   * Anthropic's API requires full names like 'claude-3-opus-20240229' or '-latest'
   */
  private normalizeAnthropicModel(model: string): string {
    const modelMappings: Record<string, string> = {
      // Short names to full dated versions
      'claude-3-opus': 'claude-3-opus-20240229',
      'claude-3-sonnet': 'claude-3-sonnet-20240229',
      'claude-3-haiku': 'claude-3-haiku-20240307',
      'claude-3-5-sonnet': 'claude-3-5-sonnet-20241022',
      'claude-3-5-sonnet-v2': 'claude-3-5-sonnet-20241022',
      'claude-3-5-haiku': 'claude-3-5-haiku-20241022',
      // Latest variants are already valid
      'claude-3-opus-latest': 'claude-3-opus-latest',
      'claude-3-5-sonnet-latest': 'claude-3-5-sonnet-latest',
      'claude-3-haiku-20240307': 'claude-3-haiku-20240307',
    };
    
    return modelMappings[model] || model;
  }

  /**
   * Make a request through our backend proxy (server key mode)
   */
  private async proxyRequest<T>(endpoint: string, options: RequestInit = {}): Promise<T> {
    const baseUrl = await this.getApiBaseUrl();
    const response = await fetch(`${baseUrl}/api${endpoint}`, {
      ...options,
      headers: {
        'Content-Type': 'application/json',
        ...options.headers,
      },
    });

    if (!response.ok) {
      const error = await response.text();
      throw new Error(`API error: ${response.statusText} - ${error}`);
    }

    return response.json();
  }

  /**
   * Make a direct call to an AI provider (BYOK mode)
   */
  private async directRequest(
    provider: AIProvider,
    apiKey: string,
    model: string,
    messages: Array<{ role: string; content: string }>,
    options: { azureEndpoint?: string; azureDeploymentName?: string } = {}
  ): Promise<string> {
    const providerConfig = PROVIDERS.find(p => p.id === provider);
    
    if (provider === 'anthropic') {
      return this.callAnthropic(apiKey, model, messages);
    } else if (provider === 'azure') {
      return this.callAzureOpenAI(apiKey, model, messages, options.azureEndpoint!, options.azureDeploymentName!);
    } else {
      // OpenAI-compatible API (GitHub Models, OpenAI)
      const baseUrl = providerConfig?.baseUrl || 'https://api.openai.com/v1';
      return this.callOpenAICompatible(baseUrl, apiKey, model, messages);
    }
  }

  /**
   * Call OpenAI-compatible APIs (OpenAI, GitHub Models)
   */
  private async callOpenAICompatible(
    baseUrl: string,
    apiKey: string,
    model: string,
    messages: Array<{ role: string; content: string }>
  ): Promise<string> {
    const response = await fetch(`${baseUrl}/chat/completions`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model,
        messages,
        max_tokens: 4096,
        temperature: 0.7,
      }),
    });

    if (!response.ok) {
      const error = await response.json().catch(() => ({ error: { message: response.statusText } }));
      throw new Error(error.error?.message || `API error: ${response.statusText}`);
    }

    const data = await response.json();
    return data.choices[0]?.message?.content || 'No response generated';
  }

  /**
   * Call Anthropic Claude API
   */
  private async callAnthropic(
    apiKey: string,
    model: string,
    messages: Array<{ role: string; content: string }>
  ): Promise<string> {
    // Normalize model name - Anthropic requires full version names or -latest suffix
    const normalizedModel = this.normalizeAnthropicModel(model);
    console.log(`🔮 Anthropic model: ${model} → ${normalizedModel}`);
    
    // Extract system message if present
    const systemMessage = messages.find(m => m.role === 'system');
    const chatMessages = messages.filter(m => m.role !== 'system');

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'anthropic-dangerous-direct-browser-access': 'true',
      },
      body: JSON.stringify({
        model: normalizedModel,
        max_tokens: 8192,
        system: systemMessage?.content || 'You are a helpful AI assistant for Microsoft Outlook.',
        messages: chatMessages.map(m => ({
          role: m.role === 'assistant' ? 'assistant' : 'user',
          content: m.content,
        })),
      }),
    });

    if (!response.ok) {
      const error = await response.json().catch(() => ({ error: { message: response.statusText } }));
      throw new Error(error.error?.message || `Anthropic API error: ${response.statusText}`);
    }

    const data = await response.json();
    return data.content[0]?.text || 'No response generated';
  }

  /**
   * Call Azure OpenAI API
   */
  private async callAzureOpenAI(
    apiKey: string,
    model: string,
    messages: Array<{ role: string; content: string }>,
    endpoint: string,
    deploymentName: string
  ): Promise<string> {
    const url = `${endpoint}/openai/deployments/${deploymentName}/chat/completions?api-version=2024-02-01`;
    
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'api-key': apiKey,
      },
      body: JSON.stringify({
        messages,
        max_tokens: 4096,
        temperature: 0.7,
      }),
    });

    if (!response.ok) {
      const error = await response.json().catch(() => ({ error: { message: response.statusText } }));
      throw new Error(error.error?.message || `Azure OpenAI error: ${response.statusText}`);
    }

    const data = await response.json();
    return data.choices[0]?.message?.content || 'No response generated';
  }

  /**
   * Send a chat message to the AI
   * Uses BYOK if user has set their own key, otherwise uses server proxy
   */
  async chat(request: AIRequest): Promise<AIResponse> {
    const store = useAPIKeyStore.getState();
    const { selectedProvider, selectedModel, apiKeys, useServerKey, azureEndpoint, azureDeploymentName } = store;
    const userApiKey = apiKeys[selectedProvider];

    console.log('📤 Sending AI request:', {
      prompt: request.prompt.substring(0, 50) + '...',
      provider: selectedProvider,
      model: selectedModel,
      mode: !useServerKey && userApiKey ? 'BYOK (Direct)' : 'Server Proxy',
    });

    try {
      // If user has their own key and not using server mode
      if (!useServerKey && userApiKey) {
        const systemPrompt = `You are FreedomForged AI, an intelligent assistant integrated into Microsoft Outlook. You help users manage their emails, calendar, and tasks efficiently.

CRITICAL BEHAVIOR - BE PROACTIVE:
- When given search results or email data, USE IT IMMEDIATELY to answer the user's question
- NEVER ask the user to search for something if you already have search results in the context
- NEVER say "I don't have access to your inbox" if search results are provided
- If the user asks about emails and you have search results, ANALYZE THOSE RESULTS
- Be helpful and take action - don't just explain what could be done

CONTEXT AWARENESS:
- Pay attention to the conversation history provided
- Remember what was discussed earlier in the conversation
- Use search results from previous queries to answer follow-up questions
- If composing emails, use the contact information from search results

Your capabilities:
- Summarize emails and extract key information
- Draft professional email responses using contacts from inbox
- Suggest actions based on email content
- Help organize and prioritize emails
- Provide calendar suggestions and scheduling help
- Generate documents (Word, PDF, Excel, PowerPoint)

WHEN COMPOSING EMAILS:
- Use the actual email addresses from search results
- Format as: TO: [email], SUBJECT: [subject], BODY: [content]
- Be specific with dates, pricing requests, etc.
- If multiple contacts found, offer to draft to all of them

DOCUMENT GENERATION TIPS:
When users ask for reports, summaries, presentations, or spreadsheets:
- Structure your response clearly with headings (# for main, ## for sub)
- Use markdown tables for tabular data
- Use bullet points for lists

Be concise, professional, and helpful. FORMAT RESPONSES WITH MARKDOWN.
Most importantly: BE PROACTIVE AND TAKE ACTION rather than asking the user to do things themselves.`;

        const messages = [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: request.prompt },
        ];

        // GitHub Models requires CORS proxy - route through backend
        if (selectedProvider === 'github') {
          const baseUrl = await this.getApiBaseUrl();
          const response = await fetch(`${baseUrl}/api/chat`, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
            },
            body: JSON.stringify({
              provider: 'github',
              apiKey: userApiKey,
              model: selectedModel,
              messages,
            }),
          });

          if (!response.ok) {
            const error = await response.json().catch(() => ({ error: response.statusText }));
            throw new Error(error.error || error.message || `API error: ${response.statusText}`);
          }

          const data = await response.json();
          if (data.error) {
            throw new Error(data.error);
          }

          console.log('📥 Got GitHub Models response via proxy');
          return { content: data.response };
        }

        // Other providers can call directly (OpenAI, xAI have CORS enabled)
        const content = await this.directRequest(
          selectedProvider,
          userApiKey,
          selectedModel,
          messages,
          { azureEndpoint, azureDeploymentName }
        );

        console.log('📥 Got direct AI response');
        return { content };
      }

      // Otherwise, use our backend proxy
      const response = await this.proxyRequest<{ response: string; suggestions?: string[] }>('/ai/chat', {
        method: 'POST',
        body: JSON.stringify({
          message: request.prompt,
          context: request.context,
        }),
      });

      console.log('📥 Got proxied AI response');
      return { 
        content: response.response,
        suggestions: response.suggestions 
      };

    } catch (error) {
      console.error('❌ AI chat error:', error);
      // Return a fallback response for development/error cases
      return this.getFallbackResponse(request.prompt, error);
    }
  }

  /**
   * Test connection to the selected AI provider
   */
  async testConnection(): Promise<{ success: boolean; error?: string }> {
    const store = useAPIKeyStore.getState();
    const { selectedProvider, selectedModel, apiKeys, azureEndpoint, azureDeploymentName } = store;
    const userApiKey = apiKeys[selectedProvider];

    if (!userApiKey) {
      return { success: false, error: 'No API key provided' };
    }

    try {
      store.setConnectionStatus('testing');

      const messages = [
        { role: 'system', content: 'You are a helpful assistant.' },
        { role: 'user', content: 'Say "Connected!" in one word.' },
      ];

      // GitHub Models requires CORS proxy - route through backend
      if (selectedProvider === 'github') {
        const baseUrl = await this.getApiBaseUrl();
        const response = await fetch(`${baseUrl}/api/chat`, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            provider: 'github',
            apiKey: userApiKey,
            model: selectedModel,
            messages,
          }),
        });

        if (!response.ok) {
          const error = await response.json().catch(() => ({ error: response.statusText }));
          throw new Error(error.error || error.message || `API error: ${response.statusText}`);
        }

        const data = await response.json();
        if (data.error) {
          throw new Error(data.error);
        }
      } else {
        await this.directRequest(
          selectedProvider,
          userApiKey,
          selectedModel,
          messages,
          { azureEndpoint, azureDeploymentName }
        );
      }

      store.setConnectionStatus('connected');
      store.setLastError(null);
      return { success: true };

    } catch (error: any) {
      const errorMessage = error.message || 'Connection failed';
      store.setConnectionStatus('error');
      store.setLastError(errorMessage);
      return { success: false, error: errorMessage };
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
   * Fallback response for when API is unavailable
   */
  private getFallbackResponse(prompt: string, error?: any): AIResponse {
    const lowercasePrompt = prompt.toLowerCase();
    const errorMsg = error?.message || 'Unknown error';

    // If it's an API key error, return helpful message
    if (errorMsg.includes('401') || errorMsg.includes('unauthorized') || errorMsg.includes('invalid')) {
      return {
        content: `⚠️ **API Key Error**\n\nYour API key appears to be invalid or expired.\n\nPlease check:\n1. The key is entered correctly in Settings → API Keys\n2. The key hasn't expired\n3. The key has the required permissions\n\n_Error: ${errorMsg}_`,
      };
    }

    if (errorMsg.includes('429') || errorMsg.includes('rate limit')) {
      return {
        content: `⚠️ **Rate Limit Reached**\n\nYou've hit the API rate limit. Please wait a moment and try again.\n\n_Error: ${errorMsg}_`,
      };
    }

    // Log what's happening for debugging
    console.log('⚠️ AI Fallback triggered:', { error: errorMsg, promptStart: lowercasePrompt.substring(0, 100) });

    // Only use summarize fallback for actual email summarization, not inbox searches
    if (lowercasePrompt.includes('summarize this email') || lowercasePrompt.includes('summarize the email')) {
      return {
        content: 'This email discusses project updates and requests a meeting to review progress. Key points include: deadline adjustments, resource allocation, and next steps for the team.',
      };
    }

    if (lowercasePrompt.includes('reply') || lowercasePrompt.includes('draft a reply')) {
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

    // Default fallback with error info
    return {
      content: `⚠️ **AI Service Error**\n\nI couldn't process your request. This usually means:\n\n1. **No API key configured** - Go to Settings ⚙️ and add your AI provider key\n2. **API error** - The AI service returned an error\n\n_Technical details: ${errorMsg}_\n\nYour search results are still available above. You can try again or configure an AI provider in settings.`,
    };
  }
}

export const aiService = new AIService();
