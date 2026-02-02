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
  private lastRequestTime = 0;
  private minRequestInterval = 500; // Minimum 500ms between requests

  /**
   * Throttle requests to avoid hitting rate limits
   */
  private async throttle(): Promise<void> {
    const now = Date.now();
    const elapsed = now - this.lastRequestTime;
    if (elapsed < this.minRequestInterval) {
      await new Promise(r => setTimeout(r, this.minRequestInterval - elapsed));
    }
    this.lastRequestTime = Date.now();
  }

  /**
   * Retry a function with exponential backoff for rate limits
   */
  private async withRetry<T>(
    fn: () => Promise<T>,
    maxRetries = 3,
    baseDelay = 2000
  ): Promise<T> {
    let lastError: Error | null = null;
    
    for (let attempt = 0; attempt < maxRetries; attempt++) {
      try {
        await this.throttle();
        return await fn();
      } catch (error: any) {
        lastError = error;
        const errorMsg = error?.message || '';
        
        // Check if it's a rate limit error
        if (errorMsg.includes('429') || errorMsg.toLowerCase().includes('rate limit')) {
          if (attempt < maxRetries - 1) {
            const delay = baseDelay * Math.pow(2, attempt); // 2s, 4s, 8s
            console.log(`⏳ Rate limited (attempt ${attempt + 1}/${maxRetries}), retrying in ${delay}ms...`);
            await new Promise(r => setTimeout(r, delay));
            continue;
          }
        }
        // For non-rate-limit errors, don't retry
        throw error;
      }
    }
    throw lastError || new Error('Max retries exceeded');
  }

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

      Current model: ${selectedProvider}/${selectedModel}. If the user asks which model or API is being used, answer with this explicitly.

      General knowledge is allowed. If the user asks about non-Outlook topics (e.g., weather, math, definitions), answer fully and directly. Do not refuse or redirect. If asked for real-time data, explain you don’t have live access and provide the best general guidance instead.

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

        // Truncate prompt if too long for GitHub Models (8000 token limit ≈ 30000 chars)
        let userPrompt = request.prompt;
        const maxChars = 25000; // Leave room for system prompt
        if (userPrompt.length > maxChars) {
          console.warn(`⚠️ Prompt too long (${userPrompt.length} chars), truncating to ${maxChars}`);
          userPrompt = userPrompt.substring(0, maxChars) + '\n\n[Content truncated due to length...]';
        }

        const messages = [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: userPrompt },
        ];

        // GitHub Models requires CORS proxy - route through backend
        if (selectedProvider === 'github') {
          const baseUrl = await this.getApiBaseUrl();
          
          // Use retry logic for rate limit handling
          const data = await this.withRetry(async () => {
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

            const result = await response.json();
            if (result.error) {
              throw new Error(result.error);
            }
            return result;
          });

          console.log('📥 Got GitHub Models response via proxy');
          return { content: data.response };
        }

        // Other providers can call directly (OpenAI, xAI have CORS enabled)
        // Use retry logic for rate limit handling
        const content = await this.withRetry(async () => {
          return this.directRequest(
            selectedProvider,
            userApiKey,
            selectedModel,
            messages,
            { azureEndpoint, azureDeploymentName }
          );
        });

        console.log('📥 Got direct AI response');
        return { content };
      }

      // Otherwise, use our backend proxy
      const response = await this.proxyRequest<AIResponse & { response?: string; suggestions?: string[] }>('/ai/chat', {
        method: 'POST',
        body: JSON.stringify({
          prompt: request.prompt,
          context: request.context,
          provider: selectedProvider,
          model: selectedModel,
        }),
      });

      console.log('📥 Got proxied AI response');
      return { 
        content: response.content || response.response || '',
        suggestions: response.suggestions,
        suggestedActions: response.suggestedActions,
        extractedTasks: response.extractedTasks,
        extractedEvents: response.extractedEvents,
        draftReply: response.draftReply,
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
    const errorMsg = error?.message || 'Unknown error';

    // Always show the actual error so user knows what's wrong
    console.error('🚨 AI Fallback triggered:', { error: errorMsg, promptLength: prompt.length });

    // If it's an API key error, return helpful message
    if (errorMsg.includes('401') || errorMsg.includes('unauthorized') || errorMsg.includes('invalid')) {
      return {
        content: `⚠️ **API Key Error**\n\nYour API key appears to be invalid or expired.\n\nPlease check:\n1. The key is entered correctly in Settings → API Keys\n2. The key hasn't expired\n3. The key has the required permissions\n\n_Error: ${errorMsg}_`,
      };
    }

    if (errorMsg.includes('400')) {
      return {
        content: `⚠️ **Bad Request Error**\n\nThe AI model couldn't process the request.\n\nThis could mean:\n1. The model name is incorrect\n2. The request format is invalid\n\n_Error: ${errorMsg}_\n\nTry going to Settings → Config and selecting a different model.`,
      };
    }

    if (errorMsg.includes('429') || errorMsg.includes('rate limit')) {
      return {
        content: `⚠️ **Rate Limit Reached**\n\nYou've hit the API rate limit. Please wait a moment and try again.\n\n_Error: ${errorMsg}_`,
      };
    }

    if (errorMsg.includes('Failed to fetch') || errorMsg.includes('network') || errorMsg.includes('CORS')) {
      return {
        content: `⚠️ **Network Error**\n\nCouldn't connect to the AI service.\n\nPossible causes:\n1. No internet connection\n2. The AI service is down\n3. CORS blocking (try a different provider)\n\n_Error: ${errorMsg}_`,
      };
    }

    // Default fallback with clear error info
    return {
      content: `⚠️ **AI Service Error**\n\nI couldn't process your request.\n\n**Error:** ${errorMsg}\n\n**To fix:**\n1. Go to **Settings** ⚙️ (Config tab)\n2. Make sure you have an API key entered\n3. Test the connection\n\nIf using **Anthropic/Claude**, make sure the model name is correct (e.g., "claude-3-5-sonnet-latest")`,
    };
  }
}

export const aiService = new AIService();
