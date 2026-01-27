import { Router, Request, Response } from 'express';
import { OpenAIProvider } from '../services/ai/openaiProvider';
import { AnthropicProvider } from '../services/ai/anthropicProvider';
import { GitHubModelsProvider } from '../services/ai/githubModelsProvider';
import { AIRequest, AIResponse } from '../../shared/types';

const router = Router();

// Initialize AI providers
const openaiProvider = new OpenAIProvider();
const anthropicProvider = new AnthropicProvider();
const githubModelsProvider = new GitHubModelsProvider();

/**
 * Get the appropriate AI provider based on configuration
 */
function getProvider(providerName?: string) {
  const provider = providerName || process.env.DEFAULT_AI_PROVIDER || 'github';
  
  switch (provider) {
    case 'anthropic':
      return anthropicProvider;
    case 'openai':
      return openaiProvider;
    case 'github':
    default:
      return githubModelsProvider;
  }
}

/**
 * POST /api/ai/chat
 * Main chat endpoint for AI interactions
 */
router.post('/chat', async (req: Request, res: Response) => {
  try {
    const request: AIRequest = req.body;
    const provider = getProvider(req.body.provider);

    if (!request.prompt) {
      return res.status(400).json({ 
        success: false, 
        error: 'Prompt is required' 
      });
    }

    const response = await provider.chat(request);
    res.json(response);
  } catch (error: any) {
    console.error('AI chat error:', error);
    res.status(500).json({ 
      success: false, 
      error: 'Failed to process AI request',
      message: error.message,
    });
  }
});

/**
 * POST /api/ai/summarize
 * Summarize email content
 */
router.post('/summarize', async (req: Request, res: Response) => {
  try {
    const { content, provider: providerName } = req.body;

    if (!content) {
      return res.status(400).json({ 
        success: false, 
        error: 'Content is required' 
      });
    }

    const request: AIRequest = {
      prompt: `Please provide a concise summary of the following email. Focus on key points, action items, and important dates:\n\n${content}`,
    };

    const provider = getProvider(providerName);
    const response = await provider.chat(request);

    res.json({
      success: true,
      summary: response.content,
    });
  } catch (error: any) {
    console.error('Summarize error:', error);
    res.status(500).json({ 
      success: false, 
      error: 'Failed to summarize content' 
    });
  }
});

/**
 * POST /api/ai/draft-reply
 * Generate email reply draft
 */
router.post('/draft-reply', async (req: Request, res: Response) => {
  try {
    const { originalEmail, tone, instructions, provider: providerName } = req.body;

    if (!originalEmail) {
      return res.status(400).json({ 
        success: false, 
        error: 'Original email is required' 
      });
    }

    const toneInstruction = tone ? `Use a ${tone} tone.` : 'Use a professional tone.';
    const additionalInstructions = instructions ? `Additional instructions: ${instructions}` : '';

    const request: AIRequest = {
      prompt: `Draft a reply to the following email. ${toneInstruction} ${additionalInstructions}

Original Email:
${originalEmail}

Please write a clear, concise reply that addresses the main points of the original email.`,
    };

    const provider = getProvider(providerName);
    const response = await provider.chat(request);

    res.json({
      success: true,
      draftReply: response.content,
    });
  } catch (error: any) {
    console.error('Draft reply error:', error);
    res.status(500).json({ 
      success: false, 
      error: 'Failed to generate reply' 
    });
  }
});

/**
 * POST /api/ai/extract-tasks
 * Extract action items from email
 */
router.post('/extract-tasks', async (req: Request, res: Response) => {
  try {
    const { content, provider: providerName } = req.body;

    if (!content) {
      return res.status(400).json({ 
        success: false, 
        error: 'Content is required' 
      });
    }

    const request: AIRequest = {
      prompt: `Analyze the following email and extract all action items, tasks, and deadlines. Format the response as a JSON array with objects containing "title", "priority" (high/normal/low), and "dueDate" (if mentioned).

Email:
${content}

Return ONLY valid JSON, no additional text.`,
    };

    const provider = getProvider(providerName);
    const response = await provider.chat(request);

    // Try to parse as JSON, fallback to text parsing
    let tasks = [];
    try {
      tasks = JSON.parse(response.content);
    } catch {
      // Parse as text if JSON fails
      const lines = response.content.split('\n').filter(line => line.trim());
      tasks = lines.map((line, index) => ({
        title: line.replace(/^[-•*\d.]\s*/, '').trim(),
        priority: 'normal',
        order: index,
      }));
    }

    res.json({
      success: true,
      tasks,
    });
  } catch (error: any) {
    console.error('Extract tasks error:', error);
    res.status(500).json({ 
      success: false, 
      error: 'Failed to extract tasks' 
    });
  }
});

/**
 * POST /api/ai/plan-day
 * Generate a day plan based on calendar and tasks
 */
router.post('/plan-day', async (req: Request, res: Response) => {
  try {
    const { events, tasks, preferences, provider: providerName } = req.body;

    const request: AIRequest = {
      prompt: `Create an optimized day plan based on the following schedule and tasks.

Calendar Events:
${JSON.stringify(events || [], null, 2)}

Pending Tasks:
${JSON.stringify(tasks || [], null, 2)}

User Preferences:
${preferences ? JSON.stringify(preferences) : 'Standard work day (9 AM - 5 PM)'}

Please create a structured day plan that:
1. Respects existing calendar commitments
2. Allocates time for high-priority tasks
3. Includes breaks and buffer time
4. Suggests optimal times for focused work
5. Provides practical recommendations

Format the plan in a clear, easy-to-follow structure.`,
    };

    const provider = getProvider(providerName);
    const response = await provider.chat(request);

    res.json({
      success: true,
      plan: response.content,
    });
  } catch (error: any) {
    console.error('Plan day error:', error);
    res.status(500).json({ 
      success: false, 
      error: 'Failed to generate day plan' 
    });
  }
});

/**
 * GET /api/ai/models
 * List available AI models
 */
router.get('/models', async (req: Request, res: Response) => {
  try {
    const models = await githubModelsProvider.listModels();
    res.json({
      success: true,
      provider: 'github',
      models,
    });
  } catch (error: any) {
    res.status(500).json({ 
      success: false, 
      error: 'Failed to list models' 
    });
  }
});

export default router;
