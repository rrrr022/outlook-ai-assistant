import { create } from 'zustand';
import { persist } from 'zustand/middleware';

export type AIProvider = 'github' | 'openai' | 'anthropic' | 'azure' | 'xai';

export interface ProviderConfig {
  id: AIProvider;
  name: string;
  description: string;
  icon: string;
  baseUrl: string;
  models: { id: string; name: string }[];
  placeholder: string;
  helpUrl: string;
}

export const PROVIDERS: ProviderConfig[] = [
  {
    id: 'github',
    name: 'GitHub Models',
    description: 'Free tier with GPT-4o, Claude, Llama, Mistral & more',
    icon: '🐙',
    baseUrl: 'https://models.inference.ai.azure.com',
    models: [
      // === OpenAI Models ===
      { id: 'gpt-4o', name: '🧠 GPT-4o (Recommended)' },
      { id: 'gpt-4o-mini', name: '🧠 GPT-4o Mini (Fast & Cheap)' },
      { id: 'o1', name: '🧠 o1 (Advanced Reasoning)' },
      { id: 'o1-mini', name: '🧠 o1-mini (Fast Reasoning)' },
      { id: 'o1-preview', name: '🧠 o1-preview (Preview)' },
      { id: 'o3-mini', name: '🧠 o3-mini (Latest Reasoning)' },
      
      // === Anthropic Claude Models ===
      { id: 'claude-3-5-sonnet', name: '🔮 Claude 3.5 Sonnet (Best Quality)' },
      { id: 'claude-3-5-sonnet-v2', name: '🔮 Claude 3.5 Sonnet v2' },
      { id: 'claude-3-opus', name: '🔮 Claude 3 Opus (Most Capable)' },
      { id: 'claude-3-haiku', name: '🔮 Claude 3 Haiku (Fast)' },
      
      // === Meta Llama Models ===
      { id: 'Meta-Llama-3.1-405B-Instruct', name: '🦙 Llama 3.1 405B (Largest)' },
      { id: 'Meta-Llama-3.1-70B-Instruct', name: '🦙 Llama 3.1 70B' },
      { id: 'Meta-Llama-3.1-8B-Instruct', name: '🦙 Llama 3.1 8B (Fast)' },
      { id: 'Meta-Llama-3-70B-Instruct', name: '🦙 Llama 3 70B' },
      { id: 'Meta-Llama-3-8B-Instruct', name: '🦙 Llama 3 8B' },
      { id: 'Llama-3.2-90B-Vision-Instruct', name: '🦙 Llama 3.2 90B Vision' },
      { id: 'Llama-3.2-11B-Vision-Instruct', name: '🦙 Llama 3.2 11B Vision' },
      { id: 'Llama-3.3-70B-Instruct', name: '🦙 Llama 3.3 70B (Latest)' },
      
      // === Mistral Models ===
      { id: 'Mistral-Large-2411', name: '🌀 Mistral Large (Latest)' },
      { id: 'Mistral-Large', name: '🌀 Mistral Large' },
      { id: 'Mistral-Small', name: '🌀 Mistral Small (Fast)' },
      { id: 'Mistral-Nemo', name: '🌀 Mistral Nemo' },
      { id: 'Ministral-3B', name: '🌀 Ministral 3B (Tiny)' },
      { id: 'Codestral-2501', name: '🌀 Codestral (Code Expert)' },
      
      // === Microsoft Phi Models ===
      { id: 'Phi-4', name: '🔷 Phi-4 (Latest Microsoft)' },
      { id: 'Phi-3.5-MoE-instruct', name: '🔷 Phi-3.5 MoE' },
      { id: 'Phi-3.5-mini-instruct', name: '🔷 Phi-3.5 Mini' },
      { id: 'Phi-3.5-vision-instruct', name: '🔷 Phi-3.5 Vision' },
      { id: 'Phi-3-medium-128k-instruct', name: '🔷 Phi-3 Medium 128K' },
      { id: 'Phi-3-mini-128k-instruct', name: '🔷 Phi-3 Mini 128K' },
      
      // === Cohere Models ===
      { id: 'Cohere-command-r-plus-08-2024', name: '💎 Command R+ (Latest)' },
      { id: 'Cohere-command-r-08-2024', name: '💎 Command R' },
      { id: 'Cohere-command-r-plus', name: '💎 Command R+ (Aug)' },
      { id: 'Cohere-command-r', name: '💎 Command R' },
      
      // === AI21 Jamba Models ===
      { id: 'AI21-Jamba-1.5-Large', name: '🔶 Jamba 1.5 Large' },
      { id: 'AI21-Jamba-1.5-Mini', name: '🔶 Jamba 1.5 Mini' },
      
      // === DeepSeek Models ===
      { id: 'DeepSeek-R1', name: '🌊 DeepSeek R1 (Reasoning)' },
      { id: 'DeepSeek-V3', name: '🌊 DeepSeek V3' },
    ],
    placeholder: 'ghp_xxxxxxxxxxxxxxxxxxxx',
    helpUrl: 'https://github.com/settings/tokens',
  },
  {
    id: 'openai',
    name: 'OpenAI',
    description: 'Direct access to GPT-4, GPT-3.5',
    icon: '🧠',
    baseUrl: 'https://api.openai.com/v1',
    models: [
      { id: 'gpt-4o', name: 'GPT-4o (Latest)' },
      { id: 'gpt-4-turbo', name: 'GPT-4 Turbo' },
      { id: 'gpt-4', name: 'GPT-4' },
      { id: 'gpt-3.5-turbo', name: 'GPT-3.5 Turbo (Cheap)' },
    ],
    placeholder: 'sk-xxxxxxxxxxxxxxxxxxxx',
    helpUrl: 'https://platform.openai.com/api-keys',
  },
  {
    id: 'anthropic',
    name: 'Anthropic',
    description: 'Claude 3.5 Sonnet, Claude 3 Opus',
    icon: '🔮',
    baseUrl: 'https://api.anthropic.com/v1',
    models: [
      { id: 'claude-3-5-sonnet-latest', name: 'Claude 3.5 Sonnet (Best)' },
      { id: 'claude-3-opus-latest', name: 'Claude 3 Opus' },
      { id: 'claude-3-haiku-20240307', name: 'Claude 3 Haiku (Fast)' },
    ],
    placeholder: 'sk-ant-xxxxxxxxxxxxxxxxxxxx',
    helpUrl: 'https://console.anthropic.com/settings/keys',
  },
  {
    id: 'xai',
    name: 'xAI (Grok)',
    description: 'Grok - witty AI with real-time knowledge',
    icon: '🚀',
    baseUrl: 'https://api.x.ai/v1',
    models: [
      { id: 'grok-beta', name: 'Grok Beta' },
      { id: 'grok-2-1212', name: 'Grok 2' },
    ],
    placeholder: 'xai-xxxxxxxxxxxxxxxxxxxx',
    helpUrl: 'https://console.x.ai',
  },
  {
    id: 'azure',
    name: 'Azure OpenAI',
    description: 'Enterprise Azure deployment',
    icon: '☁️',
    baseUrl: '', // User provides their endpoint
    models: [
      { id: 'gpt-4o', name: 'GPT-4o' },
      { id: 'gpt-4', name: 'GPT-4' },
      { id: 'gpt-35-turbo', name: 'GPT-3.5 Turbo' },
    ],
    placeholder: 'your-azure-api-key',
    helpUrl: 'https://portal.azure.com',
  },
];

interface APIKeyState {
  // Selected provider
  selectedProvider: AIProvider;
  selectedModel: string;
  
  // API Keys (stored locally, never sent to our server unless needed for proxy)
  apiKeys: Record<AIProvider, string>;
  
  // Azure-specific settings
  azureEndpoint: string;
  azureDeploymentName: string;
  
  // Use server's key (fallback to hosted mode)
  useServerKey: boolean;
  
  // Connection status
  connectionStatus: 'untested' | 'testing' | 'connected' | 'error';
  lastError: string | null;
  
  // Actions
  setSelectedProvider: (provider: AIProvider) => void;
  setSelectedModel: (model: string) => void;
  setApiKey: (provider: AIProvider, key: string) => void;
  setAzureEndpoint: (endpoint: string) => void;
  setAzureDeploymentName: (name: string) => void;
  setUseServerKey: (use: boolean) => void;
  setConnectionStatus: (status: APIKeyState['connectionStatus']) => void;
  setLastError: (error: string | null) => void;
  clearApiKey: (provider: AIProvider) => void;
  clearAllKeys: () => void;
  
  // Getters
  getCurrentApiKey: () => string;
  getCurrentProvider: () => ProviderConfig;
  hasApiKey: (provider?: AIProvider) => boolean;
}

export const useAPIKeyStore = create<APIKeyState>()(
  persist(
    (set, get) => ({
      // Initial state
      selectedProvider: 'github',
      selectedModel: 'gpt-4o',
      apiKeys: {
        github: '',
        openai: '',
        anthropic: '',
        azure: '',
        xai: '',
      },
      azureEndpoint: '',
      azureDeploymentName: '',
      useServerKey: true, // Default to using server's key
      connectionStatus: 'untested',
      lastError: null,

      // Actions
      setSelectedProvider: (provider) => {
        const providerConfig = PROVIDERS.find(p => p.id === provider);
        set({ 
          selectedProvider: provider,
          selectedModel: providerConfig?.models[0]?.id || 'gpt-4o',
          connectionStatus: 'untested',
          lastError: null,
        });
      },
      
      setSelectedModel: (model) => set({ selectedModel: model }),
      
      setApiKey: (provider, key) => set((state) => ({
        apiKeys: { ...state.apiKeys, [provider]: key },
        useServerKey: false, // When user sets a key, switch to BYOK mode
        connectionStatus: 'untested',
        lastError: null,
      })),
      
      setAzureEndpoint: (endpoint) => set({ azureEndpoint: endpoint }),
      
      setAzureDeploymentName: (name) => set({ azureDeploymentName: name }),
      
      setUseServerKey: (use) => set({ 
        useServerKey: use,
        connectionStatus: 'untested',
      }),
      
      setConnectionStatus: (status) => set({ connectionStatus: status }),
      
      setLastError: (error) => set({ lastError: error }),
      
      clearApiKey: (provider) => set((state) => ({
        apiKeys: { ...state.apiKeys, [provider]: '' },
      })),
      
      clearAllKeys: () => set({
        apiKeys: { github: '', openai: '', anthropic: '', azure: '', xai: '' },
        useServerKey: true,
        connectionStatus: 'untested',
      }),

      // Getters
      getCurrentApiKey: () => {
        const state = get();
        return state.apiKeys[state.selectedProvider] || '';
      },
      
      getCurrentProvider: () => {
        const state = get();
        return PROVIDERS.find(p => p.id === state.selectedProvider) || PROVIDERS[0];
      },
      
      hasApiKey: (provider) => {
        const state = get();
        const targetProvider = provider || state.selectedProvider;
        return !!state.apiKeys[targetProvider]?.trim();
      },
    }),
    {
      name: 'outlook-ai-keys', // localStorage key
      partialize: (state) => ({
        // Only persist these fields
        selectedProvider: state.selectedProvider,
        selectedModel: state.selectedModel,
        apiKeys: state.apiKeys,
        azureEndpoint: state.azureEndpoint,
        azureDeploymentName: state.azureDeploymentName,
        useServerKey: state.useServerKey,
      }),
    }
  )
);

// Helper to mask API keys for display
export function maskApiKey(key: string): string {
  if (!key || key.length < 10) return '••••••••';
  return key.substring(0, 8) + '••••••••' + key.substring(key.length - 4);
}
