import React, { useState } from 'react';
import { useAPIKeyStore, PROVIDERS, AIProvider } from '../store/apiKeyStore';
import { aiService } from '../services/aiService';

interface OnboardingProps {
  onComplete: () => void;
}

const OnboardingPanel: React.FC<OnboardingProps> = ({ onComplete }) => {
  const [currentStep, setCurrentStep] = useState(0);
  const [selectedProvider, setSelectedProvider] = useState<AIProvider | null>(null);
  const [apiKey, setApiKey] = useState('');
  const [isTestingKey, setIsTestingKey] = useState(false);
  const [testResult, setTestResult] = useState<'success' | 'error' | null>(null);
  const [showKey, setShowKey] = useState(false);

  const { setApiKey: saveApiKey, setSelectedProvider: saveProvider, setUseServerKey } = useAPIKeyStore();

  // Provider details with setup instructions
  const providerDetails: Record<string, {
    icon: string;
    name: string;
    tagline: string;
    features: string[];
    steps: string[];
    link: string;
    linkText: string;
    placeholder: string;
    recommended?: boolean;
    free?: boolean;
  }> = {
    github: {
      icon: '🐙',
      name: 'GitHub Models',
      tagline: 'FREE access to GPT-4o, Claude, Llama & more!',
      features: [
        '✅ Completely FREE tier available',
        '✅ Access to multiple top AI models',
        '✅ Simple GitHub account login',
        '✅ Great for getting started',
      ],
      steps: [
        'Go to github.com and sign in (or create free account)',
        'Click your profile picture → Settings',
        'Scroll down to "Developer settings" (bottom left)',
        'Click "Personal access tokens" → "Tokens (classic)"',
        'Click "Generate new token (classic)"',
        'Give it a name like "Outlook AI"',
        'Select expiration (90 days recommended)',
        'No special scopes needed - just click "Generate token"',
        'Copy the token starting with "ghp_" and paste below!',
      ],
      link: 'https://github.com/settings/tokens',
      linkText: 'Open GitHub Token Settings →',
      placeholder: 'ghp_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx',
      recommended: true,
      free: true,
    },
    openai: {
      icon: '🧠',
      name: 'OpenAI (ChatGPT)',
      tagline: 'The original ChatGPT - GPT-4o & GPT-4 Turbo',
      features: [
        '✅ Most popular AI models',
        '✅ GPT-4o with vision capabilities',
        '✅ Excellent for writing & analysis',
        '💰 Pay-as-you-go pricing',
      ],
      steps: [
        'Go to platform.openai.com',
        'Sign in or create an account',
        'Click your profile icon → "API keys"',
        'Click "+ Create new secret key"',
        'Name it "Outlook AI" and click "Create"',
        'Copy the key starting with "sk-" immediately!',
        '⚠️ You won\'t be able to see it again',
        'Add billing info if you haven\'t (Settings → Billing)',
      ],
      link: 'https://platform.openai.com/api-keys',
      linkText: 'Open OpenAI API Keys →',
      placeholder: 'sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx',
    },
    xai: {
      icon: '🚀',
      name: 'xAI (Grok)',
      tagline: 'Elon Musk\'s Grok - witty & up-to-date',
      features: [
        '✅ Real-time information access',
        '✅ Witty personality',
        '✅ Great for current events',
        '💰 Pay-as-you-go pricing',
      ],
      steps: [
        'Go to console.x.ai',
        'Sign in with your X (Twitter) account',
        'Navigate to API Keys section',
        'Click "Create API Key"',
        'Name it "Outlook AI"',
        'Copy the generated key',
        'Store it safely - you may not see it again!',
      ],
      link: 'https://console.x.ai',
      linkText: 'Open xAI Console →',
      placeholder: 'xai-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx',
    },
    anthropic: {
      icon: '🔮',
      name: 'Anthropic (Claude)',
      tagline: 'Claude 3.5 Sonnet - thoughtful & capable',
      features: [
        '✅ Excellent reasoning abilities',
        '✅ 200K context window',
        '✅ Great for long documents',
        '💰 Pay-as-you-go pricing',
      ],
      steps: [
        'Go to console.anthropic.com',
        'Sign up or sign in',
        'Go to Settings → API Keys',
        'Click "Create Key"',
        'Name it "Outlook AI"',
        'Copy the key starting with "sk-ant-"',
        'Add billing info in Settings → Plans',
      ],
      link: 'https://console.anthropic.com/settings/keys',
      linkText: 'Open Anthropic Console →',
      placeholder: 'sk-ant-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx',
    },
    azure: {
      icon: '☁️',
      name: 'Microsoft Azure OpenAI',
      tagline: 'Enterprise-grade with Microsoft integration',
      features: [
        '✅ Enterprise security & compliance',
        '✅ Microsoft ecosystem integration',
        '✅ Copilot-like capabilities',
        '💼 Best for business/enterprise',
      ],
      steps: [
        'Go to portal.azure.com',
        'Search for "Azure OpenAI" in the search bar',
        'Create an Azure OpenAI resource',
        'Wait for approval (may take 1-2 days)',
        'Once approved, go to your resource',
        'Click "Keys and Endpoint" in the left menu',
        'Copy Key 1 and your Endpoint URL',
        'Deploy a model in Azure OpenAI Studio',
      ],
      link: 'https://portal.azure.com/#create/Microsoft.CognitiveServicesOpenAI',
      linkText: 'Open Azure Portal →',
      placeholder: 'your-azure-api-key',
    },
  };

  const capabilities = [
    {
      icon: '📧',
      title: 'Email Superpowers',
      items: [
        'Summarize long email threads instantly',
        'Draft professional replies in seconds',
        'Extract action items automatically',
        'Analyze tone and sentiment',
        'Translate emails to any language',
      ],
    },
    {
      icon: '📅',
      title: 'Calendar Intelligence',
      items: [
        'Smart meeting scheduling suggestions',
        'Conflict detection and resolution',
        'Automatic meeting prep summaries',
        'Time zone management',
        'Meeting notes and follow-ups',
      ],
    },
    {
      icon: '📊',
      title: 'Document Creation',
      items: [
        'Generate Excel spreadsheets from data',
        'Create PowerPoint presentations',
        'Build charts and visualizations',
        'Write reports and summaries',
        'Format documents professionally',
      ],
    },
    {
      icon: '🤖',
      title: 'Automation Magic',
      items: [
        'Auto-categorize incoming emails',
        'Smart priority inbox sorting',
        'Automated follow-up reminders',
        'Custom workflow triggers',
        'Batch email processing',
      ],
    },
  ];

  const handleTestConnection = async () => {
    if (!selectedProvider || !apiKey) return;
    
    setIsTestingKey(true);
    setTestResult(null);
    
    // Save temporarily for testing
    saveProvider(selectedProvider as AIProvider);
    saveApiKey(selectedProvider as AIProvider, apiKey);
    setUseServerKey(false);
    
    const result = await aiService.testConnection();
    
    setTestResult(result.success ? 'success' : 'error');
    setIsTestingKey(false);
  };

  const handleComplete = () => {
    if (selectedProvider && apiKey) {
      saveProvider(selectedProvider as AIProvider);
      saveApiKey(selectedProvider as AIProvider, apiKey);
      setUseServerKey(false);
    }
    onComplete();
  };

  const handleSkip = () => {
    setUseServerKey(true);
    onComplete();
  };

  const renderWelcomeStep = () => (
    <div className="onboarding-step welcome-step">
      <div className="welcome-header">
        <div className="logo-container">
          <span className="logo-icon">🤖</span>
          <h1 className="logo-text">FreedomForged AI</h1>
        </div>
        <p className="welcome-tagline">for Outlook</p>
      </div>

      <div className="welcome-intro">
        <p className="intro-text">
          Welcome to the future of email and calendar management. 
          FreedomForged AI transforms your Outlook into an intelligent assistant 
          that understands, organizes, and acts on your behalf.
        </p>
      </div>

      <div className="capabilities-grid">
        {capabilities.map((cap, index) => (
          <div key={index} className="capability-card">
            <div className="capability-icon">{cap.icon}</div>
            <h3 className="capability-title">{cap.title}</h3>
            <ul className="capability-list">
              {cap.items.slice(0, 3).map((item, i) => (
                <li key={i}>{item}</li>
              ))}
            </ul>
          </div>
        ))}
      </div>

      <div className="welcome-cta">
        <button className="cta-button primary" onClick={() => setCurrentStep(1)}>
          <span>Get Started</span>
          <span className="cta-arrow">→</span>
        </button>
        <button className="cta-button secondary" onClick={handleSkip}>
          Skip for now
        </button>
      </div>
    </div>
  );

  const renderProviderStep = () => (
    <div className="onboarding-step provider-step">
      <div className="step-header">
        <h2>Choose Your AI Provider</h2>
        <p>Select the AI service you'd like to use. You can change this anytime in Settings.</p>
      </div>

      <div className="provider-selection">
        {Object.entries(providerDetails).map(([id, provider]) => (
          <div
            key={id}
            className={`provider-option ${selectedProvider === id ? 'selected' : ''} ${provider.recommended ? 'recommended' : ''}`}
            onClick={() => setSelectedProvider(id as AIProvider)}
          >
            {provider.recommended && <span className="recommended-badge">⭐ RECOMMENDED</span>}
            {provider.free && <span className="free-badge">FREE</span>}
            <div className="provider-header">
              <span className="provider-icon">{provider.icon}</span>
              <div className="provider-info">
                <h3 className="provider-name">{provider.name}</h3>
                <p className="provider-tagline">{provider.tagline}</p>
              </div>
            </div>
            <ul className="provider-features">
              {provider.features.map((feature, i) => (
                <li key={i}>{feature}</li>
              ))}
            </ul>
          </div>
        ))}
      </div>

      <div className="step-actions">
        <button className="cta-button secondary" onClick={() => setCurrentStep(0)}>
          ← Back
        </button>
        <button 
          className="cta-button primary" 
          onClick={() => setCurrentStep(2)}
          disabled={!selectedProvider}
        >
          Continue →
        </button>
      </div>
    </div>
  );

  const renderSetupStep = () => {
    const provider = selectedProvider ? providerDetails[selectedProvider] : null;
    if (!provider) return null;

    return (
      <div className="onboarding-step setup-step">
        <div className="step-header">
          <div className="provider-badge">
            <span>{provider.icon}</span>
            <span>{provider.name}</span>
          </div>
          <h2>Get Your API Key</h2>
          <p>Follow these steps to get your personal API key:</p>
        </div>

        <div className="setup-instructions">
          <ol className="instruction-list">
            {provider.steps.map((step, i) => (
              <li key={i} className="instruction-item">
                <span className="step-number">{i + 1}</span>
                <span className="step-text">{step}</span>
              </li>
            ))}
          </ol>

          <a 
            href={provider.link} 
            target="_blank" 
            rel="noopener noreferrer"
            className="external-link-btn"
          >
            {provider.linkText}
            <span className="external-icon">↗</span>
          </a>
        </div>

        <div className="api-key-input-section">
          <label className="input-label">Paste your API key here:</label>
          <div className="key-input-wrapper">
            <input
              type={showKey ? 'text' : 'password'}
              value={apiKey}
              onChange={(e) => {
                setApiKey(e.target.value);
                setTestResult(null);
              }}
              placeholder={provider.placeholder}
              className="api-key-input"
            />
            <button 
              className="toggle-visibility-btn"
              onClick={() => setShowKey(!showKey)}
            >
              {showKey ? '🙈' : '👁️'}
            </button>
          </div>

          {apiKey && (
            <div className="test-section">
              <button 
                className="test-btn"
                onClick={handleTestConnection}
                disabled={isTestingKey || !apiKey}
              >
                {isTestingKey ? '◐ Testing...' : '🔌 Test Connection'}
              </button>
              
              {testResult === 'success' && (
                <div className="test-result success">
                  ✅ Connection successful! Your API key is working.
                </div>
              )}
              {testResult === 'error' && (
                <div className="test-result error">
                  ❌ Connection failed. Please check your API key and try again.
                </div>
              )}
            </div>
          )}
        </div>

        <div className="security-info">
          <span className="security-icon">🔒</span>
          <p>Your API key is stored securely in your browser's local storage and is never sent to our servers.</p>
        </div>

        <div className="step-actions">
          <button className="cta-button secondary" onClick={() => setCurrentStep(1)}>
            ← Back
          </button>
          <button 
            className="cta-button primary"
            onClick={() => setCurrentStep(3)}
            disabled={!apiKey}
          >
            Continue →
          </button>
        </div>
      </div>
    );
  };

  const renderCompleteStep = () => (
    <div className="onboarding-step complete-step">
      <div className="completion-animation">
        <div className="success-ring">
          <span className="success-icon">✓</span>
        </div>
      </div>

      <h2 className="completion-title">You're All Set!</h2>
      <p className="completion-subtitle">Outlook AI is now connected and ready to supercharge your Outlook.</p>

      <div className="quick-tips">
        <h3>🚀 Quick Tips to Get Started</h3>
        <div className="tips-grid">
          <div className="tip-card">
            <span className="tip-icon">💬</span>
            <h4>Chat Tab</h4>
            <p>Ask questions about your emails or request help with any task</p>
          </div>
          <div className="tip-card">
            <span className="tip-icon">📨</span>
            <h4>Mail Tab</h4>
            <p>Click "Summarize" or "Draft Reply" on any email</p>
          </div>
          <div className="tip-card">
            <span className="tip-icon">📅</span>
            <h4>Calendar Tab</h4>
            <p>Get AI-suggested meeting times and prep notes</p>
          </div>
          <div className="tip-card">
            <span className="tip-icon">⚙️</span>
            <h4>Config Tab</h4>
            <p>Change AI providers or update your API key anytime</p>
          </div>
        </div>
      </div>

      <div className="example-prompts">
        <h3>💡 Try These First</h3>
        <div className="prompt-examples">
          <code>"Summarize my unread emails"</code>
          <code>"Draft a polite follow-up email"</code>
          <code>"What meetings do I have today?"</code>
          <code>"Help me write a professional response"</code>
        </div>
      </div>

      <button className="cta-button primary launch-btn" onClick={handleComplete}>
        <span className="launch-icon">🤖</span>
        Launch Outlook AI
      </button>
    </div>
  );

  const steps = [
    { id: 'welcome', label: 'Welcome', render: renderWelcomeStep },
    { id: 'provider', label: 'Choose AI', render: renderProviderStep },
    { id: 'setup', label: 'Get Key', render: renderSetupStep },
    { id: 'complete', label: 'Ready!', render: renderCompleteStep },
  ];

  return (
    <div className="onboarding-container">
      {/* Progress indicator */}
      <div className="progress-bar">
        {steps.map((step, index) => (
          <div 
            key={step.id}
            className={`progress-step ${index <= currentStep ? 'active' : ''} ${index < currentStep ? 'completed' : ''}`}
          >
            <div className="progress-dot">
              {index < currentStep ? '✓' : index + 1}
            </div>
            <span className="progress-label">{step.label}</span>
          </div>
        ))}
        <div 
          className="progress-line" 
          style={{ width: `${(currentStep / (steps.length - 1)) * 100}%` }}
        />
      </div>

      {/* Current step content */}
      <div className="onboarding-content">
        {steps[currentStep].render()}
      </div>
    </div>
  );
};

export default OnboardingPanel;
