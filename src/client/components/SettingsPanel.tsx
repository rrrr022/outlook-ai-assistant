import React, { useState, useMemo } from 'react';
import { Switch, Spinner } from './ui/NativeComponents';
import { useAppStore } from '../store/appStore';
import { useAPIKeyStore, PROVIDERS, maskApiKey, AIProvider } from '../store/apiKeyStore';
import { AutomationRule } from '@shared/types';
import { aiService } from '../services/aiService';
import { v4 as uuidv4 } from 'uuid';

const SettingsPanel: React.FC = () => {
  const { 
    settings, 
    updateSettings, 
    automationRules, 
    addAutomationRule, 
    updateAutomationRule,
    deleteAutomationRule 
  } = useAppStore();

  const {
    selectedProvider,
    selectedModel,
    apiKeys,
    useServerKey,
    connectionStatus,
    lastError,
    azureEndpoint,
    azureDeploymentName,
    setSelectedProvider,
    setSelectedModel,
    setApiKey,
    setAzureEndpoint,
    setAzureDeploymentName,
    setUseServerKey,
    clearApiKey,
    getCurrentProvider,
  } = useAPIKeyStore();
  
  const [showAddRule, setShowAddRule] = useState(false);
  const [newRuleName, setNewRuleName] = useState('');
  const [newRuleDescription, setNewRuleDescription] = useState('');
  const [showApiKey, setShowApiKey] = useState(false);
  const [tempApiKey, setTempApiKey] = useState('');
  const [isTesting, setIsTesting] = useState(false);
  const [modelFilter, setModelFilter] = useState('');
  const [showModelSelector, setShowModelSelector] = useState(false);

  const currentProvider = getCurrentProvider();
  const currentApiKey = apiKeys[selectedProvider] || '';

  // Filter models based on search
  const filteredModels = useMemo(() => {
    if (!modelFilter.trim()) return currentProvider.models;
    const search = modelFilter.toLowerCase();
    return currentProvider.models.filter(m => 
      m.name.toLowerCase().includes(search) || 
      m.id.toLowerCase().includes(search)
    );
  }, [currentProvider.models, modelFilter]);

  // Get current model name
  const currentModelName = currentProvider.models.find(m => m.id === selectedModel)?.name || selectedModel;

  const handleAddRule = () => {
    if (!newRuleName.trim()) return;

    const newRule: AutomationRule = {
      id: uuidv4(),
      name: newRuleName.trim(),
      description: newRuleDescription.trim(),
      isEnabled: true,
      priority: automationRules.length + 1,
      conditions: [],
      conditionLogic: 'and',
      actions: [],
      createdAt: new Date(),
      updatedAt: new Date(),
      triggerCount: 0,
    };

    addAutomationRule(newRule);
    setNewRuleName('');
    setNewRuleDescription('');
    setShowAddRule(false);
  };

  const handleSaveApiKey = () => {
    if (tempApiKey.trim()) {
      setApiKey(selectedProvider, tempApiKey.trim());
      setTempApiKey('');
    }
  };

  const handleTestConnection = async () => {
    setIsTesting(true);
    await aiService.testConnection();
    setIsTesting(false);
  };

  const presetRules = [
    {
      name: 'Auto-reply to VIP contacts',
      description: 'Automatically acknowledge emails from important contacts',
      icon: '⭐',
    },
    {
      name: 'Create tasks from flagged emails',
      description: 'Automatically create tasks when you flag an email',
      icon: '🚩',
    },
    {
      name: 'Schedule follow-ups',
      description: 'Remind me to follow up on unanswered emails after 3 days',
      icon: '⏰',
    },
    {
      name: 'Meeting prep reminder',
      description: 'Send summary of relevant emails before scheduled meetings',
      icon: '📋',
    },
  ];

  return (
    <div>
      <h2 className="section-title">⚙️ Settings</h2>

      {/* API Keys Section - BYOK */}
      <div className="card api-keys-card">
        <h3 className="section-header--lg-margin">
          🔐 API Keys (BYOK)
        </h3>
        
        <p className="helper-text mb-12">
          <strong>Bring Your Own Key</strong> - Your keys are stored locally in your browser and never sent to our servers.
        </p>

        {/* Mode Toggle */}
        <div className="mode-toggle mb-16">
          <div 
            className={`mode-option ${useServerKey ? 'active' : ''}`}
            onClick={() => setUseServerKey(true)}
          >
            <span className="mode-icon">☁️</span>
            <div className="mode-info">
              <div className="mode-title">Hosted Mode</div>
              <div className="mode-desc">Use app's shared API</div>
            </div>
          </div>
          <div 
            className={`mode-option ${!useServerKey ? 'active' : ''}`}
            onClick={() => setUseServerKey(false)}
          >
            <span className="mode-icon">🔑</span>
            <div className="mode-info">
              <div className="mode-title">Your Own Key</div>
              <div className="mode-desc">Direct to AI provider</div>
            </div>
          </div>
        </div>

        {/* Provider Selection */}
        <div className="form-row">
          <label className="form-label">AI Provider</label>
          <div className="provider-grid">
            {PROVIDERS.map((provider) => (
              <div
                key={provider.id}
                className={`provider-card ${selectedProvider === provider.id ? 'selected' : ''} ${useServerKey ? 'disabled' : ''}`}
                onClick={() => !useServerKey && setSelectedProvider(provider.id)}
              >
                <span className="provider-icon">{provider.icon}</span>
                <div className="provider-info">
                  <div className="provider-name">{provider.name}</div>
                  <div className="provider-desc">{provider.description}</div>
                </div>
                {apiKeys[provider.id] && (
                  <span className="key-badge">✓ Key Set</span>
                )}
              </div>
            ))}
          </div>
        </div>

        {/* Model Selection */}
        {!useServerKey && (
          <div className="form-row mt-12">
            <label className="form-label">AI Model</label>
            <div className="model-selector-container">
              <button
                type="button"
                className="model-selector-btn"
                onClick={() => setShowModelSelector(!showModelSelector)}
              >
                <span className="model-current">{currentModelName}</span>
                <span className="model-dropdown-icon">{showModelSelector ? '▲' : '▼'}</span>
              </button>
              
              {showModelSelector && (
                <div className="model-dropdown">
                  <input
                    type="text"
                    placeholder="🔍 Search models..."
                    value={modelFilter}
                    onChange={(e) => setModelFilter(e.target.value)}
                    className="model-search-input"
                    autoFocus
                  />
                  <div className="model-list">
                    {filteredModels.length === 0 ? (
                      <div className="model-empty">No models found</div>
                    ) : (
                      filteredModels.map((model) => (
                        <div
                          key={model.id}
                          className={`model-item ${selectedModel === model.id ? 'active' : ''}`}
                          onClick={() => {
                            setSelectedModel(model.id);
                            setShowModelSelector(false);
                            setModelFilter('');
                          }}
                        >
                          <span className="model-name">{model.name}</span>
                          {selectedModel === model.id && <span className="model-check">✓</span>}
                        </div>
                      ))
                    )}
                  </div>
                  <div className="model-count">
                    {currentProvider.models.length} models available
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

        {/* Azure-specific fields */}
        {!useServerKey && selectedProvider === 'azure' && (
          <>
            <div className="form-row mt-12">
              <label className="form-label">Azure Endpoint</label>
              <input
                type="text"
                value={azureEndpoint}
                onChange={(e) => setAzureEndpoint(e.target.value)}
                placeholder="https://your-resource.openai.azure.com"
                className="chat-input input-rounded"
              />
            </div>
            <div className="form-row mt-8">
              <label className="form-label">Deployment Name</label>
              <input
                type="text"
                value={azureDeploymentName}
                onChange={(e) => setAzureDeploymentName(e.target.value)}
                placeholder="gpt-4o-deployment"
                className="chat-input input-rounded"
              />
            </div>
          </>
        )}

        {/* API Key Input */}
        {!useServerKey && (
          <div className="form-row mt-12">
            <label className="form-label">
              API Key
              <a 
                href={currentProvider.helpUrl} 
                target="_blank" 
                rel="noopener noreferrer"
                className="help-link"
              >
                Get key →
              </a>
            </label>
            
            {currentApiKey ? (
              <div className="key-display">
                <span className="masked-key">
                  {showApiKey ? currentApiKey : maskApiKey(currentApiKey)}
                </span>
                <div className="key-actions">
                  <button
                    className="icon-btn"
                    onClick={() => setShowApiKey(!showApiKey)}
                    title={showApiKey ? 'Hide' : 'Show'}
                  >
                    {showApiKey ? '🙈' : '👁️'}
                  </button>
                  <button
                    className="icon-btn danger"
                    onClick={() => clearApiKey(selectedProvider)}
                    title="Remove key"
                  >
                    🗑️
                  </button>
                </div>
              </div>
            ) : (
              <div className="key-input-group">
                <input
                  type="password"
                  value={tempApiKey}
                  onChange={(e) => setTempApiKey(e.target.value)}
                  placeholder={currentProvider.placeholder}
                  className="chat-input input-rounded"
                />
                <button
                  className="action-button primary"
                  onClick={handleSaveApiKey}
                  disabled={!tempApiKey.trim()}
                >
                  Save
                </button>
              </div>
            )}
          </div>
        )}

        {/* Connection Status & Test */}
        {!useServerKey && currentApiKey && (
          <div className="connection-status mt-12">
            <div className={`status-indicator status-${connectionStatus}`}>
              {connectionStatus === 'untested' && '○ Not tested'}
              {connectionStatus === 'testing' && '◐ Testing...'}
              {connectionStatus === 'connected' && '● Connected'}
              {connectionStatus === 'error' && '○ Error'}
            </div>
            <button
              className="action-button secondary action-button--sm"
              onClick={handleTestConnection}
              disabled={isTesting}
            >
              {isTesting ? <Spinner size="small" /> : 'Test Connection'}
            </button>
          </div>
        )}

        {lastError && (
          <div className="error-message mt-8">
            ⚠️ {lastError}
          </div>
        )}

        {/* Security Notice */}
        <div className="security-notice mt-12">
          <span className="security-icon">🛡️</span>
          <span>Keys stored in browser localStorage only. Never transmitted to our servers.</span>
        </div>
      </div>

      {/* Feature Toggles */}
      <div className="card mt-12">
        <h3 className="section-header--lg-margin">
          🔧 Features
        </h3>

        <div className="rule-item">
          <div className="rule-info">
            <div className="rule-name">Auto-Reply Suggestions</div>
            <div className="rule-description">Get AI-generated reply suggestions for emails</div>
          </div>
          <Switch
            checked={settings.autoReplyEnabled}
            onChange={(e: React.ChangeEvent<HTMLInputElement>) => updateSettings({ autoReplyEnabled: e.target.checked })}
          />
        </div>

        <div className="rule-item">
          <div className="rule-info">
            <div className="rule-name">Calendar Sync</div>
            <div className="rule-description">Sync calendar events for context-aware assistance</div>
          </div>
          <Switch
            checked={settings.calendarSyncEnabled}
            onChange={(e: React.ChangeEvent<HTMLInputElement>) => updateSettings({ calendarSyncEnabled: e.target.checked })}
          />
        </div>

        <div className="rule-item">
          <div className="rule-info">
            <div className="rule-name">Task Extraction</div>
            <div className="rule-description">Automatically detect action items in emails</div>
          </div>
          <Switch
            checked={settings.taskExtractionEnabled}
            onChange={(e: React.ChangeEvent<HTMLInputElement>) => updateSettings({ taskExtractionEnabled: e.target.checked })}
          />
        </div>
      </div>

      {/* Approval Settings */}
      <div className="card mt-12">
        <h3 className="section-header--lg-margin">
          🛡️ Approval Controls
        </h3>
        <p className="helper-text mb-12">
          Require your approval before the AI performs these actions
        </p>

        <div className="rule-item">
          <div className="rule-info">
            <div className="rule-name">Email Sending</div>
            <div className="rule-description">Ask before sending or replying to emails</div>
          </div>
          <Switch
            checked={settings.requireApprovalForEmails}
            onChange={(e: React.ChangeEvent<HTMLInputElement>) => updateSettings({ requireApprovalForEmails: e.target.checked })}
          />
        </div>

        <div className="rule-item">
          <div className="rule-info">
            <div className="rule-name">Meeting Creation</div>
            <div className="rule-description">Ask before scheduling meetings or events</div>
          </div>
          <Switch
            checked={settings.requireApprovalForMeetings}
            onChange={(e: React.ChangeEvent<HTMLInputElement>) => updateSettings({ requireApprovalForMeetings: e.target.checked })}
          />
        </div>

        <div className="rule-item">
          <div className="rule-info">
            <div className="rule-name">Automation Rules</div>
            <div className="rule-description">Ask before automation rules execute actions</div>
          </div>
          <Switch
            checked={settings.requireApprovalForAutomations}
            onChange={(e: React.ChangeEvent<HTMLInputElement>) => updateSettings({ requireApprovalForAutomations: e.target.checked })}
          />
        </div>
      </div>

      {/* Automation Rules */}
      <div className="card mt-12">
        <div className="calendar-header">
          <h3 className="section-header--no-margin">
            🤖 Automation Rules
          </h3>
          <button 
            className="action-button primary action-button--sm"
            onClick={() => setShowAddRule(true)}
          >
            + Add Rule
          </button>
        </div>

        {/* Add Rule Form */}
        {showAddRule && (
          <div className="add-rule-form">
            <input
              type="text"
              placeholder="Rule name"
              value={newRuleName}
              onChange={(e) => setNewRuleName(e.target.value)}
              className="chat-input input-rounded mb-8"
            />
            <input
              type="text"
              placeholder="Description (optional)"
              value={newRuleDescription}
              onChange={(e) => setNewRuleDescription(e.target.value)}
              className="chat-input input-rounded mb-8"
            />
            <div className="add-rule-buttons">
              <button className="action-button primary" onClick={handleAddRule}>
                Create
              </button>
              <button 
                className="action-button secondary" 
                onClick={() => setShowAddRule(false)}
              >
                Cancel
              </button>
            </div>
          </div>
        )}

        {/* Existing Rules */}
        {automationRules.length > 0 ? (
          automationRules.map((rule) => (
            <div key={rule.id} className="rule-item">
              <div className="rule-info">
                <div className="rule-name">{rule.name}</div>
                {rule.description && (
                  <div className="rule-description">{rule.description}</div>
                )}
                <div className="rule-description rule-triggered">
                  Triggered {rule.triggerCount} times
                </div>
              </div>
              <div className="rule-actions">
                <Switch
                  checked={rule.isEnabled}
                  onChange={(e: React.ChangeEvent<HTMLInputElement>) => 
                    updateAutomationRule(rule.id, { isEnabled: e.target.checked })
                  }
                />
                <button
                  onClick={() => deleteAutomationRule(rule.id)}
                  className="delete-btn delete-btn--sm"
                  aria-label={`Delete ${rule.name}`}
                >
                  🗑️
                </button>
              </div>
            </div>
          ))
        ) : (
          <div className="empty-state">
            <p className="empty-state-text">No automation rules configured</p>
          </div>
        )}

        {/* Preset Rules Suggestions */}
        <div className="mt-16">
          <div className="suggested-header">
            Suggested automations:
          </div>
          {presetRules.map((preset, index) => (
            <div 
              key={index}
              className="quick-action-btn preset-suggestion mb-8"
              onClick={() => {
                setNewRuleName(preset.name);
                setNewRuleDescription(preset.description);
                setShowAddRule(true);
              }}
            >
              <span className="preset-suggestion-icon">{preset.icon}</span>
              <div className="preset-suggestion-content">
                <div className="preset-suggestion-title">{preset.name}</div>
                <div className="preset-suggestion-desc">{preset.description}</div>
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* About */}
      <div className="card mt-12">
        <h3 className="section-header">
          ℹ️ About
        </h3>
        <p className="text-small text-muted">
          <strong>Outlook AI</strong> v1.5.0.1
          <br />
          Your intelligent email, calendar, and task management companion.
          <br /><br />
          Built by <strong>FreedomForged_AI</strong>
        </p>
      </div>
    </div>
  );
};

export default SettingsPanel;
