import React, { useState } from 'react';
import { Switch } from '@fluentui/react-components';
import { useAppStore } from '../store/appStore';
import { AutomationRule } from '@shared/types';
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
  
  const [showAddRule, setShowAddRule] = useState(false);
  const [newRuleName, setNewRuleName] = useState('');
  const [newRuleDescription, setNewRuleDescription] = useState('');

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

      {/* AI Provider Settings */}
      <div className="card">
        <h3 className="section-header--lg-margin">
          🤖 AI Configuration
        </h3>
        
        <div className="form-row">
          <label className="form-label">
            AI Provider
          </label>
          <select
            value={settings.aiProvider}
            onChange={(e) => updateSettings({ aiProvider: e.target.value as 'openai' | 'anthropic' })}
            className="select-input"
          >
            <option value="github">GitHub Models (GPT-4o, Claude) ⭐</option>
            <option value="openai">OpenAI (Direct API)</option>
            <option value="anthropic">Anthropic (Direct API)</option>
          </select>
        </div>

        <p className="helper-text">
          💡 GitHub Models uses your GitHub account - no extra API costs!
          <br />
          Get a token from: github.com/settings/tokens
        </p>
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
            onChange={(_, data) => updateSettings({ autoReplyEnabled: data.checked })}
          />
        </div>

        <div className="rule-item">
          <div className="rule-info">
            <div className="rule-name">Calendar Sync</div>
            <div className="rule-description">Sync calendar events for context-aware assistance</div>
          </div>
          <Switch
            checked={settings.calendarSyncEnabled}
            onChange={(_, data) => updateSettings({ calendarSyncEnabled: data.checked })}
          />
        </div>

        <div className="rule-item">
          <div className="rule-info">
            <div className="rule-name">Task Extraction</div>
            <div className="rule-description">Automatically detect action items in emails</div>
          </div>
          <Switch
            checked={settings.taskExtractionEnabled}
            onChange={(_, data) => updateSettings({ taskExtractionEnabled: data.checked })}
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
            onChange={(_, data) => updateSettings({ requireApprovalForEmails: data.checked })}
          />
        </div>

        <div className="rule-item">
          <div className="rule-info">
            <div className="rule-name">Meeting Creation</div>
            <div className="rule-description">Ask before scheduling meetings or events</div>
          </div>
          <Switch
            checked={settings.requireApprovalForMeetings}
            onChange={(_, data) => updateSettings({ requireApprovalForMeetings: data.checked })}
          />
        </div>

        <div className="rule-item">
          <div className="rule-info">
            <div className="rule-name">Automation Rules</div>
            <div className="rule-description">Ask before automation rules execute actions</div>
          </div>
          <Switch
            checked={settings.requireApprovalForAutomations}
            onChange={(_, data) => updateSettings({ requireApprovalForAutomations: data.checked })}
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
                  onChange={(_, data) => 
                    updateAutomationRule(rule.id, { isEnabled: data.checked })
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
          Outlook AI Assistant v1.0.0
          <br />
          Your intelligent email, calendar, and task management companion.
        </p>
      </div>
    </div>
  );
};

export default SettingsPanel;
