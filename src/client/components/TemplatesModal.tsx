import React, { useState } from 'react';
import { aiService } from '../services/aiService';
import { outlookService } from '../services/outlookService';

interface EmailTemplate {
  id: string;
  name: string;
  icon: string;
  description: string;
  prompt: string;
}

const templates: EmailTemplate[] = [
  {
    id: 'meeting-request',
    name: 'Meeting Request',
    icon: '📅',
    description: 'Request a meeting with someone',
    prompt: 'Draft a professional meeting request email',
  },
  {
    id: 'follow-up',
    name: 'Follow Up',
    icon: '🔄',
    description: 'Follow up on a previous conversation',
    prompt: 'Draft a polite follow-up email',
  },
  {
    id: 'thank-you',
    name: 'Thank You',
    icon: '🙏',
    description: 'Express gratitude',
    prompt: 'Draft a thank you email',
  },
  {
    id: 'introduction',
    name: 'Introduction',
    icon: '👋',
    description: 'Introduce yourself or someone else',
    prompt: 'Draft a professional introduction email',
  },
  {
    id: 'decline-politely',
    name: 'Polite Decline',
    icon: '🙅',
    description: 'Decline a request professionally',
    prompt: 'Draft a polite email declining a request while maintaining good relations',
  },
  {
    id: 'project-update',
    name: 'Project Update',
    icon: '📊',
    description: 'Provide a status update',
    prompt: 'Draft a project status update email',
  },
  {
    id: 'request-info',
    name: 'Request Info',
    icon: '❓',
    description: 'Ask for information or clarification',
    prompt: 'Draft an email requesting information',
  },
  {
    id: 'out-of-office',
    name: 'Out of Office',
    icon: '🏖️',
    description: 'Auto-reply for absence',
    prompt: 'Draft an out-of-office auto-reply message',
  },
];

interface TemplatesModalProps {
  isOpen: boolean;
  onClose: () => void;
  onSelectTemplate: (content: string) => void;
}

const TemplatesModal: React.FC<TemplatesModalProps> = ({ isOpen, onClose, onSelectTemplate }) => {
  const [selectedTemplate, setSelectedTemplate] = useState<EmailTemplate | null>(null);
  const [customContext, setCustomContext] = useState('');
  const [generatedContent, setGeneratedContent] = useState('');
  const [loading, setLoading] = useState(false);

  if (!isOpen) return null;

  const handleGenerate = async () => {
    if (!selectedTemplate) return;

    setLoading(true);
    try {
      const response = await aiService.chat({
        prompt: `${selectedTemplate.prompt}. ${customContext ? `Additional context: ${customContext}` : ''}. 
        
        Format the email professionally with:
        - A clear subject line suggestion
        - Appropriate greeting
        - Concise body
        - Professional closing`,
      });
      setGeneratedContent(response.content);
    } catch (error) {
      console.error('Error generating template:', error);
      setGeneratedContent('Error generating content. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  const handleUse = async () => {
    if (generatedContent) {
      onSelectTemplate(generatedContent);
      // Try to insert into compose window
      try {
        await outlookService.insertTextToCompose(generatedContent);
      } catch (e) {
        // User might not be in compose mode
      }
      onClose();
    }
  };

  const handleReset = () => {
    setSelectedTemplate(null);
    setCustomContext('');
    setGeneratedContent('');
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-content" onClick={(e) => e.stopPropagation()}>
        <div className="modal-header">
          <h3>📝 Email Templates</h3>
          <button className="delete-btn" onClick={onClose}>×</button>
        </div>

        {!selectedTemplate ? (
          <div className="template-grid">
            {templates.map((template) => (
              <button
                key={template.id}
                className="template-btn"
                onClick={() => setSelectedTemplate(template)}
              >
                <span className="text-xl">{template.icon}</span>
                <div>
                  <div className="text-bold text-small">{template.name}</div>
                  <div className="text-small text-muted">{template.description}</div>
                </div>
              </button>
            ))}
          </div>
        ) : !generatedContent ? (
          <div>
            <div className="flex items-center gap-8 mb-12">
              <button className="action-button secondary" onClick={handleReset}>
                ← Back
              </button>
              <span className="text-xl">{selectedTemplate.icon}</span>
              <span className="text-bold">{selectedTemplate.name}</span>
            </div>

            <label className="form-label">Add context (optional):</label>
            <textarea
              className="chat-input input-rounded"
              placeholder="E.g., Meeting is about Q4 planning, recipient is John from Sales..."
              value={customContext}
              onChange={(e) => setCustomContext(e.target.value)}
              rows={3}
              style={{ width: '100%', resize: 'vertical' }}
            />

            <button 
              className="action-button primary mt-12"
              onClick={handleGenerate}
              disabled={loading}
              style={{ width: '100%' }}
            >
              {loading ? '⏳ Generating...' : '✨ Generate Email'}
            </button>
          </div>
        ) : (
          <div>
            <div className="flex items-center gap-8 mb-12">
              <button className="action-button secondary" onClick={handleReset}>
                ← Back
              </button>
              <span className="text-bold">Generated Email</span>
            </div>

            <div className="card" style={{ maxHeight: 300, overflow: 'auto' }}>
              <pre className="text-small pre-wrap">{generatedContent}</pre>
            </div>

            <div className="flex gap-8 mt-12">
              <button 
                className="action-button primary flex-1"
                onClick={handleUse}
              >
                📋 Use This Email
              </button>
              <button 
                className="action-button secondary"
                onClick={handleGenerate}
                disabled={loading}
              >
                🔄 Regenerate
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default TemplatesModal;
