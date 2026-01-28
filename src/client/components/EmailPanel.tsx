import React, { useEffect, useState } from 'react';
import { Button, Spinner } from './ui/NativeComponents';
import { useAppStore } from '../store/appStore';
import { outlookService } from '../services/outlookService';
import { aiService } from '../services/aiService';
import { approvalService } from '../services/approvalService';
import { EmailSummary } from '../../shared/types';
import TemplatesModal from './TemplatesModal';

const EmailPanel: React.FC = () => {
  const { currentEmail, setCurrentEmail, emailSummaries, setEmailSummaries, settings } = useAppStore();
  const [loading, setLoading] = useState(false);
  const [summary, setSummary] = useState<string | null>(null);
  const [draftReply, setDraftReply] = useState<string | null>(null);
  const [showTemplates, setShowTemplates] = useState(false);
  const [sentiment, setSentiment] = useState<'positive' | 'neutral' | 'negative' | null>(null);
  const [quickReplies, setQuickReplies] = useState<string[]>([]);

  useEffect(() => {
    loadCurrentEmail();
    loadRecentEmails();
  }, []);

  const loadCurrentEmail = async () => {
    try {
      const email = await outlookService.getCurrentEmailContext();
      if (email) {
        setCurrentEmail(email);
      }
    } catch (error) {
      console.error('Error loading current email:', error);
    }
  };

  const loadRecentEmails = async () => {
    try {
      setLoading(true);
      const emails = await outlookService.getRecentEmails(10);
      setEmailSummaries(emails);
    } catch (error) {
      console.error('Error loading emails:', error);
    } finally {
      setLoading(false);
    }
  };

  const handleSummarize = async () => {
    if (!currentEmail) return;
    
    setLoading(true);
    try {
      const response = await aiService.chat({
        prompt: `Please summarize this email concisely and also analyze the sentiment (positive, neutral, or negative). Start with "Sentiment: [sentiment]" on the first line, then provide the summary:\n\nFrom: ${currentEmail.sender}\nSubject: ${currentEmail.subject}\n\nPreview: ${currentEmail.preview}`,
      });
      
      // Parse sentiment from response
      const lines = response.content.split('\n');
      const sentimentLine = lines[0].toLowerCase();
      if (sentimentLine.includes('positive')) setSentiment('positive');
      else if (sentimentLine.includes('negative')) setSentiment('negative');
      else setSentiment('neutral');
      
      setSummary(lines.slice(1).join('\n').trim() || response.content);
      
      // Generate quick reply suggestions
      generateQuickReplies();
    } catch (error) {
      console.error('Error summarizing:', error);
    } finally {
      setLoading(false);
    }
  };

  const generateQuickReplies = async () => {
    if (!currentEmail) return;
    try {
      const response = await aiService.chat({
        prompt: `Generate 3 short quick reply options (each under 10 words) for this email. Return only the replies, one per line:\n\nFrom: ${currentEmail.sender}\nSubject: ${currentEmail.subject}`,
      });
      const replies = response.content.split('\n').filter(r => r.trim()).slice(0, 3);
      setQuickReplies(replies);
    } catch (e) {
      setQuickReplies(['Thanks for the update!', 'I\'ll look into this.', 'Let\'s discuss this further.']);
    }
  };

  const handleDraftReply = async () => {
    if (!currentEmail) return;
    
    setLoading(true);
    try {
      const response = await aiService.chat({
        prompt: `Draft a professional reply to this email:\n\nFrom: ${currentEmail.sender}\nSubject: ${currentEmail.subject}\n\nContent: ${currentEmail.preview}`,
      });
      setDraftReply(response.content);
    } catch (error) {
      console.error('Error drafting reply:', error);
    } finally {
      setLoading(false);
    }
  };

  const handleInsertReply = async () => {
    if (!draftReply || !currentEmail) return;
    
    // Request approval before sending
    if (settings.requireApprovalForEmails) {
      await approvalService.requestReplyApproval({
        to: currentEmail.sender,
        subject: currentEmail.subject,
        body: draftReply,
        source: 'ai',
      });
    } else {
      try {
        await outlookService.insertTextToCompose(draftReply);
      } catch (error) {
        console.error('Error inserting reply:', error);
      }
    }
  };

  const handleQuickReply = async (reply: string) => {
    if (!currentEmail) return;
    
    // Request approval for quick replies too
    if (settings.requireApprovalForEmails) {
      await approvalService.requestReplyApproval({
        to: currentEmail.sender,
        subject: currentEmail.subject,
        body: reply,
        source: 'ai',
      });
    } else {
      try {
        await outlookService.insertTextToCompose(reply);
      } catch (error) {
        setDraftReply(reply);
      }
    }
  };

  const handleTemplateSelect = (content: string) => {
    setDraftReply(content);
  };

  return (
    <div>
      <h2 className="section-title">📧 Email Assistant</h2>

      {/* Templates Button */}
      <div className="card mb-12">
        <button 
          className="template-btn" 
          onClick={() => setShowTemplates(true)}
          style={{ width: '100%', justifyContent: 'center' }}
        >
          <span>📝</span>
          <span>Use Email Template</span>
        </button>
      </div>

      {/* Current Email Section */}
      {currentEmail ? (
        <div className="card">
          <h3 className="section-header">
            Current Email
          </h3>
          <div className="email-subject">{currentEmail.subject}</div>
          <div className="email-sender">From: {currentEmail.sender}</div>
          <div className="email-preview">{currentEmail.preview}</div>
          
          <div className="email-actions">
            <button 
              className="action-button primary" 
              onClick={handleSummarize}
              disabled={loading}
            >
              📝 Summarize
            </button>
            <button 
              className="action-button secondary" 
              onClick={handleDraftReply}
              disabled={loading}
            >
              ✉️ Draft Reply
            </button>
          </div>
        </div>
      ) : (
        <div className="card">
          <div className="empty-state">
            <div className="empty-state-icon">📭</div>
            <p className="empty-state-text">
              Select an email to get started
            </p>
          </div>
        </div>
      )}

      {/* AI Summary */}
      {summary && (
        <div className="card">
          <h3 className="section-header">
            📝 Summary
            {sentiment && (
              <span className={`sentiment-indicator sentiment-${sentiment}`} style={{ marginLeft: 8 }}>
                {sentiment === 'positive' ? '😊' : sentiment === 'negative' ? '😟' : '😐'} {sentiment}
              </span>
            )}
          </h3>
          <p className="text-medium text-dark">{summary}</p>
          
          {/* Quick Reply Chips */}
          {quickReplies.length > 0 && (
            <div className="quick-replies">
              {quickReplies.map((reply, index) => (
                <button
                  key={index}
                  className="quick-reply-chip"
                  onClick={() => handleQuickReply(reply)}
                >
                  {reply}
                </button>
              ))}
            </div>
          )}
        </div>
      )}

      {/* Draft Reply */}
      {draftReply && (
        <div className="card">
          <h3 className="section-header">
            ✉️ Draft Reply
          </h3>
          <p className="text-medium text-dark pre-wrap">
            {draftReply}
          </p>
          <button 
            className="action-button primary action-button--mt" 
            onClick={handleInsertReply}
          >
            Insert into Reply
          </button>
        </div>
      )}

      {/* Recent Emails */}
      <div className="mt-16">
        <h3 className="section-title">Recent Emails</h3>
        {loading ? (
          <div className="loading-spinner">
            <Spinner size="small" />
          </div>
        ) : emailSummaries.length > 0 ? (
          emailSummaries.map((email) => (
            <div 
              key={email.id} 
              className="card email-item card--clickable"
              onClick={() => setCurrentEmail(email)}
            >
              <div className="email-subject">
                {!email.isRead && <span className="unread-indicator">● </span>}
                {email.subject}
              </div>
              <div className="email-sender">{email.sender}</div>
              <div className="email-preview">{email.preview}</div>
            </div>
          ))
        ) : (
          <div className="empty-state">
            <p className="empty-state-text">No recent emails</p>
          </div>
        )}
      </div>

      {/* Templates Modal */}
      <TemplatesModal 
        isOpen={showTemplates} 
        onClose={() => setShowTemplates(false)}
        onSelectTemplate={handleTemplateSelect}
      />
    </div>
  );
};

export default EmailPanel;
