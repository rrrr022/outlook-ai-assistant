import React from 'react';

export type ApprovalAction = 
  | 'send_email' 
  | 'create_meeting' 
  | 'reply_email' 
  | 'forward_email'
  | 'create_task'
  | 'delete_email'
  | 'auto_reply';

export interface ApprovalRequest {
  id: string;
  action: ApprovalAction;
  title: string;
  description: string;
  details: {
    to?: string;
    subject?: string;
    body?: string;
    startTime?: Date;
    endTime?: Date;
    location?: string;
    attendees?: string[];
    taskTitle?: string;
    ruleName?: string;
  };
  timestamp: Date;
  source: 'user' | 'automation' | 'ai';
}

interface ApprovalModalProps {
  request: ApprovalRequest | null;
  onApprove: (request: ApprovalRequest) => void;
  onReject: (request: ApprovalRequest) => void;
  onEdit?: (request: ApprovalRequest) => void;
}

const getActionIcon = (action: ApprovalAction): string => {
  switch (action) {
    case 'send_email':
    case 'reply_email':
    case 'forward_email':
      return '✉️';
    case 'create_meeting':
      return '📅';
    case 'create_task':
      return '✅';
    case 'delete_email':
      return '🗑️';
    case 'auto_reply':
      return '🤖';
    default:
      return '⚡';
  }
};

const getActionLabel = (action: ApprovalAction): string => {
  switch (action) {
    case 'send_email':
      return 'Send Email';
    case 'reply_email':
      return 'Send Reply';
    case 'forward_email':
      return 'Forward Email';
    case 'create_meeting':
      return 'Create Meeting';
    case 'create_task':
      return 'Create Task';
    case 'delete_email':
      return 'Delete Email';
    case 'auto_reply':
      return 'Auto-Reply';
    default:
      return 'Action';
  }
};

const getSourceLabel = (source: 'user' | 'automation' | 'ai'): string => {
  switch (source) {
    case 'user':
      return 'Your Request';
    case 'automation':
      return 'Automation Rule';
    case 'ai':
      return 'AI Assistant';
  }
};

const ApprovalModal: React.FC<ApprovalModalProps> = ({ 
  request, 
  onApprove, 
  onReject, 
  onEdit 
}) => {
  if (!request) return null;

  const formatDateTime = (date: Date) => {
    return new Date(date).toLocaleString('en-US', {
      weekday: 'short',
      month: 'short',
      day: 'numeric',
      hour: 'numeric',
      minute: '2-digit',
    });
  };

  return (
    <div className="modal-overlay">
      <div className="modal-content approval-modal">
        {/* Header */}
        <div className="approval-header">
          <div className="approval-icon-large">
            {getActionIcon(request.action)}
          </div>
          <div>
            <h3 className="approval-title">Confirm Action</h3>
            <p className="approval-subtitle">
              {getSourceLabel(request.source)} wants to {getActionLabel(request.action).toLowerCase()}
            </p>
          </div>
        </div>

        {/* Warning Banner */}
        <div className="approval-warning">
          <span>⚠️</span>
          <span>Please review before approving. This action cannot be undone.</span>
        </div>

        {/* Action Details */}
        <div className="approval-details">
          <h4>{request.title}</h4>
          <p className="text-muted">{request.description}</p>

          {/* Email Details */}
          {(request.action === 'send_email' || 
            request.action === 'reply_email' || 
            request.action === 'forward_email') && (
            <div className="detail-section">
              {request.details.to && (
                <div className="detail-row">
                  <span className="detail-label">To:</span>
                  <span className="detail-value">{request.details.to}</span>
                </div>
              )}
              {request.details.subject && (
                <div className="detail-row">
                  <span className="detail-label">Subject:</span>
                  <span className="detail-value">{request.details.subject}</span>
                </div>
              )}
              {request.details.body && (
                <div className="detail-row">
                  <span className="detail-label">Message:</span>
                  <div className="detail-preview">
                    {request.details.body.substring(0, 300)}
                    {request.details.body.length > 300 && '...'}
                  </div>
                </div>
              )}
            </div>
          )}

          {/* Meeting Details */}
          {request.action === 'create_meeting' && (
            <div className="detail-section">
              {request.details.startTime && (
                <div className="detail-row">
                  <span className="detail-label">When:</span>
                  <span className="detail-value">
                    {formatDateTime(request.details.startTime)}
                    {request.details.endTime && ` - ${formatDateTime(request.details.endTime)}`}
                  </span>
                </div>
              )}
              {request.details.location && (
                <div className="detail-row">
                  <span className="detail-label">Where:</span>
                  <span className="detail-value">{request.details.location}</span>
                </div>
              )}
              {request.details.attendees && request.details.attendees.length > 0 && (
                <div className="detail-row">
                  <span className="detail-label">Attendees:</span>
                  <span className="detail-value">{request.details.attendees.join(', ')}</span>
                </div>
              )}
            </div>
          )}

          {/* Task Details */}
          {request.action === 'create_task' && request.details.taskTitle && (
            <div className="detail-section">
              <div className="detail-row">
                <span className="detail-label">Task:</span>
                <span className="detail-value">{request.details.taskTitle}</span>
              </div>
            </div>
          )}

          {/* Automation Rule */}
          {request.source === 'automation' && request.details.ruleName && (
            <div className="detail-row mt-8">
              <span className="detail-label">Triggered by:</span>
              <span className="detail-value">{request.details.ruleName}</span>
            </div>
          )}
        </div>

        {/* Timestamp */}
        <div className="approval-timestamp">
          Requested {formatDateTime(request.timestamp)}
        </div>

        {/* Actions */}
        <div className="approval-actions">
          <button 
            className="action-button secondary"
            onClick={() => onReject(request)}
          >
            ❌ Reject
          </button>
          {onEdit && (
            <button 
              className="action-button secondary"
              onClick={() => onEdit(request)}
            >
              ✏️ Edit
            </button>
          )}
          <button 
            className="action-button primary approval-approve-btn"
            onClick={() => onApprove(request)}
          >
            ✅ Approve & {getActionLabel(request.action)}
          </button>
        </div>
      </div>
    </div>
  );
};

export default ApprovalModal;
