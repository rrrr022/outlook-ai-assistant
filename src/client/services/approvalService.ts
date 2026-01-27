import { v4 as uuidv4 } from 'uuid';
import { useAppStore } from '../store/appStore';
import { ApprovalRequest, ApprovalAction } from '../components/ApprovalModal';
import { outlookService } from './outlookService';

type ApprovalCallback = (approved: boolean, request: ApprovalRequest) => void;

class ApprovalService {
  private callbacks: Map<string, ApprovalCallback> = new Map();

  /**
   * Request approval before sending an email
   */
  async requestEmailApproval(params: {
    to: string;
    subject: string;
    body: string;
    source?: 'user' | 'automation' | 'ai';
    ruleName?: string;
  }): Promise<boolean> {
    const store = useAppStore.getState();
    
    // If approval not required, proceed directly
    if (!store.settings.requireApprovalForEmails && params.source !== 'automation') {
      return true;
    }

    const request: ApprovalRequest = {
      id: uuidv4(),
      action: 'send_email',
      title: `Send email to ${params.to}`,
      description: `Subject: ${params.subject}`,
      details: {
        to: params.to,
        subject: params.subject,
        body: params.body,
        ruleName: params.ruleName,
      },
      timestamp: new Date(),
      source: params.source || 'user',
    };

    return this.requestApproval(request);
  }

  /**
   * Request approval before sending a reply
   */
  async requestReplyApproval(params: {
    to: string;
    subject: string;
    body: string;
    source?: 'user' | 'automation' | 'ai';
  }): Promise<boolean> {
    const store = useAppStore.getState();
    
    if (!store.settings.requireApprovalForEmails) {
      return true;
    }

    const request: ApprovalRequest = {
      id: uuidv4(),
      action: 'reply_email',
      title: `Reply to ${params.to}`,
      description: `Re: ${params.subject}`,
      details: {
        to: params.to,
        subject: `Re: ${params.subject}`,
        body: params.body,
      },
      timestamp: new Date(),
      source: params.source || 'user',
    };

    return this.requestApproval(request);
  }

  /**
   * Request approval before creating a meeting
   */
  async requestMeetingApproval(params: {
    subject: string;
    startTime: Date;
    endTime: Date;
    attendees?: string[];
    location?: string;
    source?: 'user' | 'automation' | 'ai';
    ruleName?: string;
  }): Promise<boolean> {
    const store = useAppStore.getState();
    
    if (!store.settings.requireApprovalForMeetings && params.source !== 'automation') {
      return true;
    }

    const request: ApprovalRequest = {
      id: uuidv4(),
      action: 'create_meeting',
      title: params.subject,
      description: `Meeting with ${params.attendees?.length || 0} attendee(s)`,
      details: {
        subject: params.subject,
        startTime: params.startTime,
        endTime: params.endTime,
        attendees: params.attendees,
        location: params.location,
        ruleName: params.ruleName,
      },
      timestamp: new Date(),
      source: params.source || 'user',
    };

    return this.requestApproval(request);
  }

  /**
   * Request approval for an automation rule action
   */
  async requestAutomationApproval(params: {
    action: ApprovalAction;
    title: string;
    description: string;
    ruleName: string;
    details: ApprovalRequest['details'];
  }): Promise<boolean> {
    const store = useAppStore.getState();
    
    if (!store.settings.requireApprovalForAutomations) {
      return true;
    }

    const request: ApprovalRequest = {
      id: uuidv4(),
      action: params.action,
      title: params.title,
      description: params.description,
      details: {
        ...params.details,
        ruleName: params.ruleName,
      },
      timestamp: new Date(),
      source: 'automation',
    };

    return this.requestApproval(request);
  }

  /**
   * Core approval request method
   */
  private requestApproval(request: ApprovalRequest): Promise<boolean> {
    return new Promise((resolve) => {
      // Store the callback
      this.callbacks.set(request.id, (approved) => {
        resolve(approved);
      });

      // Add to the approval queue
      useAppStore.getState().addApprovalRequest(request);
    });
  }

  /**
   * Handle approval response
   */
  async handleApproval(request: ApprovalRequest): Promise<void> {
    const callback = this.callbacks.get(request.id);
    
    // Execute the actual action
    try {
      switch (request.action) {
        case 'send_email':
        case 'reply_email':
        case 'forward_email':
          // Insert the email body (actual sending would be done via Office.js)
          if (request.details.body) {
            await outlookService.insertTextToCompose(request.details.body);
          }
          break;
        
        case 'create_meeting':
          if (request.details.startTime && request.details.endTime) {
            await outlookService.createCalendarEvent({
              subject: request.details.subject || 'New Meeting',
              start: request.details.startTime,
              end: request.details.endTime,
              location: request.details.location,
            });
          }
          break;
        
        case 'create_task':
          // Task creation is handled in the store
          break;
      }
      
      if (callback) {
        callback(true, request);
        this.callbacks.delete(request.id);
      }
    } catch (error) {
      console.error('Error executing approved action:', error);
      if (callback) {
        callback(false, request);
        this.callbacks.delete(request.id);
      }
    }

    // Remove from queue
    useAppStore.getState().removeApprovalRequest(request.id);
  }

  /**
   * Handle rejection
   */
  handleRejection(request: ApprovalRequest): void {
    const callback = this.callbacks.get(request.id);
    
    if (callback) {
      callback(false, request);
      this.callbacks.delete(request.id);
    }

    // Remove from queue
    useAppStore.getState().removeApprovalRequest(request.id);
  }

  /**
   * Get pending approval count
   */
  getPendingCount(): number {
    return useAppStore.getState().pendingApprovals.length;
  }
}

export const approvalService = new ApprovalService();
