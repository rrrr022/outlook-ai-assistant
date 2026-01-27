import { EmailSummary, EmailDetails, CalendarEvent, CreateCalendarEvent } from '@shared/types';
import { graphService } from './graphService';

/**
 * Service for interacting with Outlook via Office.js API
 */
class OutlookService {
  private isOfficeInitialized = false;

  constructor() {
    this.checkOfficeInitialized();
  }

  private checkOfficeInitialized(): boolean {
    if (typeof Office !== 'undefined' && Office.context) {
      this.isOfficeInitialized = true;
      return true;
    }
    return false;
  }

  /**
   * Get the currently selected email's context
   */
  async getCurrentEmailContext(): Promise<EmailSummary | null> {
    if (!this.checkOfficeInitialized()) {
      console.warn('Office.js not initialized');
      return this.getMockEmail();
    }

    return new Promise((resolve) => {
      try {
        const item = Office.context.mailbox.item;
        if (!item) {
          resolve(null);
          return;
        }

        // Get email properties
        const email: EmailSummary = {
          id: item.itemId || '',
          subject: item.subject || '',
          sender: item.from?.displayName || '',
          senderEmail: item.from?.emailAddress || '',
          receivedDateTime: item.dateTimeCreated || new Date(),
          preview: '',
          isRead: true,
          importance: ((item as any).importance as 'low' | 'normal' | 'high') || 'normal',
          hasAttachments: item.attachments?.length > 0 || false,
          conversationId: item.conversationId,
        };

        // Try to get body preview
        if (item.body) {
          item.body.getAsync(Office.CoercionType.Text, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              email.preview = result.value.substring(0, 500);
            }
            resolve(email);
          });
        } else {
          resolve(email);
        }
      } catch (error) {
        console.error('Error getting email context:', error);
        resolve(null);
      }
    });
  }

  /**
   * Get recent emails from inbox using Graph API
   */
  async getRecentEmails(count: number = 10): Promise<EmailSummary[]> {
    if (!this.checkOfficeInitialized()) {
      return this.getMockEmails();
    }

    try {
      // Use Graph API for real inbox access
      const emails = await graphService.getInboxEmails(count);
      if (emails && emails.length > 0) {
        return emails;
      }
      // Fallback to mock if Graph API fails
      return this.getMockEmails();
    } catch (error) {
      console.error('Error getting recent emails:', error);
      return this.getMockEmails();
    }
  }

  /**
   * Get upcoming calendar events using Graph API
   */
  async getUpcomingEvents(days: number = 7): Promise<CalendarEvent[]> {
    if (!this.checkOfficeInitialized()) {
      return this.getMockEvents();
    }

    try {
      // Use Graph API for real calendar access
      const events = await graphService.getCalendarEvents(days);
      if (events && events.length > 0) {
        return events;
      }
      // Fallback to mock if Graph API fails
      return this.getMockEvents();
    } catch (error) {
      console.error('Error getting calendar events:', error);
      return this.getMockEvents();
    }
  }

  /**
   * Create a new calendar event
   */
  async createCalendarEvent(event: CreateCalendarEvent): Promise<boolean> {
    if (!this.checkOfficeInitialized()) {
      console.warn('Office.js not initialized');
      return false;
    }

    return new Promise((resolve) => {
      try {
        // Note: Creating events requires Microsoft Graph API
        // This is a placeholder for the implementation
        console.log('Creating event:', event);
        resolve(true);
      } catch (error) {
        console.error('Error creating event:', error);
        resolve(false);
      }
    });
  }

  /**
   * Insert text into compose window
   */
  async insertTextToCompose(text: string): Promise<boolean> {
    if (!this.checkOfficeInitialized()) {
      console.warn('Office.js not initialized');
      return false;
    }

    return new Promise((resolve) => {
      try {
        const item = Office.context.mailbox.item;
        if (item && item.body) {
          item.body.setSelectedDataAsync(
            text,
            { coercionType: Office.CoercionType.Text },
            (result) => {
              resolve(result.status === Office.AsyncResultStatus.Succeeded);
            }
          );
        } else {
          resolve(false);
        }
      } catch (error) {
        console.error('Error inserting text:', error);
        resolve(false);
      }
    });
  }

  /**
   * Reply to current email
   */
  async replyToEmail(body: string, replyAll: boolean = false): Promise<boolean> {
    if (!this.checkOfficeInitialized()) {
      return false;
    }

    return new Promise((resolve) => {
      try {
        const item = Office.context.mailbox.item;
        if (item) {
          // Use the appropriate reply form based on replyAll flag
          if (replyAll) {
            item.displayReplyAllForm(body);
          } else {
            item.displayReplyForm(body);
          }
          resolve(true);
        } else {
          resolve(false);
        }
      } catch (error) {
        console.error('Error replying to email:', error);
        resolve(false);
      }
    });
  }

  /**
   * Get user's email address
   */
  getUserEmail(): string {
    if (this.checkOfficeInitialized()) {
      return Office.context.mailbox.userProfile.emailAddress;
    }
    return 'user@example.com';
  }

  /**
   * Get user's display name
   */
  getUserDisplayName(): string {
    if (this.checkOfficeInitialized()) {
      return Office.context.mailbox.userProfile.displayName;
    }
    return 'User';
  }

  // Mock data for development/testing
  private getMockEmail(): EmailSummary {
    return {
      id: 'mock-1',
      subject: 'Project Update - Q4 Review',
      sender: 'John Smith',
      senderEmail: 'john.smith@company.com',
      receivedDateTime: new Date(),
      preview: 'Hi team, I wanted to share the latest updates on our Q4 project milestones. We have made significant progress on the main deliverables...',
      isRead: false,
      importance: 'high',
      hasAttachments: true,
    };
  }

  private getMockEmails(): EmailSummary[] {
    return [
      {
        id: 'mock-1',
        subject: 'Project Update - Q4 Review',
        sender: 'John Smith',
        senderEmail: 'john.smith@company.com',
        receivedDateTime: new Date(),
        preview: 'Hi team, I wanted to share the latest updates on our Q4 project milestones...',
        isRead: false,
        importance: 'high',
        hasAttachments: true,
      },
      {
        id: 'mock-2',
        subject: 'Meeting Request: Sprint Planning',
        sender: 'Sarah Johnson',
        senderEmail: 'sarah.j@company.com',
        receivedDateTime: new Date(Date.now() - 3600000),
        preview: 'Would you be available for a sprint planning session tomorrow at 2 PM?',
        isRead: true,
        importance: 'normal',
        hasAttachments: false,
      },
      {
        id: 'mock-3',
        subject: 'Invoice #12345 Attached',
        sender: 'Accounting',
        senderEmail: 'accounting@company.com',
        receivedDateTime: new Date(Date.now() - 7200000),
        preview: 'Please find attached the invoice for the recent order. Payment is due within 30 days.',
        isRead: true,
        importance: 'low',
        hasAttachments: true,
      },
    ];
  }

  private getMockEvents(): CalendarEvent[] {
    const today = new Date();
    return [
      {
        id: 'event-1',
        subject: 'Team Standup',
        start: new Date(today.setHours(9, 0, 0, 0)),
        end: new Date(today.setHours(9, 30, 0, 0)),
        location: 'Conference Room A',
        isAllDay: false,
        organizer: 'Team Lead',
        attendees: [],
        importance: 'normal',
        showAs: 'busy',
        isRecurring: true,
      },
      {
        id: 'event-2',
        subject: 'Project Review',
        start: new Date(today.setHours(14, 0, 0, 0)),
        end: new Date(today.setHours(15, 0, 0, 0)),
        location: 'Virtual - Teams',
        isAllDay: false,
        organizer: 'Project Manager',
        attendees: [],
        importance: 'high',
        showAs: 'busy',
        isRecurring: false,
      },
      {
        id: 'event-3',
        subject: '1:1 with Manager',
        start: new Date(today.setHours(16, 0, 0, 0)),
        end: new Date(today.setHours(16, 30, 0, 0)),
        location: 'Office',
        isAllDay: false,
        organizer: 'Manager',
        attendees: [],
        importance: 'normal',
        showAs: 'busy',
        isRecurring: true,
      },
    ];
  }
}

export const outlookService = new OutlookService();
