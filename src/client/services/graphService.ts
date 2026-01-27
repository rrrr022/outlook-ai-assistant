import { EmailSummary, CalendarEvent } from '../../shared/types';
import { msalService } from './msalService';

/**
 * Microsoft Graph API Service for accessing full mailbox
 * Uses MSAL.js for authentication with Azure AD
 */
class GraphService {
  private _isRealDataMode: boolean = false;
  private _lastError: string | null = null;
  private _isSignedIn: boolean = false;

  /**
   * Check if we're using real data (vs mock data)
   */
  get isRealDataMode(): boolean {
    return this._isRealDataMode;
  }

  /**
   * Get last error message
   */
  get lastError(): string | null {
    return this._lastError;
  }

  /**
   * Check if user is signed in to Microsoft Graph
   */
  get isSignedIn(): boolean {
    return this._isSignedIn || msalService.isSignedIn();
  }

  /**
   * Sign in to Microsoft Graph
   */
  async signIn(): Promise<boolean> {
    try {
      await msalService.signIn();
      this._isSignedIn = msalService.isSignedIn();
      this._isRealDataMode = this._isSignedIn;
      this._lastError = null;
      return this._isSignedIn;
    } catch (error: any) {
      this._lastError = error.message || 'Sign in failed';
      console.error('Graph sign in failed:', error);
      return false;
    }
  }

  /**
   * Sign out from Microsoft Graph
   */
  async signOut(): Promise<void> {
    await msalService.signOut();
    this._isSignedIn = false;
    this._isRealDataMode = false;
  }

  /**
   * Get user info
   */
  getUserInfo() {
    return msalService.getUserInfo();
  }

  /**
   * Get access token using MSAL
   */
  async getAccessToken(): Promise<string | null> {
    // First try MSAL (for full inbox access)
    if (msalService.isSignedIn()) {
      const token = await msalService.getAccessToken();
      if (token) {
        this._isRealDataMode = true;
        this._isSignedIn = true;
        this._lastError = null;
        console.log('✅ Got MSAL token for Graph API');
        return token;
      }
    }

    // Fallback to Office.js REST token (limited to current email)
    if (typeof Office !== 'undefined' && Office.context?.mailbox) {
      return new Promise((resolve) => {
        Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            this._isRealDataMode = true;
            this._lastError = null;
            console.log('✅ Got Office.js REST token');
            resolve(result.value);
          } else {
            this._isRealDataMode = false;
            this._lastError = result.error?.message || 'Token error';
            resolve(null);
          }
        });
      });
    }

    this._isRealDataMode = false;
    this._lastError = 'Not signed in. Click "Sign in to Microsoft" for full inbox access.';
    return null;
  }

  /**
   * Get the REST API endpoint URL
   */
  private getRestUrl(): string {
    // When using MSAL, use Microsoft Graph API
    if (msalService.isSignedIn()) {
      return 'https://graph.microsoft.com/v1.0';
    }
    // When using Office.js token, use the Outlook REST API
    if (typeof Office !== 'undefined' && Office.context?.mailbox?.restUrl) {
      return Office.context.mailbox.restUrl;
    }
    return 'https://graph.microsoft.com/v1.0';
  }

  /**
   * Fetch emails from inbox using REST API
   */
  async getInboxEmails(count: number = 50, skip: number = 0): Promise<EmailSummary[]> {
    const token = await this.getAccessToken();
    if (!token) {
      console.warn('No access token available, returning mock data');
      return this.getMockEmails(count);
    }

    try {
      const restUrl = this.getRestUrl();
      const url = `${restUrl}/me/mailfolders/inbox/messages?$top=${count}&$skip=${skip}&$select=id,subject,from,receivedDateTime,bodyPreview,isRead,importance,hasAttachments,conversationId&$orderby=receivedDateTime desc`;

      const response = await fetch(url, {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        throw new Error(`Graph API error: ${response.status}`);
      }

      const data = await response.json();
      return data.value.map((msg: any) => ({
        id: msg.id,
        subject: msg.subject || '(No Subject)',
        sender: msg.from?.emailAddress?.name || 'Unknown',
        senderEmail: msg.from?.emailAddress?.address || '',
        receivedDateTime: new Date(msg.receivedDateTime),
        preview: msg.bodyPreview || '',
        isRead: msg.isRead,
        importance: msg.importance?.toLowerCase() || 'normal',
        hasAttachments: msg.hasAttachments,
        conversationId: msg.conversationId,
      }));
    } catch (error) {
      console.error('Error fetching inbox emails:', error);
      return this.getMockEmails(count);
    }
  }

  /**
   * Get unread email count
   */
  async getUnreadCount(): Promise<number> {
    const token = await this.getAccessToken();
    if (!token) return 0;

    try {
      const restUrl = this.getRestUrl();
      const url = `${restUrl}/me/mailfolders/inbox?$select=unreadItemCount`;

      const response = await fetch(url, {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) return 0;

      const data = await response.json();
      return data.unreadItemCount || 0;
    } catch (error) {
      console.error('Error getting unread count:', error);
      return 0;
    }
  }

  /**
   * Search emails in inbox
   */
  async searchEmails(query: string, count: number = 20): Promise<EmailSummary[]> {
    const token = await this.getAccessToken();
    if (!token) return [];

    try {
      const restUrl = this.getRestUrl();
      const url = `${restUrl}/me/messages?$search="${encodeURIComponent(query)}"&$top=${count}&$select=id,subject,from,receivedDateTime,bodyPreview,isRead,importance,hasAttachments,conversationId`;

      const response = await fetch(url, {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) return [];

      const data = await response.json();
      return data.value.map((msg: any) => ({
        id: msg.id,
        subject: msg.subject || '(No Subject)',
        sender: msg.from?.emailAddress?.name || 'Unknown',
        senderEmail: msg.from?.emailAddress?.address || '',
        receivedDateTime: new Date(msg.receivedDateTime),
        preview: msg.bodyPreview || '',
        isRead: msg.isRead,
        importance: msg.importance?.toLowerCase() || 'normal',
        hasAttachments: msg.hasAttachments,
        conversationId: msg.conversationId,
      }));
    } catch (error) {
      console.error('Error searching emails:', error);
      return [];
    }
  }

  /**
   * Get emails from a specific sender
   */
  async getEmailsFromSender(senderEmail: string, count: number = 20): Promise<EmailSummary[]> {
    const token = await this.getAccessToken();
    if (!token) return [];

    try {
      const restUrl = this.getRestUrl();
      const url = `${restUrl}/me/messages?$filter=from/emailAddress/address eq '${encodeURIComponent(senderEmail)}'&$top=${count}&$select=id,subject,from,receivedDateTime,bodyPreview,isRead,importance,hasAttachments,conversationId&$orderby=receivedDateTime desc`;

      const response = await fetch(url, {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) return [];

      const data = await response.json();
      return data.value.map((msg: any) => ({
        id: msg.id,
        subject: msg.subject || '(No Subject)',
        sender: msg.from?.emailAddress?.name || 'Unknown',
        senderEmail: msg.from?.emailAddress?.address || '',
        receivedDateTime: new Date(msg.receivedDateTime),
        preview: msg.bodyPreview || '',
        isRead: msg.isRead,
        importance: msg.importance?.toLowerCase() || 'normal',
        hasAttachments: msg.hasAttachments,
        conversationId: msg.conversationId,
      }));
    } catch (error) {
      console.error('Error getting emails from sender:', error);
      return [];
    }
  }

  /**
   * Get full email details including body
   */
  async getEmailDetails(emailId: string): Promise<any | null> {
    const token = await this.getAccessToken();
    if (!token) return null;

    try {
      const restUrl = this.getRestUrl();
      const url = `${restUrl}/me/messages/${emailId}?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,bodyPreview,isRead,importance,hasAttachments,conversationId,attachments`;

      const response = await fetch(url, {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) return null;

      return await response.json();
    } catch (error) {
      console.error('Error getting email details:', error);
      return null;
    }
  }

  /**
   * Get inbox summary statistics
   */
  async getInboxSummary(): Promise<{
    totalEmails: number;
    unreadCount: number;
    topSenders: { name: string; email: string; count: number }[];
    recentEmails: EmailSummary[];
    isRealData: boolean;
    error?: string;
  }> {
    // First try to get real data
    const [emails, unreadCount] = await Promise.all([
      this.getInboxEmails(100),
      this.getUnreadCount(),
    ]);

    const isRealData = this._isRealDataMode;

    // Calculate top senders
    const senderCounts: Record<string, { name: string; email: string; count: number }> = {};
    emails.forEach((email) => {
      const key = email.senderEmail;
      if (!senderCounts[key]) {
        senderCounts[key] = { name: email.sender, email: email.senderEmail, count: 0 };
      }
      senderCounts[key].count++;
    });

    const topSenders = Object.values(senderCounts)
      .sort((a, b) => b.count - a.count)
      .slice(0, 10);

    return {
      totalEmails: emails.length,
      unreadCount,
      topSenders,
      recentEmails: emails.slice(0, 20),
      isRealData,
      error: this._lastError || undefined,
    };
  }

  /**
   * Get calendar events for the specified number of days
   */
  async getCalendarEvents(days: number = 7): Promise<CalendarEvent[]> {
    const token = await this.getAccessToken();
    if (!token) {
      return this.getMockCalendarEvents();
    }

    try {
      const restUrl = this.getRestUrl();
      const startDate = new Date().toISOString();
      const endDate = new Date(Date.now() + days * 24 * 60 * 60 * 1000).toISOString();
      
      const url = `${restUrl}/me/calendarview?startDateTime=${startDate}&endDateTime=${endDate}&$select=id,subject,start,end,location,isAllDay,organizer,attendees,body,importance,showAs,isRecurring&$orderby=start/dateTime`;

      const response = await fetch(url, {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
          'Prefer': 'outlook.timezone="UTC"',
        },
      });

      if (!response.ok) {
        console.error('Calendar API error:', response.status);
        return this.getMockCalendarEvents();
      }

      const data = await response.json();
      return data.value.map((event: any) => ({
        id: event.id,
        subject: event.subject || '(No Title)',
        start: new Date(event.start?.dateTime || event.start),
        end: new Date(event.end?.dateTime || event.end),
        location: event.location?.displayName || '',
        isAllDay: event.isAllDay || false,
        organizer: event.organizer?.emailAddress?.name || 'Unknown',
        attendees: (event.attendees || []).map((a: any) => ({
          name: a.emailAddress?.name || '',
          email: a.emailAddress?.address || '',
          response: a.status?.response || 'none',
          type: a.type || 'required',
        })),
        body: event.body?.content || '',
        importance: event.importance?.toLowerCase() || 'normal',
        showAs: event.showAs?.toLowerCase() || 'busy',
        isRecurring: event.isRecurring || false,
      }));
    } catch (error) {
      console.error('Error fetching calendar events:', error);
      return this.getMockCalendarEvents();
    }
  }

  /**
   * Mock calendar events for development/testing
   */
  private getMockCalendarEvents(): CalendarEvent[] {
    const now = new Date();
    return [
      {
        id: 'mock-event-1',
        subject: 'Team Standup',
        start: new Date(now.getTime() + 2 * 60 * 60 * 1000),
        end: new Date(now.getTime() + 2.5 * 60 * 60 * 1000),
        location: 'Conference Room A',
        isAllDay: false,
        organizer: 'You',
        attendees: [],
        importance: 'normal',
        showAs: 'busy',
        isRecurring: true,
      },
      {
        id: 'mock-event-2',
        subject: 'Project Review',
        start: new Date(now.getTime() + 24 * 60 * 60 * 1000),
        end: new Date(now.getTime() + 25 * 60 * 60 * 1000),
        location: 'Virtual - Teams',
        isAllDay: false,
        organizer: 'Manager',
        attendees: [],
        importance: 'high',
        showAs: 'busy',
        isRecurring: false,
      },
    ];
  }

  /**
   * Mark a single email as read
   */
  async markEmailAsRead(emailId: string): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      const restUrl = this.getRestUrl();
      const url = `${restUrl}/me/messages/${emailId}`;

      const response = await fetch(url, {
        method: 'PATCH',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ isRead: true }),
      });

      return response.ok;
    } catch (error) {
      console.error('Error marking email as read:', error);
      return false;
    }
  }

  /**
   * Mark multiple emails as read
   */
  async markEmailsAsRead(emailIds: string[]): Promise<{ success: number; failed: number }> {
    let success = 0;
    let failed = 0;

    for (const id of emailIds) {
      const result = await this.markEmailAsRead(id);
      if (result) success++;
      else failed++;
    }

    return { success, failed };
  }

  /**
   * Get all unread emails
   */
  async getUnreadEmails(count: number = 50): Promise<EmailSummary[]> {
    const token = await this.getAccessToken();
    if (!token) return [];

    try {
      const restUrl = this.getRestUrl();
      const url = `${restUrl}/me/mailfolders/inbox/messages?$filter=isRead eq false&$top=${count}&$select=id,subject,from,receivedDateTime,bodyPreview,isRead,importance,hasAttachments,conversationId&$orderby=receivedDateTime desc`;

      const response = await fetch(url, {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) return [];

      const data = await response.json();
      return data.value.map((msg: any) => ({
        id: msg.id,
        subject: msg.subject || '(No Subject)',
        sender: msg.from?.emailAddress?.name || 'Unknown',
        senderEmail: msg.from?.emailAddress?.address || '',
        receivedDateTime: new Date(msg.receivedDateTime),
        preview: msg.bodyPreview || '',
        isRead: msg.isRead,
        importance: msg.importance?.toLowerCase() || 'normal',
        hasAttachments: msg.hasAttachments,
        conversationId: msg.conversationId,
      }));
    } catch (error) {
      console.error('Error getting unread emails:', error);
      return [];
    }
  }

  /**
   * Mock emails for development/testing
   */
  private getMockEmails(count: number = 10): EmailSummary[] {
    const mockEmails: EmailSummary[] = [];
    const senders = [
      { name: 'John Smith', email: 'john.smith@company.com' },
      { name: 'Sarah Johnson', email: 'sarah.j@client.org' },
      { name: 'Mike Wilson', email: 'mwilson@vendor.net' },
      { name: 'Emily Brown', email: 'emily.brown@partner.com' },
      { name: 'Tech Support', email: 'support@service.com' },
    ];
    const subjects = [
      'Q4 Report Review Required',
      'Meeting Tomorrow at 2 PM',
      'Project Update - Phase 2',
      'Invoice #12345 Attached',
      'Quick Question about the proposal',
      'Follow-up from our call',
      'Action Required: Contract Review',
      'Weekly Status Update',
      'RE: Budget Approval',
      'FW: Important Announcement',
    ];

    for (let i = 0; i < Math.min(count, 10); i++) {
      const sender = senders[i % senders.length];
      mockEmails.push({
        id: `mock-${i}`,
        subject: subjects[i % subjects.length],
        sender: sender.name,
        senderEmail: sender.email,
        receivedDateTime: new Date(Date.now() - i * 3600000),
        preview: `This is a preview of email ${i + 1}. Lorem ipsum dolor sit amet, consectetur adipiscing elit...`,
        isRead: i > 2,
        importance: i === 0 ? 'high' : 'normal',
        hasAttachments: i % 3 === 0,
        conversationId: `conv-${i}`,
      });
    }

    return mockEmails;
  }
}

export const graphService = new GraphService();
