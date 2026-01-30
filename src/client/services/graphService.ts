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
   * Initialize auth state (restores cached sign-in after reload)
   */
  async initializeAuth(): Promise<void> {
    await msalService.initialize();
    this._isSignedIn = msalService.isSignedIn();
    this._isRealDataMode = this._isSignedIn;
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
   * Search emails in inbox - uses Microsoft Graph $search for full-text search
   * Searches across all folders by default (not just inbox)
   */
  async searchEmails(query: string, count: number = 250): Promise<EmailSummary[]> {
    const token = await this.getAccessToken();
    if (!token) return [];

    try {
      const restUrl = this.getRestUrl();
      let nextUrl = `${restUrl}/me/messages?$search="${encodeURIComponent(query)}"&$select=id,subject,from,receivedDateTime,bodyPreview,isRead,importance,hasAttachments,conversationId`;

      const results: EmailSummary[] = [];
      while (nextUrl && results.length < count) {
        const pageSize = Math.min(100, count - results.length);
        const pageUrl = nextUrl.includes('$top=') ? nextUrl : `${nextUrl}&$top=${pageSize}`;

        const response = await fetch(pageUrl, {
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
            'ConsistencyLevel': 'eventual',
          },
        });

        if (!response.ok) return [];

        const data = await response.json();
        const page = data.value.map((msg: any) => ({
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

        results.push(...page);
        nextUrl = data['@odata.nextLink'] || '';
      }

      return results;
    } catch (error) {
      console.error('Error searching emails:', error);
      return [];
    }
  }

  /**
   * Get emails from a specific sender
   */
  async getEmailsFromSender(senderEmail: string, count: number = 100): Promise<EmailSummary[]> {
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
   * Advanced email search with filters and pagination
   * This provides more powerful search capabilities using Graph API filters
   */
  async advancedSearchEmails(options: {
    searchText?: string;           // Full-text search in subject, body, sender
    senderEmail?: string;          // Filter by sender email
    senderDomain?: string;         // Filter by sender domain (e.g., "@company.com")
    subjectContains?: string;      // Filter by subject keyword
    hasAttachments?: boolean;      // Filter by attachment presence
    isRead?: boolean;              // Filter by read/unread status
    importance?: 'high' | 'normal' | 'low';
    startDate?: Date;              // Filter emails received after this date
    endDate?: Date;                // Filter emails received before this date
    folderId?: string;             // Search in specific folder
    count?: number;                // Max results (default 500)
    skip?: number;                 // Pagination offset
  }): Promise<{ emails: EmailSummary[]; totalEstimate: number; hasMore: boolean }> {
    const token = await this.getAccessToken();
    if (!token) return { emails: [], totalEstimate: 0, hasMore: false };

    const count = options.count || 500;
    const skip = options.skip || 0;

    try {
      const restUrl = this.getRestUrl();
      
      // Build filter conditions
      const filters: string[] = [];
      
      if (options.senderEmail) {
        filters.push(`from/emailAddress/address eq '${options.senderEmail}'`);
      }
      
      if (options.senderDomain) {
        // Use contains for domain matching
        filters.push(`contains(from/emailAddress/address, '${options.senderDomain}')`);
      }
      
      if (options.subjectContains) {
        filters.push(`contains(subject, '${options.subjectContains}')`);
      }
      
      if (options.hasAttachments !== undefined) {
        filters.push(`hasAttachments eq ${options.hasAttachments}`);
      }
      
      if (options.isRead !== undefined) {
        filters.push(`isRead eq ${options.isRead}`);
      }
      
      if (options.importance) {
        filters.push(`importance eq '${options.importance}'`);
      }
      
      if (options.startDate) {
        filters.push(`receivedDateTime ge ${options.startDate.toISOString()}`);
      }
      
      if (options.endDate) {
        filters.push(`receivedDateTime le ${options.endDate.toISOString()}`);
      }
      
      // Build the URL
      let baseUrl = options.folderId 
        ? `${restUrl}/me/mailfolders/${options.folderId}/messages`
        : `${restUrl}/me/messages`;
      
      const params: string[] = [];
      
      // Add $search if provided (for full-text search)
      if (options.searchText) {
        params.push(`$search="${encodeURIComponent(options.searchText)}"`);
      }
      
      // Add filters
      if (filters.length > 0) {
        params.push(`$filter=${filters.join(' and ')}`);
      }
      
      // Add selection, count, pagination, and order
      params.push(`$select=id,subject,from,receivedDateTime,bodyPreview,isRead,importance,hasAttachments,conversationId,parentFolderId`);
      params.push(`$count=true`);
      params.push(`$top=${count}`);
      if (skip > 0) {
        params.push(`$skip=${skip}`);
      }
      params.push(`$orderby=receivedDateTime desc`);
      
      const url = `${baseUrl}?${params.join('&')}`;
      console.log(`🔍 Advanced search URL: ${url}`);

      const response = await fetch(url, {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
          'ConsistencyLevel': 'eventual',  // Required for $count
        },
      });

      if (!response.ok) {
        console.error('Advanced search failed:', await response.text());
        return { emails: [], totalEstimate: 0, hasMore: false };
      }

      const data = await response.json();
      const totalCount = data['@odata.count'] || data.value?.length || 0;
      
      const emails = data.value.map((msg: any) => ({
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
      
      console.log(`📬 Advanced search found ${emails.length} emails (total estimate: ${totalCount})`);
      
      return {
        emails,
        totalEstimate: totalCount,
        hasMore: skip + emails.length < totalCount,
      };
    } catch (error) {
      console.error('Error in advanced search:', error);
      return { emails: [], totalEstimate: 0, hasMore: false };
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
   * Get all unread emails - searches across all folders
   */
  async getUnreadEmails(count: number = 250): Promise<EmailSummary[]> {
    const token = await this.getAccessToken();
    if (!token) return [];

    try {
      const restUrl = this.getRestUrl();
      // Search all messages (not just inbox) for unread emails
      const url = `${restUrl}/me/messages?$filter=isRead eq false&$top=${count}&$select=id,subject,from,receivedDateTime,bodyPreview,isRead,importance,hasAttachments,conversationId,parentFolderId&$orderby=receivedDateTime desc`;

      const response = await fetch(url, {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) return [];

      const data = await response.json();
      console.log(`📬 Found ${data.value?.length || 0} unread emails across all folders`);
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
   * Send an email via Microsoft Graph API
   */
  async sendEmail(to: string, subject: string, body: string, cc?: string, bcc?: string): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) {
      console.error('Cannot send email: not authenticated');
      return false;
    }

    try {
      const restUrl = this.getRestUrl();
      
      // Build the message object
      const message: any = {
        subject,
        body: {
          contentType: 'HTML',
          content: body.replace(/\n/g, '<br>'),
        },
        toRecipients: to.split(/[,;]/).map(email => ({
          emailAddress: { address: email.trim() }
        })),
      };
      
      if (cc) {
        message.ccRecipients = cc.split(/[,;]/).map(email => ({
          emailAddress: { address: email.trim() }
        }));
      }
      
      if (bcc) {
        message.bccRecipients = bcc.split(/[,;]/).map(email => ({
          emailAddress: { address: email.trim() }
        }));
      }

      const response = await fetch(`${restUrl}/me/sendMail`, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          message,
          saveToSentItems: true,
        }),
      });

      if (!response.ok) {
        const error = await response.text();
        console.error('Send email error:', response.status, error);
        return false;
      }

      console.log('✅ Email sent successfully');
      return true;
    } catch (error) {
      console.error('Error sending email:', error);
      return false;
    }
  }

  /**
   * Reply to an email
   */
  async replyToEmail(emailId: string, body: string): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      const restUrl = this.getRestUrl();
      
      const response = await fetch(`${restUrl}/me/messages/${emailId}/reply`, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          comment: body.replace(/\n/g, '<br>'),
        }),
      });

      return response.ok;
    } catch (error) {
      console.error('Error replying to email:', error);
      return false;
    }
  }

  /**
   * Delete an email
   */
  async deleteEmail(emailId: string): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      const restUrl = this.getRestUrl();
      
      const response = await fetch(`${restUrl}/me/messages/${emailId}`, {
        method: 'DELETE',
        headers: {
          'Authorization': `Bearer ${token}`,
        },
      });

      return response.ok;
    } catch (error) {
      console.error('Error deleting email:', error);
      return false;
    }
  }

  /**
   * Create a calendar event
   */
  async createCalendarEvent(
    subject: string,
    start: Date,
    end: Date,
    attendees?: string[],
    body?: string,
    location?: string
  ): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      const restUrl = this.getRestUrl();
      
      const event: any = {
        subject,
        body: body ? {
          contentType: 'HTML',
          content: body,
        } : undefined,
        start: {
          dateTime: start.toISOString(),
          timeZone: 'UTC',
        },
        end: {
          dateTime: end.toISOString(),
          timeZone: 'UTC',
        },
        location: location ? { displayName: location } : undefined,
      };
      
      if (attendees && attendees.length > 0) {
        event.attendees = attendees.map(email => ({
          emailAddress: { address: email.trim() },
          type: 'required',
        }));
      }

      const response = await fetch(`${restUrl}/me/events`, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(event),
      });

      return response.ok;
    } catch (error) {
      console.error('Error creating calendar event:', error);
      return false;
    }
  }

  /**
   * Forward an email to specified recipients
   */
  async forwardEmail(
    emailId: string,
    toRecipients: string[],
    comment?: string
  ): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      const requestBody: any = {
        toRecipients: toRecipients.map(email => ({
          emailAddress: { address: email }
        }))
      };
      
      if (comment) {
        requestBody.comment = comment;
      }

      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/messages/${emailId}/forward`,
        {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(requestBody),
        }
      );

      if (response.ok || response.status === 202) {
        console.log('Email forwarded successfully');
        return true;
      } else {
        console.error('Failed to forward email:', response.status);
        return false;
      }
    } catch (error) {
      console.error('Error forwarding email:', error);
      return false;
    }
  }

  /**
   * Get all mail folders
   */
  async getFolders(): Promise<{ id: string; displayName: string; unreadItemCount: number }[]> {
    const token = await this.getAccessToken();
    if (!token) return [];

    try {
      const response = await fetch(
        'https://graph.microsoft.com/v1.0/me/mailFolders?$top=50',
        {
          headers: { 'Authorization': `Bearer ${token}` },
        }
      );

      if (response.ok) {
        const data = await response.json();
        return data.value.map((folder: any) => ({
          id: folder.id,
          displayName: folder.displayName,
          unreadItemCount: folder.unreadItemCount || 0,
        }));
      }
      return [];
    } catch (error) {
      console.error('Error getting folders:', error);
      return [];
    }
  }

  /**
   * Move an email to a specific folder
   */
  async moveEmailToFolder(emailId: string, folderId: string): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/messages/${emailId}/move`,
        {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({ destinationId: folderId }),
        }
      );

      if (response.ok) {
        console.log('Email moved successfully');
        return true;
      } else {
        console.error('Failed to move email:', response.status);
        return false;
      }
    } catch (error) {
      console.error('Error moving email:', error);
      return false;
    }
  }

  /**
   * Archive an email (move to Archive folder)
   */
  async archiveEmail(emailId: string): Promise<boolean> {
    const folders = await this.getFolders();
    const archiveFolder = folders.find(f => 
      f.displayName.toLowerCase() === 'archive'
    );
    
    if (archiveFolder) {
      return this.moveEmailToFolder(emailId, archiveFolder.id);
    } else {
      console.error('Archive folder not found');
      return false;
    }
  }

  /**
   * Flag or unflag an email
   */
  async flagEmail(emailId: string, flagged: boolean): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/messages/${emailId}`,
        {
          method: 'PATCH',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            flag: {
              flagStatus: flagged ? 'flagged' : 'notFlagged'
            }
          }),
        }
      );

      if (response.ok) {
        console.log(`Email ${flagged ? 'flagged' : 'unflagged'} successfully`);
        return true;
      } else {
        console.error('Failed to flag email:', response.status);
        return false;
      }
    } catch (error) {
      console.error('Error flagging email:', error);
      return false;
    }
  }

  /**
   * Create a draft email
   */
  async createDraft(
    to: string[],
    subject: string,
    body: string,
    cc?: string[],
    bcc?: string[]
  ): Promise<{ id: string; webLink: string } | null> {
    const token = await this.getAccessToken();
    if (!token) return null;

    try {
      const emailMessage: any = {
        subject,
        body: {
          contentType: 'HTML',
          content: body.replace(/\n/g, '<br>'),
        },
        toRecipients: to.map(email => ({
          emailAddress: { address: email }
        })),
      };

      if (cc && cc.length > 0) {
        emailMessage.ccRecipients = cc.map(email => ({
          emailAddress: { address: email }
        }));
      }

      if (bcc && bcc.length > 0) {
        emailMessage.bccRecipients = bcc.map(email => ({
          emailAddress: { address: email }
        }));
      }

      const response = await fetch(
        'https://graph.microsoft.com/v1.0/me/messages',
        {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(emailMessage),
        }
      );

      if (response.ok) {
        const draft = await response.json();
        console.log('Draft created successfully');
        return {
          id: draft.id,
          webLink: draft.webLink || '',
        };
      } else {
        console.error('Failed to create draft:', response.status);
        return null;
      }
    } catch (error) {
      console.error('Error creating draft:', error);
      return null;
    }
  }

  /**
   * Mark an email as unread
   */
  async markEmailAsUnread(emailId: string): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/messages/${emailId}`,
        {
          method: 'PATCH',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({ isRead: false }),
        }
      );

      if (response.ok) {
        console.log('Email marked as unread');
        return true;
      }
      return false;
    } catch (error) {
      console.error('Error marking email as unread:', error);
      return false;
    }
  }

  /**
   * Create a To-Do task
   */
  async createTask(
    title: string,
    dueDate?: Date,
    body?: string,
    linkedEmailId?: string
  ): Promise<{ id: string } | null> {
    const token = await this.getAccessToken();
    if (!token) return null;

    try {
      // First get the default task list
      const listsResponse = await fetch(
        'https://graph.microsoft.com/v1.0/me/todo/lists',
        {
          headers: { 'Authorization': `Bearer ${token}` },
        }
      );

      if (!listsResponse.ok) {
        console.error('Failed to get task lists');
        return null;
      }

      const listsData = await listsResponse.json();
      const defaultList = listsData.value.find((list: any) => 
        list.wellknownListName === 'defaultList' || list.displayName === 'Tasks'
      ) || listsData.value[0];

      if (!defaultList) {
        console.error('No task list found');
        return null;
      }

      const taskBody: any = {
        title,
        importance: 'normal',
      };

      if (dueDate) {
        taskBody.dueDateTime = {
          dateTime: dueDate.toISOString(),
          timeZone: 'UTC',
        };
      }

      if (body) {
        taskBody.body = {
          content: body,
          contentType: 'text',
        };
      }

      if (linkedEmailId) {
        taskBody.linkedResources = [{
          webUrl: `https://outlook.office.com/mail/id/${linkedEmailId}`,
          applicationName: 'Outlook',
          displayName: 'Related Email',
        }];
      }

      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/todo/lists/${defaultList.id}/tasks`,
        {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(taskBody),
        }
      );

      if (response.ok) {
        const task = await response.json();
        console.log('Task created successfully');
        return { id: task.id };
      } else {
        console.error('Failed to create task:', response.status);
        return null;
      }
    } catch (error) {
      console.error('Error creating task:', error);
      return null;
    }
  }

  /**
   * Get upcoming tasks
   */
  async getTasks(count: number = 20): Promise<{
    id: string;
    title: string;
    status: string;
    dueDate?: Date;
    importance: string;
  }[]> {
    const token = await this.getAccessToken();
    if (!token) return [];

    try {
      // Get all task lists
      const listsResponse = await fetch(
        'https://graph.microsoft.com/v1.0/me/todo/lists',
        {
          headers: { 'Authorization': `Bearer ${token}` },
        }
      );

      if (!listsResponse.ok) return [];

      const listsData = await listsResponse.json();
      const allTasks: any[] = [];

      // Get tasks from each list
      for (const list of listsData.value) {
        const tasksResponse = await fetch(
          `https://graph.microsoft.com/v1.0/me/todo/lists/${list.id}/tasks?$filter=status ne 'completed'&$top=${count}`,
          {
            headers: { 'Authorization': `Bearer ${token}` },
          }
        );

        if (tasksResponse.ok) {
          const tasksData = await tasksResponse.json();
          allTasks.push(...tasksData.value);
        }
      }

      return allTasks.slice(0, count).map((task: any) => ({
        id: task.id,
        title: task.title,
        status: task.status,
        dueDate: task.dueDateTime ? new Date(task.dueDateTime.dateTime) : undefined,
        importance: task.importance,
      }));
    } catch (error) {
      console.error('Error getting tasks:', error);
      return [];
    }
  }

  /**
   * Delete a calendar event
   */
  async deleteCalendarEvent(eventId: string): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/calendar/events/${eventId}`,
        {
          method: 'DELETE',
          headers: { 'Authorization': `Bearer ${token}` },
        }
      );

      if (response.ok || response.status === 204) {
        console.log('Calendar event deleted successfully');
        return true;
      } else {
        console.error('Failed to delete calendar event:', response.status);
        return false;
      }
    } catch (error) {
      console.error('Error deleting calendar event:', error);
      return false;
    }
  }

  /**
   * Get contacts from address book
   */
  async getContacts(count: number = 50): Promise<{
    id: string;
    displayName: string;
    emailAddresses: string[];
    jobTitle?: string;
    companyName?: string;
  }[]> {
    const token = await this.getAccessToken();
    if (!token) return [];

    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/contacts?$top=${count}&$select=id,displayName,emailAddresses,jobTitle,companyName`,
        {
          headers: { 'Authorization': `Bearer ${token}` },
        }
      );

      if (response.ok) {
        const data = await response.json();
        return data.value.map((contact: any) => ({
          id: contact.id,
          displayName: contact.displayName || '',
          emailAddresses: (contact.emailAddresses || []).map((e: any) => e.address),
          jobTitle: contact.jobTitle,
          companyName: contact.companyName,
        }));
      }
      return [];
    } catch (error) {
      console.error('Error getting contacts:', error);
      return [];
    }
  }

  /**
   * Search contacts by name or email
   */
  async searchContacts(query: string): Promise<{
    displayName: string;
    emailAddress: string;
    source: string;
  }[]> {
    const token = await this.getAccessToken();
    if (!token) return [];

    try {
      // Use the people API for better results
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/people?$search="${encodeURIComponent(query)}"&$top=10`,
        {
          headers: { 'Authorization': `Bearer ${token}` },
        }
      );

      if (response.ok) {
        const data = await response.json();
        return data.value
          .filter((person: any) => person.scoredEmailAddresses?.length > 0)
          .map((person: any) => ({
            displayName: person.displayName || '',
            emailAddress: person.scoredEmailAddresses[0]?.address || '',
            source: person.personType?.class || 'Unknown',
          }));
      }
      return [];
    } catch (error) {
      console.error('Error searching contacts:', error);
      return [];
    }
  }

  // ============================================================
  // ATTACHMENT FUNCTIONS
  // ============================================================

  /**
   * Get attachments for an email
   */
  async getAttachments(emailId: string): Promise<{
    id: string;
    name: string;
    contentType: string;
    size: number;
    isInline: boolean;
  }[]> {
    const token = await this.getAccessToken();
    if (!token) return [];

    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/messages/${emailId}/attachments`,
        {
          headers: { 'Authorization': `Bearer ${token}` },
        }
      );

      if (response.ok) {
        const data = await response.json();
        return data.value.map((att: any) => ({
          id: att.id,
          name: att.name,
          contentType: att.contentType,
          size: att.size,
          isInline: att.isInline || false,
        }));
      }
      return [];
    } catch (error) {
      console.error('Error getting attachments:', error);
      return [];
    }
  }

  /**
   * Download attachment content (base64)
   */
  async downloadAttachment(emailId: string, attachmentId: string): Promise<{ name: string; content: string } | null> {
    const token = await this.getAccessToken();
    if (!token) return null;

    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/messages/${emailId}/attachments/${attachmentId}`,
        {
          headers: { 'Authorization': `Bearer ${token}` },
        }
      );

      if (response.ok) {
        const data = await response.json();
        return {
          name: data.name,
          content: data.contentBytes || '',
        };
      }
      return null;
    } catch (error) {
      console.error('Error downloading attachment:', error);
      return null;
    }
  }

  /**
   * Add attachment to a draft email
   */
  async addAttachment(
    emailId: string,
    name: string,
    contentBytes: string,
    contentType: string = 'application/octet-stream'
  ): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/messages/${emailId}/attachments`,
        {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            '@odata.type': '#microsoft.graph.fileAttachment',
            name,
            contentBytes,
            contentType,
          }),
        }
      );

      return response.ok;
    } catch (error) {
      console.error('Error adding attachment:', error);
      return false;
    }
  }

  // ============================================================
  // EMAIL CATEGORY & IMPORTANCE FUNCTIONS
  // ============================================================

  /**
   * Set categories on an email
   */
  async setEmailCategories(emailId: string, categories: string[]): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/messages/${emailId}`,
        {
          method: 'PATCH',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({ categories }),
        }
      );

      return response.ok;
    } catch (error) {
      console.error('Error setting categories:', error);
      return false;
    }
  }

  /**
   * Get available categories
   */
  async getCategories(): Promise<{ displayName: string; color: string }[]> {
    const token = await this.getAccessToken();
    if (!token) return [];

    try {
      const response = await fetch(
        'https://graph.microsoft.com/v1.0/me/outlook/masterCategories',
        {
          headers: { 'Authorization': `Bearer ${token}` },
        }
      );

      if (response.ok) {
        const data = await response.json();
        return data.value.map((cat: any) => ({
          displayName: cat.displayName,
          color: cat.color,
        }));
      }
      return [];
    } catch (error) {
      console.error('Error getting categories:', error);
      return [];
    }
  }

  /**
   * Set email importance
   */
  async setEmailImportance(emailId: string, importance: 'low' | 'normal' | 'high'): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/messages/${emailId}`,
        {
          method: 'PATCH',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({ importance }),
        }
      );

      return response.ok;
    } catch (error) {
      console.error('Error setting importance:', error);
      return false;
    }
  }

  // ============================================================
  // REPLY ALL & CONVERSATION THREAD
  // ============================================================

  /**
   * Reply all to an email
   */
  async replyAllToEmail(emailId: string, body: string): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/messages/${emailId}/replyAll`,
        {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            comment: body.replace(/\n/g, '<br>'),
          }),
        }
      );

      return response.ok || response.status === 202;
    } catch (error) {
      console.error('Error replying all:', error);
      return false;
    }
  }

  /**
   * Get conversation thread
   */
  async getConversationThread(conversationId: string, count: number = 20): Promise<EmailSummary[]> {
    const token = await this.getAccessToken();
    if (!token) return [];

    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/messages?$filter=conversationId eq '${conversationId}'&$top=${count}&$orderby=receivedDateTime desc`,
        {
          headers: { 'Authorization': `Bearer ${token}` },
        }
      );

      if (response.ok) {
        const data = await response.json();
        return data.value.map((email: any) => ({
          id: email.id,
          subject: email.subject,
          sender: email.from?.emailAddress?.name || 'Unknown',
          senderEmail: email.from?.emailAddress?.address || '',
          receivedDateTime: new Date(email.receivedDateTime),
          preview: email.bodyPreview || '',
          isRead: email.isRead,
          importance: email.importance,
          hasAttachments: email.hasAttachments,
          conversationId: email.conversationId,
        }));
      }
      return [];
    } catch (error) {
      console.error('Error getting conversation thread:', error);
      return [];
    }
  }

  // ============================================================
  // CALENDAR - MEETING RESPONSES
  // ============================================================

  /**
   * Accept a meeting invitation
   */
  async acceptMeeting(eventId: string, comment?: string, sendResponse: boolean = true): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/events/${eventId}/accept`,
        {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            comment: comment || '',
            sendResponse,
          }),
        }
      );

      return response.ok || response.status === 202;
    } catch (error) {
      console.error('Error accepting meeting:', error);
      return false;
    }
  }

  /**
   * Decline a meeting invitation
   */
  async declineMeeting(eventId: string, comment?: string, sendResponse: boolean = true): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/events/${eventId}/decline`,
        {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            comment: comment || '',
            sendResponse,
          }),
        }
      );

      return response.ok || response.status === 202;
    } catch (error) {
      console.error('Error declining meeting:', error);
      return false;
    }
  }

  /**
   * Tentatively accept a meeting
   */
  async tentativelyAcceptMeeting(eventId: string, comment?: string, sendResponse: boolean = true): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/events/${eventId}/tentativelyAccept`,
        {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            comment: comment || '',
            sendResponse,
          }),
        }
      );

      return response.ok || response.status === 202;
    } catch (error) {
      console.error('Error tentatively accepting meeting:', error);
      return false;
    }
  }

  /**
   * Update a calendar event
   */
  async updateCalendarEvent(
    eventId: string,
    updates: {
      subject?: string;
      start?: Date;
      end?: Date;
      location?: string;
      body?: string;
      attendees?: string[];
    }
  ): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      const eventUpdate: any = {};
      
      if (updates.subject) eventUpdate.subject = updates.subject;
      if (updates.location) eventUpdate.location = { displayName: updates.location };
      if (updates.body) eventUpdate.body = { contentType: 'HTML', content: updates.body };
      
      if (updates.start) {
        eventUpdate.start = {
          dateTime: updates.start.toISOString(),
          timeZone: 'UTC',
        };
      }
      
      if (updates.end) {
        eventUpdate.end = {
          dateTime: updates.end.toISOString(),
          timeZone: 'UTC',
        };
      }
      
      if (updates.attendees) {
        eventUpdate.attendees = updates.attendees.map(email => ({
          emailAddress: { address: email },
          type: 'required',
        }));
      }

      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/events/${eventId}`,
        {
          method: 'PATCH',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(eventUpdate),
        }
      );

      return response.ok;
    } catch (error) {
      console.error('Error updating calendar event:', error);
      return false;
    }
  }

  /**
   * Create a recurring calendar event
   */
  async createRecurringEvent(
    subject: string,
    start: Date,
    end: Date,
    recurrence: {
      pattern: 'daily' | 'weekly' | 'monthly' | 'yearly';
      interval: number;
      daysOfWeek?: ('sunday' | 'monday' | 'tuesday' | 'wednesday' | 'thursday' | 'friday' | 'saturday')[];
      endDate?: Date;
      occurrences?: number;
    },
    attendees?: string[],
    location?: string
  ): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      const recurrencePattern: any = {
        type: recurrence.pattern,
        interval: recurrence.interval,
      };
      
      if (recurrence.daysOfWeek) {
        recurrencePattern.daysOfWeek = recurrence.daysOfWeek;
      }

      const recurrenceRange: any = {
        type: recurrence.endDate ? 'endDate' : (recurrence.occurrences ? 'numbered' : 'noEnd'),
        startDate: start.toISOString().split('T')[0],
      };
      
      if (recurrence.endDate) {
        recurrenceRange.endDate = recurrence.endDate.toISOString().split('T')[0];
      }
      if (recurrence.occurrences) {
        recurrenceRange.numberOfOccurrences = recurrence.occurrences;
      }

      const event: any = {
        subject,
        start: { dateTime: start.toISOString(), timeZone: 'UTC' },
        end: { dateTime: end.toISOString(), timeZone: 'UTC' },
        recurrence: {
          pattern: recurrencePattern,
          range: recurrenceRange,
        },
      };

      if (attendees && attendees.length > 0) {
        event.attendees = attendees.map(email => ({
          emailAddress: { address: email },
          type: 'required',
        }));
      }

      if (location) {
        event.location = { displayName: location };
      }

      const response = await fetch(
        'https://graph.microsoft.com/v1.0/me/events',
        {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(event),
        }
      );

      return response.ok;
    } catch (error) {
      console.error('Error creating recurring event:', error);
      return false;
    }
  }

  /**
   * Get free/busy schedule
   */
  async getFreeBusy(
    emails: string[],
    start: Date,
    end: Date
  ): Promise<{
    email: string;
    availability: 'free' | 'tentative' | 'busy' | 'oof' | 'workingElsewhere' | 'unknown';
    slots: { start: Date; end: Date; status: string }[];
  }[]> {
    const token = await this.getAccessToken();
    if (!token) return [];

    try {
      const response = await fetch(
        'https://graph.microsoft.com/v1.0/me/calendar/getSchedule',
        {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            schedules: emails,
            startTime: { dateTime: start.toISOString(), timeZone: 'UTC' },
            endTime: { dateTime: end.toISOString(), timeZone: 'UTC' },
            availabilityViewInterval: 30,
          }),
        }
      );

      if (response.ok) {
        const data = await response.json();
        return data.value.map((schedule: any) => ({
          email: schedule.scheduleId,
          availability: schedule.availabilityView?.[0] === '0' ? 'free' : 'busy',
          slots: (schedule.scheduleItems || []).map((item: any) => ({
            start: new Date(item.start.dateTime),
            end: new Date(item.end.dateTime),
            status: item.status,
          })),
        }));
      }
      return [];
    } catch (error) {
      console.error('Error getting free/busy:', error);
      return [];
    }
  }

  // ============================================================
  // TASKS - COMPLETE, DELETE, UPDATE
  // ============================================================

  /**
   * Complete a task
   */
  async completeTask(taskId: string, listId?: string): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      // If no listId provided, find it
      let targetListId = listId;
      if (!targetListId) {
        const listsResponse = await fetch(
          'https://graph.microsoft.com/v1.0/me/todo/lists',
          { headers: { 'Authorization': `Bearer ${token}` } }
        );
        if (listsResponse.ok) {
          const listsData = await listsResponse.json();
          // Search all lists for the task
          for (const list of listsData.value) {
            const taskResponse = await fetch(
              `https://graph.microsoft.com/v1.0/me/todo/lists/${list.id}/tasks/${taskId}`,
              { headers: { 'Authorization': `Bearer ${token}` } }
            );
            if (taskResponse.ok) {
              targetListId = list.id;
              break;
            }
          }
        }
      }

      if (!targetListId) return false;

      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/todo/lists/${targetListId}/tasks/${taskId}`,
        {
          method: 'PATCH',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            status: 'completed',
            completedDateTime: {
              dateTime: new Date().toISOString(),
              timeZone: 'UTC',
            },
          }),
        }
      );

      return response.ok;
    } catch (error) {
      console.error('Error completing task:', error);
      return false;
    }
  }

  /**
   * Delete a task
   */
  async deleteTask(taskId: string, listId?: string): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      let targetListId = listId;
      if (!targetListId) {
        const listsResponse = await fetch(
          'https://graph.microsoft.com/v1.0/me/todo/lists',
          { headers: { 'Authorization': `Bearer ${token}` } }
        );
        if (listsResponse.ok) {
          const listsData = await listsResponse.json();
          for (const list of listsData.value) {
            const taskResponse = await fetch(
              `https://graph.microsoft.com/v1.0/me/todo/lists/${list.id}/tasks/${taskId}`,
              { headers: { 'Authorization': `Bearer ${token}` } }
            );
            if (taskResponse.ok) {
              targetListId = list.id;
              break;
            }
          }
        }
      }

      if (!targetListId) return false;

      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/todo/lists/${targetListId}/tasks/${taskId}`,
        {
          method: 'DELETE',
          headers: { 'Authorization': `Bearer ${token}` },
        }
      );

      return response.ok || response.status === 204;
    } catch (error) {
      console.error('Error deleting task:', error);
      return false;
    }
  }

  /**
   * Update a task
   */
  async updateTask(
    taskId: string,
    updates: {
      title?: string;
      body?: string;
      dueDate?: Date;
      importance?: 'low' | 'normal' | 'high';
    },
    listId?: string
  ): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      let targetListId = listId;
      if (!targetListId) {
        const listsResponse = await fetch(
          'https://graph.microsoft.com/v1.0/me/todo/lists',
          { headers: { 'Authorization': `Bearer ${token}` } }
        );
        if (listsResponse.ok) {
          const listsData = await listsResponse.json();
          for (const list of listsData.value) {
            const taskResponse = await fetch(
              `https://graph.microsoft.com/v1.0/me/todo/lists/${list.id}/tasks/${taskId}`,
              { headers: { 'Authorization': `Bearer ${token}` } }
            );
            if (taskResponse.ok) {
              targetListId = list.id;
              break;
            }
          }
        }
      }

      if (!targetListId) return false;

      const taskUpdate: any = {};
      if (updates.title) taskUpdate.title = updates.title;
      if (updates.body) taskUpdate.body = { content: updates.body, contentType: 'text' };
      if (updates.importance) taskUpdate.importance = updates.importance;
      if (updates.dueDate) {
        taskUpdate.dueDateTime = {
          dateTime: updates.dueDate.toISOString(),
          timeZone: 'UTC',
        };
      }

      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/todo/lists/${targetListId}/tasks/${taskId}`,
        {
          method: 'PATCH',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(taskUpdate),
        }
      );

      return response.ok;
    } catch (error) {
      console.error('Error updating task:', error);
      return false;
    }
  }

  // ============================================================
  // CONTACTS - CREATE, UPDATE, DELETE
  // ============================================================

  /**
   * Create a new contact
   */
  async createContact(contact: {
    givenName?: string;
    surname?: string;
    displayName?: string;
    emailAddresses?: { address: string; name?: string }[];
    businessPhones?: string[];
    mobilePhone?: string;
    companyName?: string;
    jobTitle?: string;
  }): Promise<{ id: string } | null> {
    const token = await this.getAccessToken();
    if (!token) return null;

    try {
      const response = await fetch(
        'https://graph.microsoft.com/v1.0/me/contacts',
        {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(contact),
        }
      );

      if (response.ok) {
        const data = await response.json();
        return { id: data.id };
      }
      return null;
    } catch (error) {
      console.error('Error creating contact:', error);
      return null;
    }
  }

  /**
   * Update a contact
   */
  async updateContact(contactId: string, updates: {
    givenName?: string;
    surname?: string;
    displayName?: string;
    emailAddresses?: { address: string; name?: string }[];
    businessPhones?: string[];
    mobilePhone?: string;
    companyName?: string;
    jobTitle?: string;
  }): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/contacts/${contactId}`,
        {
          method: 'PATCH',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(updates),
        }
      );

      return response.ok;
    } catch (error) {
      console.error('Error updating contact:', error);
      return false;
    }
  }

  /**
   * Delete a contact
   */
  async deleteContact(contactId: string): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/contacts/${contactId}`,
        {
          method: 'DELETE',
          headers: { 'Authorization': `Bearer ${token}` },
        }
      );

      return response.ok || response.status === 204;
    } catch (error) {
      console.error('Error deleting contact:', error);
      return false;
    }
  }

  // ============================================================
  // FOLDERS - CREATE, RENAME, DELETE
  // ============================================================

  /**
   * Create a new mail folder
   */
  async createFolder(displayName: string, parentFolderId?: string): Promise<{ id: string } | null> {
    const token = await this.getAccessToken();
    if (!token) return null;

    try {
      const url = parentFolderId
        ? `https://graph.microsoft.com/v1.0/me/mailFolders/${parentFolderId}/childFolders`
        : 'https://graph.microsoft.com/v1.0/me/mailFolders';

      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ displayName }),
      });

      if (response.ok) {
        const data = await response.json();
        return { id: data.id };
      }
      if (response.status === 409) {
        const folders = await this.getFolders();
        const existing = folders.find(
          (folder: any) => folder.displayName?.toLowerCase() === displayName.toLowerCase()
        );
        return existing ? { id: existing.id } : null;
      }
      return null;
    } catch (error) {
      console.error('Error creating folder:', error);
      return null;
    }
  }

  /**
   * Rename a mail folder
   */
  async renameFolder(folderId: string, newName: string): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/mailFolders/${folderId}`,
        {
          method: 'PATCH',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({ displayName: newName }),
        }
      );

      return response.ok;
    } catch (error) {
      console.error('Error renaming folder:', error);
      return false;
    }
  }

  /**
   * Delete a mail folder
   */
  async deleteFolder(folderId: string): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/mailFolders/${folderId}`,
        {
          method: 'DELETE',
          headers: { 'Authorization': `Bearer ${token}` },
        }
      );

      return response.ok || response.status === 204;
    } catch (error) {
      console.error('Error deleting folder:', error);
      return false;
    }
  }

  // ============================================================
  // MAIL RULES & AUTO-REPLY (OUT OF OFFICE)
  // ============================================================

  /**
   * Get mail rules
   */
  async getMailRules(): Promise<{
    id: string;
    displayName: string;
    isEnabled: boolean;
    conditions: any;
    actions: any;
  }[]> {
    const token = await this.getAccessToken();
    if (!token) return [];

    try {
      const response = await fetch(
        'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules',
        {
          headers: { 'Authorization': `Bearer ${token}` },
        }
      );

      if (response.ok) {
        const data = await response.json();
        return data.value.map((rule: any) => ({
          id: rule.id,
          displayName: rule.displayName,
          isEnabled: rule.isEnabled,
          conditions: rule.conditions,
          actions: rule.actions,
        }));
      }
      return [];
    } catch (error) {
      console.error('Error getting mail rules:', error);
      return [];
    }
  }

  /**
   * Create a mail rule
   */
  async createMailRule(rule: {
    displayName: string;
    conditions: {
      fromAddresses?: { address: string }[];
      subjectContains?: string[];
      senderContains?: string[];
    };
    actions: {
      moveToFolder?: string;
      delete?: boolean;
      markAsRead?: boolean;
      forwardTo?: { address: string }[];
    };
  }): Promise<{ id: string } | null> {
    const token = await this.getAccessToken();
    if (!token) return null;

    try {
      const response = await fetch(
        'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules',
        {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            displayName: rule.displayName,
            sequence: 1,
            isEnabled: true,
            conditions: rule.conditions,
            actions: rule.actions,
          }),
        }
      );

      if (response.ok) {
        const data = await response.json();
        return { id: data.id };
      }
      return null;
    } catch (error) {
      console.error('Error creating mail rule:', error);
      return null;
    }
  }

  /**
   * Delete a mail rule
   */
  async deleteMailRule(ruleId: string): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules/${ruleId}`,
        {
          method: 'DELETE',
          headers: { 'Authorization': `Bearer ${token}` },
        }
      );

      return response.ok || response.status === 204;
    } catch (error) {
      console.error('Error deleting mail rule:', error);
      return false;
    }
  }

  /**
   * Get automatic replies (Out of Office) settings
   */
  async getAutoReply(): Promise<{
    status: 'disabled' | 'alwaysEnabled' | 'scheduled';
    externalMessage?: string;
    internalMessage?: string;
    scheduledStartDateTime?: Date;
    scheduledEndDateTime?: Date;
  } | null> {
    const token = await this.getAccessToken();
    if (!token) return null;

    try {
      const response = await fetch(
        'https://graph.microsoft.com/v1.0/me/mailboxSettings/automaticRepliesSetting',
        {
          headers: { 'Authorization': `Bearer ${token}` },
        }
      );

      if (response.ok) {
        const data = await response.json();
        return {
          status: data.status,
          externalMessage: data.externalReplyMessage,
          internalMessage: data.internalReplyMessage,
          scheduledStartDateTime: data.scheduledStartDateTime?.dateTime 
            ? new Date(data.scheduledStartDateTime.dateTime) : undefined,
          scheduledEndDateTime: data.scheduledEndDateTime?.dateTime 
            ? new Date(data.scheduledEndDateTime.dateTime) : undefined,
        };
      }
      return null;
    } catch (error) {
      console.error('Error getting auto-reply settings:', error);
      return null;
    }
  }

  /**
   * Set automatic replies (Out of Office)
   */
  async setAutoReply(settings: {
    status: 'disabled' | 'alwaysEnabled' | 'scheduled';
    externalMessage?: string;
    internalMessage?: string;
    scheduledStartDateTime?: Date;
    scheduledEndDateTime?: Date;
    externalAudience?: 'none' | 'contactsOnly' | 'all';
  }): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      const autoReplySettings: any = {
        status: settings.status,
      };

      if (settings.externalMessage) {
        autoReplySettings.externalReplyMessage = settings.externalMessage;
      }
      if (settings.internalMessage) {
        autoReplySettings.internalReplyMessage = settings.internalMessage;
      }
      if (settings.scheduledStartDateTime) {
        autoReplySettings.scheduledStartDateTime = {
          dateTime: settings.scheduledStartDateTime.toISOString(),
          timeZone: 'UTC',
        };
      }
      if (settings.scheduledEndDateTime) {
        autoReplySettings.scheduledEndDateTime = {
          dateTime: settings.scheduledEndDateTime.toISOString(),
          timeZone: 'UTC',
        };
      }
      if (settings.externalAudience) {
        autoReplySettings.externalAudience = settings.externalAudience;
      }

      const response = await fetch(
        'https://graph.microsoft.com/v1.0/me/mailboxSettings',
        {
          method: 'PATCH',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            automaticRepliesSetting: autoReplySettings,
          }),
        }
      );

      return response.ok;
    } catch (error) {
      console.error('Error setting auto-reply:', error);
      return false;
    }
  }

  // ============================================================
  // SENT ITEMS & DRAFTS
  // ============================================================

  /**
   * Get sent items
   */
  async getSentItems(count: number = 20): Promise<EmailSummary[]> {
    const token = await this.getAccessToken();
    if (!token) return [];

    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/mailFolders/sentitems/messages?$top=${count}&$orderby=sentDateTime desc`,
        {
          headers: { 'Authorization': `Bearer ${token}` },
        }
      );

      if (response.ok) {
        const data = await response.json();
        return data.value.map((email: any) => ({
          id: email.id,
          subject: email.subject,
          sender: email.from?.emailAddress?.name || 'You',
          senderEmail: email.from?.emailAddress?.address || '',
          receivedDateTime: new Date(email.sentDateTime),
          preview: email.bodyPreview || '',
          isRead: true,
          importance: email.importance,
          hasAttachments: email.hasAttachments,
          conversationId: email.conversationId,
        }));
      }
      return [];
    } catch (error) {
      console.error('Error getting sent items:', error);
      return [];
    }
  }

  /**
   * Get draft emails
   */
  async getDrafts(count: number = 20): Promise<EmailSummary[]> {
    const token = await this.getAccessToken();
    if (!token) return [];

    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/mailFolders/drafts/messages?$top=${count}&$orderby=lastModifiedDateTime desc`,
        {
          headers: { 'Authorization': `Bearer ${token}` },
        }
      );

      if (response.ok) {
        const data = await response.json();
        return data.value.map((email: any) => ({
          id: email.id,
          subject: email.subject || '(No subject)',
          sender: 'Draft',
          senderEmail: '',
          receivedDateTime: new Date(email.lastModifiedDateTime),
          preview: email.bodyPreview || '',
          isRead: true,
          importance: email.importance,
          hasAttachments: email.hasAttachments,
          conversationId: email.conversationId,
        }));
      }
      return [];
    } catch (error) {
      console.error('Error getting drafts:', error);
      return [];
    }
  }

  /**
   * Send a draft email
   */
  async sendDraft(draftId: string): Promise<boolean> {
    const token = await this.getAccessToken();
    if (!token) return false;

    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/messages/${draftId}/send`,
        {
          method: 'POST',
          headers: { 'Authorization': `Bearer ${token}` },
        }
      );

      return response.ok || response.status === 202;
    } catch (error) {
      console.error('Error sending draft:', error);
      return false;
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
