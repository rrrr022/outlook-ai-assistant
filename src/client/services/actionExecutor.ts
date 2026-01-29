/**
 * Action Executor - Bridges AI intent to Outlook operations
 * This module parses AI responses for action commands and executes them
 */

import { graphService } from './graphService';

export interface ActionResult {
  success: boolean;
  action: string;
  message: string;
  data?: any;
}

export interface ParsedAction {
  type: string;
  params: Record<string, any>;
}

/**
 * Available actions the AI can trigger
 */
export const AVAILABLE_ACTIONS = {
  // Email actions
  SEND_EMAIL: 'send_email',
  REPLY_EMAIL: 'reply_email',
  REPLY_ALL: 'reply_all',
  FORWARD_EMAIL: 'forward_email',
  DELETE_EMAIL: 'delete_email',
  ARCHIVE_EMAIL: 'archive_email',
  MARK_READ: 'mark_read',
  MARK_UNREAD: 'mark_unread',
  FLAG_EMAIL: 'flag_email',
  UNFLAG_EMAIL: 'unflag_email',
  MOVE_EMAIL: 'move_email',
  CREATE_DRAFT: 'create_draft',
  SEND_DRAFT: 'send_draft',
  SET_CATEGORIES: 'set_categories',
  SET_IMPORTANCE: 'set_importance',
  GET_ATTACHMENTS: 'get_attachments',
  GET_CONVERSATION: 'get_conversation',
  GET_SENT_ITEMS: 'get_sent_items',
  GET_DRAFTS: 'get_drafts',
  
  // Search/Query actions
  SEARCH_EMAILS: 'search_emails',
  GET_UNREAD: 'get_unread',
  GET_EMAIL_DETAILS: 'get_email_details',
  GET_EMAILS_FROM_SENDER: 'get_emails_from_sender',
  
  // Calendar actions
  CREATE_EVENT: 'create_event',
  CREATE_RECURRING_EVENT: 'create_recurring_event',
  UPDATE_EVENT: 'update_event',
  DELETE_EVENT: 'delete_event',
  ACCEPT_MEETING: 'accept_meeting',
  DECLINE_MEETING: 'decline_meeting',
  TENTATIVE_MEETING: 'tentative_meeting',
  GET_CALENDAR: 'get_calendar',
  GET_FREE_BUSY: 'get_free_busy',
  
  // Task actions
  CREATE_TASK: 'create_task',
  COMPLETE_TASK: 'complete_task',
  DELETE_TASK: 'delete_task',
  UPDATE_TASK: 'update_task',
  GET_TASKS: 'get_tasks',
  
  // Contact actions
  GET_CONTACTS: 'get_contacts',
  SEARCH_CONTACTS: 'search_contacts',
  CREATE_CONTACT: 'create_contact',
  UPDATE_CONTACT: 'update_contact',
  DELETE_CONTACT: 'delete_contact',
  
  // Folder actions
  GET_FOLDERS: 'get_folders',
  CREATE_FOLDER: 'create_folder',
  RENAME_FOLDER: 'rename_folder',
  DELETE_FOLDER: 'delete_folder',
  
  // Rules & Settings
  GET_MAIL_RULES: 'get_mail_rules',
  CREATE_MAIL_RULE: 'create_mail_rule',
  DELETE_MAIL_RULE: 'delete_mail_rule',
  GET_AUTO_REPLY: 'get_auto_reply',
  SET_AUTO_REPLY: 'set_auto_reply',
  GET_CATEGORIES: 'get_categories',
};

/**
 * Parse action commands from AI response
 * Looks for JSON blocks marked with [ACTION] tags or structured commands
 */
export function parseActionsFromResponse(aiResponse: string): ParsedAction[] {
  const actions: ParsedAction[] = [];
  
  // Pattern 1: [ACTION:type]{json_params}[/ACTION]
  const actionTagPattern = /\[ACTION:(\w+)\]([\s\S]*?)\[\/ACTION\]/gi;
  let match;
  
  while ((match = actionTagPattern.exec(aiResponse)) !== null) {
    try {
      const type = match[1].toLowerCase();
      const paramsStr = match[2].trim();
      const params = paramsStr ? JSON.parse(paramsStr) : {};
      actions.push({ type, params });
    } catch (e) {
      console.error('Failed to parse action:', match[0], e);
    }
  }
  
  // Pattern 2: ```action\n{type: "...", ...}\n```
  const codeBlockPattern = /```action\s*([\s\S]*?)```/gi;
  
  while ((match = codeBlockPattern.exec(aiResponse)) !== null) {
    try {
      const actionData = JSON.parse(match[1].trim());
      if (actionData.type) {
        actions.push({
          type: actionData.type.toLowerCase(),
          params: actionData,
        });
      }
    } catch (e) {
      console.error('Failed to parse action code block:', match[0], e);
    }
  }
  
  // Pattern 3: EXECUTE: action_name param1="value1" param2="value2"
  const executePattern = /EXECUTE:\s*(\w+)\s+(.+?)(?:\n|$)/gi;
  
  while ((match = executePattern.exec(aiResponse)) !== null) {
    const type = match[1].toLowerCase();
    const paramsStr = match[2];
    const params: Record<string, any> = {};
    
    // Parse key="value" pairs
    const paramPattern = /(\w+)="([^"]+)"/g;
    let paramMatch;
    while ((paramMatch = paramPattern.exec(paramsStr)) !== null) {
      params[paramMatch[1]] = paramMatch[2];
    }
    
    actions.push({ type, params });
  }
  
  return actions;
}

/**
 * Execute a single action
 */
export async function executeAction(action: ParsedAction): Promise<ActionResult> {
  const { type, params } = action;
  
  try {
    switch (type) {
      // ==================== EMAIL ACTIONS ====================
      
      case AVAILABLE_ACTIONS.SEND_EMAIL:
      case 'send': {
        const { to, subject, body, cc, bcc } = params;
        if (!to || !subject || !body) {
          return {
            success: false,
            action: type,
            message: 'Missing required parameters: to, subject, body',
          };
        }
        // Convert arrays to comma-separated strings
        const toStr = Array.isArray(to) ? to.join(', ') : to;
        const ccStr = cc ? (Array.isArray(cc) ? cc.join(', ') : cc) : undefined;
        const bccStr = bcc ? (Array.isArray(bcc) ? bcc.join(', ') : bcc) : undefined;
        
        const success = await graphService.sendEmail(toStr, subject, body, ccStr, bccStr);
        return {
          success,
          action: type,
          message: success ? `Email sent to ${toStr}` : 'Failed to send email',
        };
      }
      
      case AVAILABLE_ACTIONS.REPLY_EMAIL:
      case 'reply': {
        const { emailId, body } = params;
        if (!emailId || !body) {
          return {
            success: false,
            action: type,
            message: 'Missing required parameters: emailId, body',
          };
        }
        const success = await graphService.replyToEmail(emailId, body);
        return {
          success,
          action: type,
          message: success ? 'Reply sent successfully' : 'Failed to send reply',
        };
      }
      
      case AVAILABLE_ACTIONS.FORWARD_EMAIL:
      case 'forward': {
        const { emailId, to, comment } = params;
        if (!emailId || !to) {
          return {
            success: false,
            action: type,
            message: 'Missing required parameters: emailId, to',
          };
        }
        const toArray = Array.isArray(to) ? to : [to];
        const success = await graphService.forwardEmail(emailId, toArray, comment);
        return {
          success,
          action: type,
          message: success ? `Email forwarded to ${toArray.join(', ')}` : 'Failed to forward email',
        };
      }
      
      case AVAILABLE_ACTIONS.DELETE_EMAIL:
      case 'delete': {
        const { emailId } = params;
        if (!emailId) {
          return { success: false, action: type, message: 'Missing emailId' };
        }
        const success = await graphService.deleteEmail(emailId);
        return {
          success,
          action: type,
          message: success ? 'Email deleted' : 'Failed to delete email',
        };
      }
      
      case AVAILABLE_ACTIONS.ARCHIVE_EMAIL:
      case 'archive': {
        const { emailId } = params;
        if (!emailId) {
          return { success: false, action: type, message: 'Missing emailId' };
        }
        const success = await graphService.archiveEmail(emailId);
        return {
          success,
          action: type,
          message: success ? 'Email archived' : 'Failed to archive email',
        };
      }
      
      case AVAILABLE_ACTIONS.MARK_READ:
      case 'mark_read': {
        const { emailId, emailIds } = params;
        if (emailIds && Array.isArray(emailIds)) {
          const result = await graphService.markEmailsAsRead(emailIds);
          const success = result.success > 0;
          return {
            success,
            action: type,
            message: success 
              ? `Marked ${result.success} emails as read${result.failed > 0 ? ` (${result.failed} failed)` : ''}`
              : 'Failed to mark emails as read',
          };
        }
        if (!emailId) {
          return { success: false, action: type, message: 'Missing emailId or emailIds' };
        }
        const success = await graphService.markEmailAsRead(emailId);
        return {
          success,
          action: type,
          message: success ? 'Email marked as read' : 'Failed to mark as read',
        };
      }
      
      case AVAILABLE_ACTIONS.MARK_UNREAD:
      case 'mark_unread': {
        const { emailId } = params;
        if (!emailId) {
          return { success: false, action: type, message: 'Missing emailId' };
        }
        const success = await graphService.markEmailAsUnread(emailId);
        return {
          success,
          action: type,
          message: success ? 'Email marked as unread' : 'Failed to mark as unread',
        };
      }
      
      case AVAILABLE_ACTIONS.FLAG_EMAIL:
      case 'flag': {
        const { emailId } = params;
        if (!emailId) {
          return { success: false, action: type, message: 'Missing emailId' };
        }
        const success = await graphService.flagEmail(emailId, true);
        return {
          success,
          action: type,
          message: success ? 'Email flagged' : 'Failed to flag email',
        };
      }
      
      case AVAILABLE_ACTIONS.UNFLAG_EMAIL:
      case 'unflag': {
        const { emailId } = params;
        if (!emailId) {
          return { success: false, action: type, message: 'Missing emailId' };
        }
        const success = await graphService.flagEmail(emailId, false);
        return {
          success,
          action: type,
          message: success ? 'Email unflagged' : 'Failed to unflag email',
        };
      }
      
      case AVAILABLE_ACTIONS.MOVE_EMAIL:
      case 'move': {
        const { emailId, folderId, folderName } = params;
        if (!emailId) {
          return { success: false, action: type, message: 'Missing emailId' };
        }
        
        let targetFolderId = folderId;
        if (!targetFolderId && folderName) {
          const folders = await graphService.getFolders();
          const folder = folders.find(f => 
            f.displayName.toLowerCase() === folderName.toLowerCase()
          );
          if (folder) {
            targetFolderId = folder.id;
          } else {
            return {
              success: false,
              action: type,
              message: `Folder "${folderName}" not found`,
              data: { availableFolders: folders.map(f => f.displayName) },
            };
          }
        }
        
        if (!targetFolderId) {
          return { success: false, action: type, message: 'Missing folderId or folderName' };
        }
        
        const success = await graphService.moveEmailToFolder(emailId, targetFolderId);
        return {
          success,
          action: type,
          message: success ? 'Email moved successfully' : 'Failed to move email',
        };
      }
      
      case AVAILABLE_ACTIONS.CREATE_DRAFT:
      case 'draft': {
        const { to, subject, body, cc, bcc } = params;
        if (!to || !subject) {
          return {
            success: false,
            action: type,
            message: 'Missing required parameters: to, subject',
          };
        }
        const toArray = Array.isArray(to) ? to : [to];
        const draft = await graphService.createDraft(
          toArray,
          subject,
          body || '',
          cc ? (Array.isArray(cc) ? cc : [cc]) : undefined,
          bcc ? (Array.isArray(bcc) ? bcc : [bcc]) : undefined
        );
        return {
          success: !!draft,
          action: type,
          message: draft ? 'Draft created successfully' : 'Failed to create draft',
          data: draft,
        };
      }
      
      // ==================== SEARCH/QUERY ACTIONS ====================
      
      case AVAILABLE_ACTIONS.SEARCH_EMAILS:
      case 'search': {
        const { query, count = 20 } = params;
        if (!query) {
          return { success: false, action: type, message: 'Missing search query' };
        }
        const emails = await graphService.searchEmails(query, count);
        return {
          success: true,
          action: type,
          message: `Found ${emails.length} emails matching "${query}"`,
          data: emails,
        };
      }
      
      case AVAILABLE_ACTIONS.GET_UNREAD:
      case 'unread': {
        const { count = 50 } = params;
        const emails = await graphService.getUnreadEmails(count);
        return {
          success: true,
          action: type,
          message: `Found ${emails.length} unread emails`,
          data: emails,
        };
      }
      
      case AVAILABLE_ACTIONS.GET_EMAIL_DETAILS:
      case 'details': {
        const { emailId } = params;
        if (!emailId) {
          return { success: false, action: type, message: 'Missing emailId' };
        }
        const email = await graphService.getEmailDetails(emailId);
        return {
          success: !!email,
          action: type,
          message: email ? 'Email details retrieved' : 'Failed to get email details',
          data: email,
        };
      }
      
      case AVAILABLE_ACTIONS.GET_EMAILS_FROM_SENDER:
      case 'from_sender': {
        const { senderEmail, count = 10 } = params;
        if (!senderEmail) {
          return { success: false, action: type, message: 'Missing senderEmail' };
        }
        const emails = await graphService.getEmailsFromSender(senderEmail, count);
        return {
          success: true,
          action: type,
          message: `Found ${emails.length} emails from ${senderEmail}`,
          data: emails,
        };
      }
      
      // ==================== CALENDAR ACTIONS ====================
      
      case AVAILABLE_ACTIONS.CREATE_EVENT:
      case 'schedule':
      case 'meeting': {
        const { subject, start, end, attendees, body, location } = params;
        if (!subject || !start || !end) {
          return {
            success: false,
            action: type,
            message: 'Missing required parameters: subject, start, end',
          };
        }
        const startDate = new Date(start);
        const endDate = new Date(end);
        const success = await graphService.createCalendarEvent(
          subject,
          startDate,
          endDate,
          attendees,
          body,
          location
        );
        return {
          success,
          action: type,
          message: success ? `Event "${subject}" created` : 'Failed to create event',
        };
      }
      
      case AVAILABLE_ACTIONS.DELETE_EVENT:
      case 'cancel_meeting': {
        const { eventId } = params;
        if (!eventId) {
          return { success: false, action: type, message: 'Missing eventId' };
        }
        const success = await graphService.deleteCalendarEvent(eventId);
        return {
          success,
          action: type,
          message: success ? 'Event deleted' : 'Failed to delete event',
        };
      }
      
      case AVAILABLE_ACTIONS.GET_CALENDAR:
      case 'calendar': {
        const { days = 7 } = params;
        const events = await graphService.getCalendarEvents(days);
        return {
          success: true,
          action: type,
          message: `Found ${events.length} events in the next ${days} days`,
          data: events,
        };
      }
      
      // ==================== TASK ACTIONS ====================
      
      case AVAILABLE_ACTIONS.CREATE_TASK:
      case 'task':
      case 'todo': {
        const { title, dueDate, body, linkedEmailId } = params;
        if (!title) {
          return { success: false, action: type, message: 'Missing task title' };
        }
        const task = await graphService.createTask(
          title,
          dueDate ? new Date(dueDate) : undefined,
          body,
          linkedEmailId
        );
        return {
          success: !!task,
          action: type,
          message: task ? `Task "${title}" created` : 'Failed to create task',
          data: task,
        };
      }
      
      case AVAILABLE_ACTIONS.GET_TASKS:
      case 'tasks': {
        const { count = 20 } = params;
        const tasks = await graphService.getTasks(count);
        return {
          success: true,
          action: type,
          message: `Found ${tasks.length} tasks`,
          data: tasks,
        };
      }
      
      // ==================== CONTACT ACTIONS ====================
      
      case AVAILABLE_ACTIONS.GET_CONTACTS:
      case 'contacts': {
        const { count = 50 } = params;
        const contacts = await graphService.getContacts(count);
        return {
          success: true,
          action: type,
          message: `Found ${contacts.length} contacts`,
          data: contacts,
        };
      }
      
      case AVAILABLE_ACTIONS.SEARCH_CONTACTS:
      case 'find_contact': {
        const { query } = params;
        if (!query) {
          return { success: false, action: type, message: 'Missing search query' };
        }
        const contacts = await graphService.searchContacts(query);
        return {
          success: true,
          action: type,
          message: `Found ${contacts.length} contacts matching "${query}"`,
          data: contacts,
        };
      }
      
      case AVAILABLE_ACTIONS.CREATE_CONTACT: {
        const { givenName, surname, displayName, emailAddresses, businessPhones, mobilePhone, companyName, jobTitle } = params;
        const contact = await graphService.createContact({
          givenName,
          surname,
          displayName,
          emailAddresses: emailAddresses?.map((e: string) => ({ address: e })),
          businessPhones,
          mobilePhone,
          companyName,
          jobTitle,
        });
        return {
          success: !!contact,
          action: type,
          message: contact ? `Contact "${displayName || givenName}" created` : 'Failed to create contact',
          data: contact,
        };
      }
      
      case AVAILABLE_ACTIONS.UPDATE_CONTACT: {
        const { contactId, ...updates } = params;
        if (!contactId) {
          return { success: false, action: type, message: 'Missing contactId' };
        }
        const success = await graphService.updateContact(contactId, updates);
        return {
          success,
          action: type,
          message: success ? 'Contact updated' : 'Failed to update contact',
        };
      }
      
      case AVAILABLE_ACTIONS.DELETE_CONTACT: {
        const { contactId } = params;
        if (!contactId) {
          return { success: false, action: type, message: 'Missing contactId' };
        }
        const success = await graphService.deleteContact(contactId);
        return {
          success,
          action: type,
          message: success ? 'Contact deleted' : 'Failed to delete contact',
        };
      }
      
      // ==================== FOLDER ACTIONS ====================
      
      case AVAILABLE_ACTIONS.GET_FOLDERS:
      case 'folders': {
        const folders = await graphService.getFolders();
        return {
          success: true,
          action: type,
          message: `Found ${folders.length} folders`,
          data: folders,
        };
      }
      
      case AVAILABLE_ACTIONS.CREATE_FOLDER: {
        const { displayName, parentFolderId } = params;
        if (!displayName) {
          return { success: false, action: type, message: 'Missing folder name' };
        }
        const folder = await graphService.createFolder(displayName, parentFolderId);
        return {
          success: !!folder,
          action: type,
          message: folder ? `Folder "${displayName}" created` : 'Failed to create folder',
          data: folder,
        };
      }
      
      case AVAILABLE_ACTIONS.RENAME_FOLDER: {
        const { folderId, newName } = params;
        if (!folderId || !newName) {
          return { success: false, action: type, message: 'Missing folderId or newName' };
        }
        const success = await graphService.renameFolder(folderId, newName);
        return {
          success,
          action: type,
          message: success ? `Folder renamed to "${newName}"` : 'Failed to rename folder',
        };
      }
      
      case AVAILABLE_ACTIONS.DELETE_FOLDER: {
        const { folderId } = params;
        if (!folderId) {
          return { success: false, action: type, message: 'Missing folderId' };
        }
        const success = await graphService.deleteFolder(folderId);
        return {
          success,
          action: type,
          message: success ? 'Folder deleted' : 'Failed to delete folder',
        };
      }
      
      // ==================== NEW EMAIL ACTIONS ====================
      
      case AVAILABLE_ACTIONS.REPLY_ALL:
      case 'reply_all': {
        const { emailId, body } = params;
        if (!emailId || !body) {
          return { success: false, action: type, message: 'Missing emailId or body' };
        }
        const success = await graphService.replyAllToEmail(emailId, body);
        return {
          success,
          action: type,
          message: success ? 'Reply all sent' : 'Failed to reply all',
        };
      }
      
      case AVAILABLE_ACTIONS.SET_CATEGORIES: {
        const { emailId, categories } = params;
        if (!emailId || !categories) {
          return { success: false, action: type, message: 'Missing emailId or categories' };
        }
        const success = await graphService.setEmailCategories(emailId, categories);
        return {
          success,
          action: type,
          message: success ? 'Categories set' : 'Failed to set categories',
        };
      }
      
      case AVAILABLE_ACTIONS.SET_IMPORTANCE: {
        const { emailId, importance } = params;
        if (!emailId || !importance) {
          return { success: false, action: type, message: 'Missing emailId or importance' };
        }
        const success = await graphService.setEmailImportance(emailId, importance);
        return {
          success,
          action: type,
          message: success ? `Importance set to ${importance}` : 'Failed to set importance',
        };
      }
      
      case AVAILABLE_ACTIONS.GET_ATTACHMENTS: {
        const { emailId } = params;
        if (!emailId) {
          return { success: false, action: type, message: 'Missing emailId' };
        }
        const attachments = await graphService.getAttachments(emailId);
        return {
          success: true,
          action: type,
          message: `Found ${attachments.length} attachments`,
          data: attachments,
        };
      }
      
      case AVAILABLE_ACTIONS.GET_CONVERSATION: {
        const { conversationId, count = 20 } = params;
        if (!conversationId) {
          return { success: false, action: type, message: 'Missing conversationId' };
        }
        const emails = await graphService.getConversationThread(conversationId, count);
        return {
          success: true,
          action: type,
          message: `Found ${emails.length} messages in conversation`,
          data: emails,
        };
      }
      
      case AVAILABLE_ACTIONS.GET_SENT_ITEMS:
      case 'sent': {
        const { count = 20 } = params;
        const emails = await graphService.getSentItems(count);
        return {
          success: true,
          action: type,
          message: `Found ${emails.length} sent items`,
          data: emails,
        };
      }
      
      case AVAILABLE_ACTIONS.GET_DRAFTS:
      case 'drafts': {
        const { count = 20 } = params;
        const drafts = await graphService.getDrafts(count);
        return {
          success: true,
          action: type,
          message: `Found ${drafts.length} drafts`,
          data: drafts,
        };
      }
      
      case AVAILABLE_ACTIONS.SEND_DRAFT: {
        const { draftId } = params;
        if (!draftId) {
          return { success: false, action: type, message: 'Missing draftId' };
        }
        const success = await graphService.sendDraft(draftId);
        return {
          success,
          action: type,
          message: success ? 'Draft sent' : 'Failed to send draft',
        };
      }
      
      case AVAILABLE_ACTIONS.GET_CATEGORIES: {
        const categories = await graphService.getCategories();
        return {
          success: true,
          action: type,
          message: `Found ${categories.length} categories`,
          data: categories,
        };
      }
      
      // ==================== NEW CALENDAR ACTIONS ====================
      
      case AVAILABLE_ACTIONS.UPDATE_EVENT: {
        const { eventId, subject, start, end, location, body, attendees } = params;
        if (!eventId) {
          return { success: false, action: type, message: 'Missing eventId' };
        }
        const success = await graphService.updateCalendarEvent(eventId, {
          subject,
          start: start ? new Date(start) : undefined,
          end: end ? new Date(end) : undefined,
          location,
          body,
          attendees,
        });
        return {
          success,
          action: type,
          message: success ? 'Event updated' : 'Failed to update event',
        };
      }
      
      case AVAILABLE_ACTIONS.CREATE_RECURRING_EVENT: {
        const { subject, start, end, recurrence, attendees, location } = params;
        if (!subject || !start || !end || !recurrence) {
          return { success: false, action: type, message: 'Missing required parameters' };
        }
        const success = await graphService.createRecurringEvent(
          subject,
          new Date(start),
          new Date(end),
          {
            pattern: recurrence.pattern,
            interval: recurrence.interval || 1,
            daysOfWeek: recurrence.daysOfWeek,
            endDate: recurrence.endDate ? new Date(recurrence.endDate) : undefined,
            occurrences: recurrence.occurrences,
          },
          attendees,
          location
        );
        return {
          success,
          action: type,
          message: success ? `Recurring event "${subject}" created` : 'Failed to create recurring event',
        };
      }
      
      case AVAILABLE_ACTIONS.ACCEPT_MEETING:
      case 'accept': {
        const { eventId, comment, sendResponse = true } = params;
        if (!eventId) {
          return { success: false, action: type, message: 'Missing eventId' };
        }
        const success = await graphService.acceptMeeting(eventId, comment, sendResponse);
        return {
          success,
          action: type,
          message: success ? 'Meeting accepted' : 'Failed to accept meeting',
        };
      }
      
      case AVAILABLE_ACTIONS.DECLINE_MEETING:
      case 'decline': {
        const { eventId, comment, sendResponse = true } = params;
        if (!eventId) {
          return { success: false, action: type, message: 'Missing eventId' };
        }
        const success = await graphService.declineMeeting(eventId, comment, sendResponse);
        return {
          success,
          action: type,
          message: success ? 'Meeting declined' : 'Failed to decline meeting',
        };
      }
      
      case AVAILABLE_ACTIONS.TENTATIVE_MEETING:
      case 'tentative': {
        const { eventId, comment, sendResponse = true } = params;
        if (!eventId) {
          return { success: false, action: type, message: 'Missing eventId' };
        }
        const success = await graphService.tentativelyAcceptMeeting(eventId, comment, sendResponse);
        return {
          success,
          action: type,
          message: success ? 'Meeting tentatively accepted' : 'Failed to tentatively accept meeting',
        };
      }
      
      case AVAILABLE_ACTIONS.GET_FREE_BUSY: {
        const { emails, start, end } = params;
        if (!emails || !start || !end) {
          return { success: false, action: type, message: 'Missing emails, start, or end' };
        }
        const schedules = await graphService.getFreeBusy(
          Array.isArray(emails) ? emails : [emails],
          new Date(start),
          new Date(end)
        );
        return {
          success: true,
          action: type,
          message: `Retrieved availability for ${schedules.length} people`,
          data: schedules,
        };
      }
      
      // ==================== NEW TASK ACTIONS ====================
      
      case AVAILABLE_ACTIONS.COMPLETE_TASK: {
        const { taskId, listId } = params;
        if (!taskId) {
          return { success: false, action: type, message: 'Missing taskId' };
        }
        const success = await graphService.completeTask(taskId, listId);
        return {
          success,
          action: type,
          message: success ? 'Task completed' : 'Failed to complete task',
        };
      }
      
      case AVAILABLE_ACTIONS.DELETE_TASK: {
        const { taskId, listId } = params;
        if (!taskId) {
          return { success: false, action: type, message: 'Missing taskId' };
        }
        const success = await graphService.deleteTask(taskId, listId);
        return {
          success,
          action: type,
          message: success ? 'Task deleted' : 'Failed to delete task',
        };
      }
      
      case AVAILABLE_ACTIONS.UPDATE_TASK: {
        const { taskId, title, body, dueDate, importance, listId } = params;
        if (!taskId) {
          return { success: false, action: type, message: 'Missing taskId' };
        }
        const success = await graphService.updateTask(taskId, {
          title,
          body,
          dueDate: dueDate ? new Date(dueDate) : undefined,
          importance,
        }, listId);
        return {
          success,
          action: type,
          message: success ? 'Task updated' : 'Failed to update task',
        };
      }
      
      // ==================== MAIL RULES & SETTINGS ====================
      
      case AVAILABLE_ACTIONS.GET_MAIL_RULES: {
        const rules = await graphService.getMailRules();
        return {
          success: true,
          action: type,
          message: `Found ${rules.length} mail rules`,
          data: rules,
        };
      }
      
      case AVAILABLE_ACTIONS.CREATE_MAIL_RULE: {
        const { displayName, conditions, actions: ruleActions } = params;
        if (!displayName || !conditions || !ruleActions) {
          return { success: false, action: type, message: 'Missing displayName, conditions, or actions' };
        }
        const rule = await graphService.createMailRule({
          displayName,
          conditions,
          actions: ruleActions,
        });
        return {
          success: !!rule,
          action: type,
          message: rule ? `Mail rule "${displayName}" created` : 'Failed to create mail rule',
          data: rule,
        };
      }
      
      case AVAILABLE_ACTIONS.DELETE_MAIL_RULE: {
        const { ruleId } = params;
        if (!ruleId) {
          return { success: false, action: type, message: 'Missing ruleId' };
        }
        const success = await graphService.deleteMailRule(ruleId);
        return {
          success,
          action: type,
          message: success ? 'Mail rule deleted' : 'Failed to delete mail rule',
        };
      }
      
      case AVAILABLE_ACTIONS.GET_AUTO_REPLY:
      case 'oof':
      case 'out_of_office': {
        const settings = await graphService.getAutoReply();
        return {
          success: !!settings,
          action: type,
          message: settings ? `Auto-reply status: ${settings.status}` : 'Failed to get auto-reply settings',
          data: settings,
        };
      }
      
      case AVAILABLE_ACTIONS.SET_AUTO_REPLY: {
        const { status, externalMessage, internalMessage, scheduledStartDateTime, scheduledEndDateTime, externalAudience } = params;
        if (!status) {
          return { success: false, action: type, message: 'Missing status' };
        }
        const success = await graphService.setAutoReply({
          status,
          externalMessage,
          internalMessage,
          scheduledStartDateTime: scheduledStartDateTime ? new Date(scheduledStartDateTime) : undefined,
          scheduledEndDateTime: scheduledEndDateTime ? new Date(scheduledEndDateTime) : undefined,
          externalAudience,
        });
        return {
          success,
          action: type,
          message: success ? `Auto-reply ${status === 'disabled' ? 'disabled' : 'enabled'}` : 'Failed to set auto-reply',
        };
      }
      
      default:
        return {
          success: false,
          action: type,
          message: `Unknown action type: ${type}`,
        };
    }
  } catch (error) {
    console.error(`Error executing action ${type}:`, error);
    return {
      success: false,
      action: type,
      message: `Error: ${error instanceof Error ? error.message : 'Unknown error'}`,
    };
  }
}

/**
 * Execute multiple actions from AI response
 */
export async function executeActionsFromResponse(
  aiResponse: string
): Promise<ActionResult[]> {
  const actions = parseActionsFromResponse(aiResponse);
  const results: ActionResult[] = [];
  
  for (const action of actions) {
    const result = await executeAction(action);
    results.push(result);
  }
  
  return results;
}

/**
 * Get a description of all available actions for the AI prompt
 */
export function getAvailableActionsDescription(): string {
  return `
## Available Actions

You can execute Outlook actions by including action commands in your response. Use ONE of these formats:

### Format 1: Action Tags
[ACTION:action_name]{"param1": "value1", "param2": "value2"}[/ACTION]

### Format 2: Code Block
\`\`\`action
{"type": "action_name", "param1": "value1"}
\`\`\`

### Available Actions:

**Email Actions:**
- send_email: Send a new email
  - Required: to (string or array), subject, body
  - Optional: cc, bcc
  
- reply_email: Reply to an email
  - Required: emailId, body
  
- forward_email: Forward an email
  - Required: emailId, to (string or array)
  - Optional: comment
  
- delete_email: Delete an email
  - Required: emailId
  
- archive_email: Archive an email
  - Required: emailId
  
- mark_read: Mark email(s) as read
  - Required: emailId OR emailIds (array)
  
- mark_unread: Mark email as unread
  - Required: emailId
  
- flag_email: Flag an email for follow-up
  - Required: emailId
  
- unflag_email: Remove flag from email
  - Required: emailId
  
- move_email: Move email to folder
  - Required: emailId, folderName OR folderId
  
- create_draft: Create a draft email
  - Required: to, subject
  - Optional: body, cc, bcc

**Calendar Actions:**
- create_event: Create a calendar event
  - Required: subject, start (ISO datetime), end (ISO datetime)
  - Optional: attendees (array), body, location
  
- delete_event: Delete a calendar event
  - Required: eventId
  
- get_calendar: Get upcoming calendar events
  - Optional: days (default 7)

**Task Actions:**
- create_task: Create a To-Do task
  - Required: title
  - Optional: dueDate (ISO datetime), body, linkedEmailId
  
- get_tasks: Get upcoming tasks
  - Optional: count (default 20)

**Search Actions:**
- search_emails: Search emails by keyword
  - Required: query
  - Optional: count (default 20)
  
- get_unread: Get unread emails
  - Optional: count (default 50)
  
- get_email_details: Get full email content
  - Required: emailId
  
- get_emails_from_sender: Get emails from specific sender
  - Required: senderEmail
  - Optional: count (default 10)

**Contact Actions:**
- get_contacts: Get contacts from address book
  - Optional: count (default 50)
  
- search_contacts: Search for a contact
  - Required: query

**Folder Actions:**
- get_folders: Get all mail folders

### Example Usage:

To send an email:
[ACTION:send_email]{"to": "john@example.com", "subject": "Meeting Tomorrow", "body": "Hi John, Let's meet at 2pm."}[/ACTION]

To create a calendar event:
[ACTION:create_event]{"subject": "Team Standup", "start": "2024-01-15T10:00:00", "end": "2024-01-15T10:30:00", "attendees": ["team@example.com"]}[/ACTION]

To create a task from an email:
[ACTION:create_task]{"title": "Review proposal", "dueDate": "2024-01-20", "linkedEmailId": "AAMkAGI..."}[/ACTION]
`;
}
