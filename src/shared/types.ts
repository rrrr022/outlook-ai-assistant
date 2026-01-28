// Message Types
export interface Message {
  id: string;
  role: 'user' | 'assistant' | 'system';
  content: string;
  timestamp: Date;
  metadata?: {
    action?: string;
    emailId?: string;
    calendarEventId?: string;
    taskId?: string;
  };
}

// Email Types
export interface EmailSummary {
  id: string;
  subject: string;
  sender: string;
  senderEmail: string;
  receivedDateTime: Date;
  preview: string;
  isRead: boolean;
  importance: 'low' | 'normal' | 'high';
  hasAttachments: boolean;
  conversationId?: string;
}

export interface EmailDetails extends EmailSummary {
  body: string;
  bodyType: 'text' | 'html';
  toRecipients: EmailRecipient[];
  ccRecipients: EmailRecipient[];
  attachments: EmailAttachment[];
}

export interface EmailRecipient {
  name: string;
  email: string;
}

export interface EmailAttachment {
  id: string;
  name: string;
  contentType: string;
  size: number;
}

export interface DraftEmail {
  to: string[];
  cc?: string[];
  bcc?: string[];
  subject: string;
  body: string;
  bodyType: 'text' | 'html';
}

// Calendar Types
export interface CalendarEvent {
  id: string;
  subject: string;
  start: Date;
  end: Date;
  location?: string;
  isAllDay: boolean;
  organizer: string;
  attendees: EventAttendee[];
  body?: string;
  importance: 'low' | 'normal' | 'high';
  showAs: 'free' | 'tentative' | 'busy' | 'oof' | 'workingElsewhere';
  isRecurring: boolean;
  recurrence?: RecurrencePattern;
}

export interface EventAttendee {
  name: string;
  email: string;
  response: 'none' | 'accepted' | 'declined' | 'tentative';
  type: 'required' | 'optional' | 'resource';
}

export interface RecurrencePattern {
  type: 'daily' | 'weekly' | 'monthly' | 'yearly';
  interval: number;
  daysOfWeek?: string[];
  endDate?: Date;
  occurrences?: number;
}

export interface CreateCalendarEvent {
  subject: string;
  start: Date;
  end: Date;
  location?: string;
  body?: string;
  attendees?: string[];
  isAllDay?: boolean;
  reminder?: number;
}

// Task Types
export interface Task {
  id: string;
  title: string;
  description?: string;
  dueDate?: Date;
  priority: 'low' | 'normal' | 'high';
  status: 'notStarted' | 'inProgress' | 'completed' | 'deferred';
  createdAt: Date;
  updatedAt: Date;
  sourceEmailId?: string;
  categories?: string[];
}

export interface CreateTask {
  title: string;
  description?: string;
  dueDate?: Date;
  priority?: 'low' | 'normal' | 'high';
  sourceEmailId?: string;
}

// Automation Rule Types
export type RuleConditionType = 
  | 'sender'
  | 'subject'
  | 'body'
  | 'hasAttachment'
  | 'importance'
  | 'isRead'
  | 'timeOfDay'
  | 'dayOfWeek';

export type RuleActionType =
  | 'autoReply'
  | 'forward'
  | 'categorize'
  | 'createTask'
  | 'scheduleEvent'
  | 'markAsRead'
  | 'moveToFolder'
  | 'notify';

export interface RuleCondition {
  type: RuleConditionType;
  operator: 'equals' | 'contains' | 'startsWith' | 'endsWith' | 'greaterThan' | 'lessThan';
  value: string | boolean | number;
}

export interface RuleAction {
  type: RuleActionType;
  parameters: Record<string, unknown>;
}

export interface AutomationRule {
  id: string;
  name: string;
  description?: string;
  isEnabled: boolean;
  priority: number;
  conditions: RuleCondition[];
  conditionLogic: 'and' | 'or';
  actions: RuleAction[];
  createdAt: Date;
  updatedAt: Date;
  lastTriggered?: Date;
  triggerCount: number;
}

// Inbox Summary Type (for full inbox access)
export interface InboxSummary {
  totalEmails: number;
  unreadCount: number;
  topSenders: Array<{ name: string; email: string; count: number }>;
  recentEmails: EmailSummary[];
}

// AI Types
export interface AIRequest {
  prompt: string;
  context?: {
    currentEmail?: EmailSummary | EmailDetails;
    recentEmails?: EmailSummary[];
    upcomingEvents?: CalendarEvent[];
    pendingTasks?: Task[];
    inboxSummary?: InboxSummary;
    searchResults?: EmailSummary[];
  };
  action?: string;
}

export interface AIResponse {
  content: string;
  suggestions?: string[];
  suggestedActions?: SuggestedAction[];
  extractedTasks?: CreateTask[];
  extractedEvents?: CreateCalendarEvent[];
  draftReply?: DraftEmail;
}

export interface SuggestedAction {
  type: string;
  label: string;
  description: string;
  parameters: Record<string, unknown>;
}

// API Response Types
export interface ApiResponse<T> {
  success: boolean;
  data?: T;
  error?: string;
  message?: string;
}

// Time Slot Types
export interface TimeSlot {
  start: Date;
  end: Date;
  available: boolean;
  event?: CalendarEvent;
}

export interface DayPlan {
  date: Date;
  slots: TimeSlot[];
  tasks: Task[];
  priorities: string[];
}
