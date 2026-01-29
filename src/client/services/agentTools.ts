/**
 * Agent Tools - Functions the AI can call directly
 * This gives the AI autonomous control over Outlook functionality
 */

import { outlookService } from './outlookService';
import { graphService } from './graphService';

export interface ToolResult {
  success: boolean;
  data?: any;
  error?: string;
}

export interface Tool {
  name: string;
  description: string;
  parameters: {
    name: string;
    type: string;
    description: string;
    required: boolean;
  }[];
  execute: (params: Record<string, any>) => Promise<ToolResult>;
  requiresApproval?: boolean;
}

/**
 * All available tools the AI agent can use
 */
export const AGENT_TOOLS: Tool[] = [
  // ============ EMAIL TOOLS ============
  {
    name: 'search_emails',
    description: 'Search inbox for emails matching a query. Returns email subjects, senders, dates, and previews.',
    parameters: [
      { name: 'query', type: 'string', description: 'Search query (keywords, sender name, subject text)', required: true },
      { name: 'limit', type: 'number', description: 'Max results to return (default 10)', required: false },
    ],
    execute: async (params) => {
      try {
        const emails = await graphService.searchEmails(params.query, params.limit || 10);
        return { success: true, data: emails };
      } catch (error: any) {
        return { success: false, error: error.message };
      }
    },
  },
  {
    name: 'get_email_details',
    description: 'Get full details of a specific email including body content',
    parameters: [
      { name: 'emailId', type: 'string', description: 'The ID of the email to retrieve', required: true },
    ],
    execute: async (params) => {
      try {
        const email = await graphService.getEmailDetails(params.emailId);
        return { success: true, data: email };
      } catch (error: any) {
        return { success: false, error: error.message };
      }
    },
  },
  {
    name: 'get_recent_emails',
    description: 'Get the most recent emails from inbox',
    parameters: [
      { name: 'count', type: 'number', description: 'Number of emails to retrieve (default 10)', required: false },
    ],
    execute: async (params) => {
      try {
        const emails = await outlookService.getRecentEmails(params.count || 10);
        return { success: true, data: emails };
      } catch (error: any) {
        return { success: false, error: error.message };
      }
    },
  },
  {
    name: 'get_current_email',
    description: 'Get the email currently open/selected in Outlook',
    parameters: [],
    execute: async () => {
      try {
        const email = await outlookService.getCurrentEmailContext();
        return { success: true, data: email };
      } catch (error: any) {
        return { success: false, error: error.message };
      }
    },
  },
  {
    name: 'draft_email',
    description: 'Create a draft email (does NOT send - just prepares it). Returns the draft content for user review.',
    parameters: [
      { name: 'to', type: 'string', description: 'Recipient email address(es), comma-separated', required: true },
      { name: 'subject', type: 'string', description: 'Email subject line', required: true },
      { name: 'body', type: 'string', description: 'Email body content (can be HTML)', required: true },
      { name: 'cc', type: 'string', description: 'CC recipients, comma-separated', required: false },
    ],
    execute: async (params) => {
      // This just returns the draft for approval - doesn't actually send
      return {
        success: true,
        data: {
          type: 'draft_email',
          to: params.to,
          subject: params.subject,
          body: params.body,
          cc: params.cc,
          status: 'pending_approval',
        },
      };
    },
    requiresApproval: true,
  },
  {
    name: 'send_email',
    description: 'Send an email. REQUIRES USER APPROVAL before execution.',
    parameters: [
      { name: 'to', type: 'string', description: 'Recipient email address(es)', required: true },
      { name: 'subject', type: 'string', description: 'Email subject', required: true },
      { name: 'body', type: 'string', description: 'Email body (HTML supported)', required: true },
    ],
    execute: async (params) => {
      try {
        await graphService.sendEmail(params.to, params.subject, params.body);
        return { success: true, data: { message: 'Email sent successfully' } };
      } catch (error: any) {
        return { success: false, error: error.message };
      }
    },
    requiresApproval: true,
  },
  {
    name: 'reply_to_email',
    description: 'Reply to an email. REQUIRES USER APPROVAL.',
    parameters: [
      { name: 'emailId', type: 'string', description: 'ID of email to reply to', required: true },
      { name: 'body', type: 'string', description: 'Reply content', required: true },
      { name: 'replyAll', type: 'boolean', description: 'Reply to all recipients (not yet supported)', required: false },
    ],
    execute: async (params) => {
      try {
        // Note: replyAll not yet supported in graphService
        await graphService.replyToEmail(params.emailId, params.body);
        return { success: true, data: { message: 'Reply sent successfully' } };
      } catch (error: any) {
        return { success: false, error: error.message };
      }
    },
    requiresApproval: true,
  },
  {
    name: 'delete_email',
    description: 'Delete an email (moves to trash). REQUIRES USER APPROVAL.',
    parameters: [
      { name: 'emailId', type: 'string', description: 'ID of email to delete', required: true },
    ],
    execute: async (params) => {
      try {
        await graphService.deleteEmail(params.emailId);
        return { success: true, data: { message: 'Email deleted' } };
      } catch (error: any) {
        return { success: false, error: error.message };
      }
    },
    requiresApproval: true,
  },
  {
    name: 'insert_text_to_compose',
    description: 'Insert text into the currently open compose window (if user is writing an email)',
    parameters: [
      { name: 'text', type: 'string', description: 'Text to insert', required: true },
    ],
    execute: async (params) => {
      try {
        await outlookService.insertTextToCompose(params.text);
        return { success: true, data: { message: 'Text inserted into compose window' } };
      } catch (error: any) {
        return { success: false, error: error.message };
      }
    },
  },

  // ============ CALENDAR TOOLS ============
  {
    name: 'get_calendar_events',
    description: 'Get upcoming calendar events',
    parameters: [
      { name: 'days', type: 'number', description: 'Number of days to look ahead (default 7)', required: false },
    ],
    execute: async (params) => {
      try {
        const events = await outlookService.getUpcomingEvents(params.days || 7);
        return { success: true, data: events };
      } catch (error: any) {
        return { success: false, error: error.message };
      }
    },
  },
  {
    name: 'create_calendar_event',
    description: 'Create a new calendar event/meeting. REQUIRES USER APPROVAL.',
    parameters: [
      { name: 'subject', type: 'string', description: 'Event title', required: true },
      { name: 'start', type: 'string', description: 'Start time (ISO format or natural language)', required: true },
      { name: 'end', type: 'string', description: 'End time (ISO format or natural language)', required: true },
      { name: 'location', type: 'string', description: 'Event location', required: false },
      { name: 'attendees', type: 'string', description: 'Attendee emails, comma-separated', required: false },
      { name: 'body', type: 'string', description: 'Event description/notes', required: false },
    ],
    execute: async (params) => {
      try {
        const attendeeList = params.attendees?.split(',').map((e: string) => e.trim()) || [];
        const event = await graphService.createCalendarEvent(
          params.subject,
          new Date(params.start),
          new Date(params.end),
          attendeeList.length > 0 ? attendeeList : undefined,
          params.body,
          params.location
        );
        return { success: true, data: event };
      } catch (error: any) {
        return { success: false, error: error.message };
      }
    },
    requiresApproval: true,
  },
  {
    name: 'find_free_time',
    description: 'Find available time slots in calendar',
    parameters: [
      { name: 'durationMinutes', type: 'number', description: 'Meeting duration in minutes', required: true },
      { name: 'withinDays', type: 'number', description: 'Search within X days (default 7)', required: false },
    ],
    execute: async (params) => {
      try {
        // Get events and calculate free slots
        const events = await outlookService.getUpcomingEvents(params.withinDays || 7);
        const workHoursStart = 9;
        const workHoursEnd = 17;
        const freeSlots: { start: Date; end: Date }[] = [];
        
        // Simple free time calculation
        const now = new Date();
        for (let day = 0; day < (params.withinDays || 7); day++) {
          const date = new Date(now);
          date.setDate(date.getDate() + day);
          
          // Skip weekends
          if (date.getDay() === 0 || date.getDay() === 6) continue;
          
          const dayStart = new Date(date);
          dayStart.setHours(workHoursStart, 0, 0, 0);
          const dayEnd = new Date(date);
          dayEnd.setHours(workHoursEnd, 0, 0, 0);
          
          // Find conflicts
          const dayEvents = events.filter(e => {
            const eventStart = new Date(e.start);
            return eventStart.toDateString() === date.toDateString();
          });
          
          if (dayEvents.length === 0) {
            freeSlots.push({ start: dayStart, end: dayEnd });
          }
        }
        
        return { success: true, data: freeSlots.slice(0, 5) };
      } catch (error: any) {
        return { success: false, error: error.message };
      }
    },
  },

  // ============ TASK TOOLS ============
  {
    name: 'create_task',
    description: 'Create a new task/to-do item',
    parameters: [
      { name: 'title', type: 'string', description: 'Task title', required: true },
      { name: 'dueDate', type: 'string', description: 'Due date (ISO format)', required: false },
      { name: 'priority', type: 'string', description: 'Priority: high, medium, low', required: false },
      { name: 'notes', type: 'string', description: 'Additional notes', required: false },
    ],
    execute: async (params) => {
      // Store in local state (could connect to MS To-Do API)
      return {
        success: true,
        data: {
          type: 'task_created',
          task: {
            id: Date.now().toString(),
            title: params.title,
            dueDate: params.dueDate,
            priority: params.priority || 'medium',
            notes: params.notes,
            completed: false,
          },
        },
      };
    },
  },

  // ============ USER INFO TOOLS ============
  {
    name: 'get_user_info',
    description: 'Get information about the current user (name, email)',
    parameters: [],
    execute: async () => {
      try {
        const userInfo = graphService.getUserInfo();
        return { success: true, data: userInfo };
      } catch (error: any) {
        return { success: false, error: error.message };
      }
    },
  },

  // ============ UTILITY TOOLS ============
  {
    name: 'get_current_datetime',
    description: 'Get the current date and time',
    parameters: [],
    execute: async () => {
      const now = new Date();
      return {
        success: true,
        data: {
          iso: now.toISOString(),
          formatted: now.toLocaleString(),
          date: now.toLocaleDateString(),
          time: now.toLocaleTimeString(),
          dayOfWeek: now.toLocaleDateString('en-US', { weekday: 'long' }),
        },
      };
    },
  },
  {
    name: 'show_message_to_user',
    description: 'Display a message or information to the user in the chat',
    parameters: [
      { name: 'message', type: 'string', description: 'Message to display', required: true },
      { name: 'type', type: 'string', description: 'Message type: info, success, warning, error', required: false },
    ],
    execute: async (params) => {
      return {
        success: true,
        data: {
          type: 'display_message',
          message: params.message,
          messageType: params.type || 'info',
        },
      };
    },
  },
  {
    name: 'request_user_input',
    description: 'Ask the user for additional information or clarification',
    parameters: [
      { name: 'question', type: 'string', description: 'Question to ask the user', required: true },
      { name: 'options', type: 'string', description: 'Optional: comma-separated list of options to choose from', required: false },
    ],
    execute: async (params) => {
      return {
        success: true,
        data: {
          type: 'request_input',
          question: params.question,
          options: params.options?.split(',').map((o: string) => o.trim()),
        },
      };
    },
  },
];

/**
 * Generate the tools documentation for the AI prompt
 */
export function generateToolsPrompt(): string {
  let prompt = `## AVAILABLE TOOLS

You have access to the following tools. To use a tool, respond with a JSON block in this exact format:

\`\`\`tool
{
  "tool": "tool_name",
  "params": {
    "param1": "value1",
    "param2": "value2"
  }
}
\`\`\`

You can call multiple tools in sequence. After each tool call, you'll receive the result and can decide what to do next.

### Tools:

`;

  for (const tool of AGENT_TOOLS) {
    prompt += `**${tool.name}**${tool.requiresApproval ? ' ⚠️ REQUIRES APPROVAL' : ''}\n`;
    prompt += `${tool.description}\n`;
    if (tool.parameters.length > 0) {
      prompt += `Parameters:\n`;
      for (const param of tool.parameters) {
        prompt += `  - ${param.name} (${param.type}${param.required ? ', required' : ', optional'}): ${param.description}\n`;
      }
    }
    prompt += '\n';
  }

  prompt += `
### Important Guidelines:
1. Always gather information before taking action (search first, then act)
2. For destructive actions (send, delete, create meeting), always confirm with user first
3. You can chain multiple tool calls to complete complex tasks
4. If a tool fails, explain the error and suggest alternatives
5. Use show_message_to_user to communicate progress on multi-step tasks
6. Use request_user_input when you need clarification

### Response Format:
- To use a tool: Include a \`\`\`tool code block with the JSON
- To respond to user: Just write your response naturally
- You can combine both: explain what you're doing AND call tools
`;

  return prompt;
}

/**
 * Parse tool calls from AI response
 */
export function parseToolCalls(response: string): { tool: string; params: Record<string, any> }[] {
  const toolCalls: { tool: string; params: Record<string, any> }[] = [];
  
  // Match ```tool ... ``` blocks
  const toolBlockRegex = /```tool\s*([\s\S]*?)```/g;
  let match;
  
  while ((match = toolBlockRegex.exec(response)) !== null) {
    try {
      const json = JSON.parse(match[1].trim());
      if (json.tool) {
        toolCalls.push({
          tool: json.tool,
          params: json.params || {},
        });
      }
    } catch (e) {
      console.error('Failed to parse tool call:', match[1], e);
    }
  }
  
  return toolCalls;
}

/**
 * Execute a tool call
 */
export async function executeTool(
  toolName: string,
  params: Record<string, any>
): Promise<ToolResult> {
  const tool = AGENT_TOOLS.find(t => t.name === toolName);
  
  if (!tool) {
    return { success: false, error: `Unknown tool: ${toolName}` };
  }
  
  try {
    return await tool.execute(params);
  } catch (error: any) {
    return { success: false, error: error.message };
  }
}

/**
 * Check if a tool requires approval
 */
export function toolRequiresApproval(toolName: string): boolean {
  const tool = AGENT_TOOLS.find(t => t.name === toolName);
  return tool?.requiresApproval || false;
}

/**
 * Get tool by name
 */
export function getTool(toolName: string): Tool | undefined {
  return AGENT_TOOLS.find(t => t.name === toolName);
}
