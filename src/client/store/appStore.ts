import { create } from 'zustand';
import { Message, AutomationRule, CalendarEvent, Task, EmailSummary } from '../../shared/types';
import { ApprovalRequest } from '../components/ApprovalModal';

interface AppState {
  // UI State
  isLoading: boolean;
  error: string | null;
  
  // Chat State
  messages: Message[];
  sessions: ChatSession[];
  activeSessionId: string | null;
  
  // Email State
  currentEmail: EmailSummary | null;
  emailSummaries: EmailSummary[];
  
  // Calendar State
  calendarEvents: CalendarEvent[];
  
  // Tasks State
  tasks: Task[];
  
  // Automation Rules
  automationRules: AutomationRule[];
  
  // Approval Queue
  pendingApprovals: ApprovalRequest[];
  currentApproval: ApprovalRequest | null;
  
  // Settings
  settings: {
    aiProvider: 'openai' | 'anthropic' | 'github';
    autoReplyEnabled: boolean;
    calendarSyncEnabled: boolean;
    taskExtractionEnabled: boolean;
    requireApprovalForEmails: boolean;
    requireApprovalForMeetings: boolean;
    requireApprovalForAutomations: boolean;
  };
  
  // Actions
  setLoading: (loading: boolean) => void;
  setError: (error: string | null) => void;
  addMessage: (message: Message) => void;
  clearMessages: () => void;
  createNewSession: () => void;
  loadSession: (sessionId: string) => void;
  deleteSession: (sessionId: string) => void;
  setCurrentEmail: (email: EmailSummary | null) => void;
  setEmailSummaries: (summaries: EmailSummary[]) => void;
  setCalendarEvents: (events: CalendarEvent[]) => void;
  setTasks: (tasks: Task[]) => void;
  addTask: (task: Task) => void;
  updateTask: (taskId: string, updates: Partial<Task>) => void;
  deleteTask: (taskId: string) => void;
  setAutomationRules: (rules: AutomationRule[]) => void;
  addAutomationRule: (rule: AutomationRule) => void;
  updateAutomationRule: (ruleId: string, updates: Partial<AutomationRule>) => void;
  deleteAutomationRule: (ruleId: string) => void;
  updateSettings: (settings: Partial<AppState['settings']>) => void;
  
  // Approval Actions
  addApprovalRequest: (request: ApprovalRequest) => void;
  removeApprovalRequest: (requestId: string) => void;
  setCurrentApproval: (request: ApprovalRequest | null) => void;
  clearApprovals: () => void;
}

interface ChatSession {
  id: string;
  title: string;
  createdAt: string;
  updatedAt: string;
  messages: Message[];
}

const SESSIONS_KEY = 'ff_ai_chat_sessions';

const loadSessions = (): ChatSession[] => {
  try {
    const raw = localStorage.getItem(SESSIONS_KEY);
    if (!raw) return [];
    const sessions = JSON.parse(raw) as ChatSession[];
    return sessions.map((session) => ({
      ...session,
      messages: session.messages.map((m: any) => ({
        ...m,
        timestamp: new Date(m.timestamp),
      })),
    }));
  } catch {
    return [];
  }
};

const saveSessions = (sessions: ChatSession[]) => {
  try {
    localStorage.setItem(SESSIONS_KEY, JSON.stringify(sessions));
  } catch {
    // ignore
  }
};

export const useAppStore = create<AppState>((set) => ({
  // Initial State
  isLoading: false,
  error: null,
  messages: [],
  sessions: loadSessions(),
  activeSessionId: null,
  currentEmail: null,
  emailSummaries: [],
  calendarEvents: [],
  tasks: [],
  automationRules: [],
  pendingApprovals: [],
  currentApproval: null,
  settings: {
    aiProvider: 'github',
    autoReplyEnabled: false,
    calendarSyncEnabled: true,
    taskExtractionEnabled: true,
    requireApprovalForEmails: true,
    requireApprovalForMeetings: true,
    requireApprovalForAutomations: true,
  },

  // Actions
  setLoading: (loading) => set({ isLoading: loading }),
  
  setError: (error) => set({ error }),
  
  addMessage: (message) =>
    set((state) => {
      const messages = [...state.messages, message];
      let sessions = [...state.sessions];
      let activeSessionId = state.activeSessionId;

      if (!activeSessionId) {
        activeSessionId = `session-${Date.now()}`;
        const newSession: ChatSession = {
          id: activeSessionId,
          title: message.role === 'user' ? message.content.substring(0, 50) : 'New Chat',
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString(),
          messages: [],
        };
        sessions = [newSession, ...sessions];
      }

      sessions = sessions.map((s) => {
        if (s.id !== activeSessionId) return s;
        const title = s.title === 'New Chat' && message.role === 'user'
          ? message.content.substring(0, 50)
          : s.title;
        return {
          ...s,
          title,
          updatedAt: new Date().toISOString(),
          messages,
        };
      });

      saveSessions(sessions);
      return { messages, sessions, activeSessionId };
    }),
  
  clearMessages: () => set({ messages: [] }),

  createNewSession: () =>
    set((state) => {
      const id = `session-${Date.now()}`;
      const newSession: ChatSession = {
        id,
        title: 'New Chat',
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
        messages: [],
      };
      const sessions = [newSession, ...state.sessions];
      saveSessions(sessions);
      return { messages: [], sessions, activeSessionId: id };
    }),

  loadSession: (sessionId) =>
    set((state) => {
      const session = state.sessions.find((s) => s.id === sessionId);
      if (!session) return state;
      return { messages: session.messages, activeSessionId: sessionId };
    }),

  deleteSession: (sessionId) =>
    set((state) => {
      const sessions = state.sessions.filter((s) => s.id !== sessionId);
      const activeSessionId = state.activeSessionId === sessionId ? null : state.activeSessionId;
      saveSessions(sessions);
      return { sessions, activeSessionId, messages: activeSessionId ? state.messages : [] };
    }),
  
  setCurrentEmail: (email) => set({ currentEmail: email }),
  
  setEmailSummaries: (summaries) => set({ emailSummaries: summaries }),
  
  setCalendarEvents: (events) => set({ calendarEvents: events }),
  
  setTasks: (tasks) => set({ tasks }),
  
  addTask: (task) =>
    set((state) => ({ tasks: [...state.tasks, task] })),
  
  updateTask: (taskId, updates) =>
    set((state) => ({
      tasks: state.tasks.map((task) =>
        task.id === taskId ? { ...task, ...updates } : task
      ),
    })),
  
  deleteTask: (taskId) =>
    set((state) => ({
      tasks: state.tasks.filter((task) => task.id !== taskId),
    })),
  
  setAutomationRules: (rules) => set({ automationRules: rules }),
  
  addAutomationRule: (rule) =>
    set((state) => ({ automationRules: [...state.automationRules, rule] })),
  
  updateAutomationRule: (ruleId, updates) =>
    set((state) => ({
      automationRules: state.automationRules.map((rule) =>
        rule.id === ruleId ? { ...rule, ...updates } : rule
      ),
    })),
  
  deleteAutomationRule: (ruleId) =>
    set((state) => ({
      automationRules: state.automationRules.filter((rule) => rule.id !== ruleId),
    })),
  
  updateSettings: (newSettings) =>
    set((state) => ({
      settings: { ...state.settings, ...newSettings },
    })),
  
  // Approval Actions
  addApprovalRequest: (request) =>
    set((state) => ({ 
      pendingApprovals: [...state.pendingApprovals, request],
      currentApproval: state.currentApproval || request, // Show first one if none showing
    })),
  
  removeApprovalRequest: (requestId) =>
    set((state) => {
      const newPending = state.pendingApprovals.filter((r) => r.id !== requestId);
      return {
        pendingApprovals: newPending,
        currentApproval: state.currentApproval?.id === requestId 
          ? (newPending[0] || null) 
          : state.currentApproval,
      };
    }),
  
  setCurrentApproval: (request) => set({ currentApproval: request }),
  
  clearApprovals: () => set({ pendingApprovals: [], currentApproval: null }),
}));
