import React, { useEffect, useState } from 'react';
import { useAppStore } from '../store/appStore';

interface Analytics {
  emailsProcessed: number;
  emailsTrend: number;
  tasksCompleted: number;
  tasksTrend: number;
  meetingsScheduled: number;
  meetingsTrend: number;
  timeSaved: number;
  responseTime: {
    average: number;
    improvement: number;
  };
  topSenders: { name: string; count: number }[];
  productivityScore: number;
  weeklyActivity: { day: string; emails: number; tasks: number }[];
}

const AnalyticsPanel: React.FC = () => {
  const { tasks, messages, calendarEvents, emailSummaries } = useAppStore();
  const [analytics, setAnalytics] = useState<Analytics | null>(null);

  useEffect(() => {
    // Calculate analytics from store data
    const completedTasks = tasks.filter(t => t.status === 'completed').length;
    const totalTasks = tasks.length;
    
    const mockAnalytics: Analytics = {
      emailsProcessed: emailSummaries.length + Math.floor(Math.random() * 50) + 20,
      emailsTrend: 12,
      tasksCompleted: completedTasks || Math.floor(Math.random() * 20) + 5,
      tasksTrend: 8,
      meetingsScheduled: calendarEvents.length || Math.floor(Math.random() * 10) + 3,
      meetingsTrend: -5,
      timeSaved: Math.floor(Math.random() * 5) + 2,
      responseTime: {
        average: 15,
        improvement: 23,
      },
      topSenders: [
        { name: 'Team Lead', count: 12 },
        { name: 'Product Manager', count: 8 },
        { name: 'HR Department', count: 6 },
        { name: 'IT Support', count: 5 },
      ],
      productivityScore: Math.min(100, Math.floor((completedTasks / Math.max(totalTasks, 1)) * 100) + 60),
      weeklyActivity: [
        { day: 'Mon', emails: 24, tasks: 8 },
        { day: 'Tue', emails: 31, tasks: 12 },
        { day: 'Wed', emails: 28, tasks: 6 },
        { day: 'Thu', emails: 35, tasks: 15 },
        { day: 'Fri', emails: 22, tasks: 9 },
        { day: 'Sat', emails: 5, tasks: 2 },
        { day: 'Sun', emails: 3, tasks: 1 },
      ],
    };
    
    setAnalytics(mockAnalytics);
  }, [tasks, calendarEvents, emailSummaries]);

  if (!analytics) {
    return (
      <div className="loading-spinner">
        <div className="spinner" />
      </div>
    );
  }

  const maxActivity = Math.max(...analytics.weeklyActivity.map(d => d.emails + d.tasks));

  return (
    <div>
      <h2 className="section-title">📊 Analytics & Insights</h2>

      {/* Productivity Score */}
      <div className="analytics-card">
        <div className="flex justify-between items-center">
          <div>
            <h4>Productivity Score</h4>
            <div className="analytics-value">{analytics.productivityScore}%</div>
            <div className="analytics-change">
              ↑ 15% vs last week
            </div>
          </div>
          <div className="text-icon">
            {analytics.productivityScore >= 80 ? '🏆' : 
             analytics.productivityScore >= 60 ? '⭐' : '📈'}
          </div>
        </div>
        <div className="progress-bar mt-8">
          <div 
            className="progress-fill" 
            style={{ width: `${analytics.productivityScore}%` }} 
          />
        </div>
      </div>

      {/* Stats Grid */}
      <div className="analytics-grid">
        <div className="card">
          <div className="text-small text-muted">Emails Processed</div>
          <div className="flex items-center gap-8 mt-4">
            <span className="text-xl text-bold">{analytics.emailsProcessed}</span>
            <span className={`text-small ${analytics.emailsTrend >= 0 ? 'text-success' : 'text-danger'}`}>
              {analytics.emailsTrend >= 0 ? '↑' : '↓'} {Math.abs(analytics.emailsTrend)}%
            </span>
          </div>
        </div>
        
        <div className="card">
          <div className="text-small text-muted">Tasks Completed</div>
          <div className="flex items-center gap-8 mt-4">
            <span className="text-xl text-bold">{analytics.tasksCompleted}</span>
            <span className={`text-small ${analytics.tasksTrend >= 0 ? 'text-success' : 'text-danger'}`}>
              {analytics.tasksTrend >= 0 ? '↑' : '↓'} {Math.abs(analytics.tasksTrend)}%
            </span>
          </div>
        </div>
        
        <div className="card">
          <div className="text-small text-muted">Meetings</div>
          <div className="flex items-center gap-8 mt-4">
            <span className="text-xl text-bold">{analytics.meetingsScheduled}</span>
            <span className={`text-small ${analytics.meetingsTrend >= 0 ? 'text-success' : 'text-danger'}`}>
              {analytics.meetingsTrend >= 0 ? '↑' : '↓'} {Math.abs(analytics.meetingsTrend)}%
            </span>
          </div>
        </div>
        
        <div className="card">
          <div className="text-small text-muted">Time Saved</div>
          <div className="flex items-center gap-8 mt-4">
            <span className="text-xl text-bold">{analytics.timeSaved}h</span>
            <span className="text-small text-success">this week</span>
          </div>
        </div>
      </div>

      {/* Weekly Activity Chart */}
      <div className="card mt-12">
        <h3 className="section-header--lg-margin">📈 Weekly Activity</h3>
        <div className="flex gap-8" style={{ height: 120, alignItems: 'flex-end' }}>
          {analytics.weeklyActivity.map((day, index) => (
            <div key={index} className="flex-1 flex flex-col items-center gap-4">
              <div 
                className="flex flex-col gap-4" 
                style={{ 
                  height: `${((day.emails + day.tasks) / maxActivity) * 80}px`,
                  width: '100%',
                }}
              >
                <div 
                  style={{ 
                    flex: day.emails, 
                    background: '#0078d4', 
                    borderRadius: '4px 4px 0 0',
                    minHeight: 4,
                  }} 
                />
                <div 
                  style={{ 
                    flex: day.tasks, 
                    background: '#107c10', 
                    borderRadius: '0 0 4px 4px',
                    minHeight: 4,
                  }} 
                />
              </div>
              <span className="text-small text-muted">{day.day}</span>
            </div>
          ))}
        </div>
        <div className="flex gap-12 mt-8 justify-center">
          <div className="flex items-center gap-4">
            <div style={{ width: 12, height: 12, background: '#0078d4', borderRadius: 2 }} />
            <span className="text-small text-muted">Emails</span>
          </div>
          <div className="flex items-center gap-4">
            <div style={{ width: 12, height: 12, background: '#107c10', borderRadius: 2 }} />
            <span className="text-small text-muted">Tasks</span>
          </div>
        </div>
      </div>

      {/* Response Time */}
      <div className="card mt-12">
        <h3 className="section-header--lg-margin">⚡ Response Time</h3>
        <div className="stats-row">
          <span className="stats-label">Average Response</span>
          <span className="stats-value">{analytics.responseTime.average} min</span>
        </div>
        <div className="stats-row">
          <span className="stats-label">Improvement</span>
          <span className="stats-value text-success">↑ {analytics.responseTime.improvement}%</span>
        </div>
      </div>

      {/* Top Senders */}
      <div className="card mt-12">
        <h3 className="section-header--lg-margin">📬 Top Email Senders</h3>
        {analytics.topSenders.map((sender, index) => (
          <div key={index} className="stats-row">
            <span className="stats-label">
              {index + 1}. {sender.name}
            </span>
            <span className="stats-value">{sender.count} emails</span>
          </div>
        ))}
      </div>

      {/* AI Insights */}
      <div className="smart-suggestion mt-12">
        <span className="smart-suggestion-icon">💡</span>
        <div className="smart-suggestion-content">
          <div className="smart-suggestion-title">AI Insight</div>
          <div className="smart-suggestion-text">
            You tend to be most productive on Thursdays. Consider scheduling important 
            tasks and meetings for that day to maximize your output.
          </div>
        </div>
      </div>

      {/* Keyboard Shortcuts */}
      <div className="card mt-12">
        <h3 className="section-header--lg-margin">⌨️ Quick Shortcuts</h3>
        <div className="stats-row">
          <span className="stats-label">Summarize Email</span>
          <span><kbd className="kbd">Ctrl</kbd> + <kbd className="kbd">S</kbd></span>
        </div>
        <div className="stats-row">
          <span className="stats-label">Draft Reply</span>
          <span><kbd className="kbd">Ctrl</kbd> + <kbd className="kbd">R</kbd></span>
        </div>
        <div className="stats-row">
          <span className="stats-label">Add Task</span>
          <span><kbd className="kbd">Ctrl</kbd> + <kbd className="kbd">T</kbd></span>
        </div>
        <div className="stats-row">
          <span className="stats-label">Plan Day</span>
          <span><kbd className="kbd">Ctrl</kbd> + <kbd className="kbd">P</kbd></span>
        </div>
      </div>
    </div>
  );
};

export default AnalyticsPanel;
