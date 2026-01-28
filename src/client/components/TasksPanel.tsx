import React, { useEffect, useState } from 'react';
import { Spinner, Checkbox } from './ui/NativeComponents';
import { useAppStore } from '../store/appStore';
import { aiService } from '../services/aiService';
import { Task } from '@shared/types';
import { v4 as uuidv4 } from 'uuid';

const TasksPanel: React.FC = () => {
  const { tasks, setTasks, addTask, updateTask, deleteTask, currentEmail } = useAppStore();
  const [loading, setLoading] = useState(false);
  const [newTaskTitle, setNewTaskTitle] = useState('');
  const [filter, setFilter] = useState<'all' | 'active' | 'completed'>('all');

  const handleAddTask = () => {
    if (!newTaskTitle.trim()) return;

    const newTask: Task = {
      id: uuidv4(),
      title: newTaskTitle.trim(),
      priority: 'normal',
      status: 'notStarted',
      createdAt: new Date(),
      updatedAt: new Date(),
    };

    addTask(newTask);
    setNewTaskTitle('');
  };

  const handleExtractTasks = async () => {
    if (!currentEmail) {
      alert('Please select an email first');
      return;
    }

    setLoading(true);
    try {
      const response = await aiService.chat({
        prompt: `Extract any action items or tasks from this email. Return them as a simple list.
        
        From: ${currentEmail.sender}
        Subject: ${currentEmail.subject}
        Content: ${currentEmail.preview}`,
      });

      // Parse the response and create tasks
      const lines = response.content.split('\n').filter(line => line.trim());
      lines.forEach((line) => {
        const taskTitle = line.replace(/^[-•*\d.]\s*/, '').trim();
        if (taskTitle && taskTitle.length > 3) {
          addTask({
            id: uuidv4(),
            title: taskTitle,
            priority: 'normal',
            status: 'notStarted',
            createdAt: new Date(),
            updatedAt: new Date(),
            sourceEmailId: currentEmail.id,
          });
        }
      });
    } catch (error) {
      console.error('Error extracting tasks:', error);
    } finally {
      setLoading(false);
    }
  };

  const handleToggleComplete = (task: Task) => {
    updateTask(task.id, {
      status: task.status === 'completed' ? 'notStarted' : 'completed',
      updatedAt: new Date(),
    });
  };

  const handleDeleteTask = (taskId: string) => {
    deleteTask(taskId);
  };

  const handlePrioritize = async () => {
    if (tasks.length === 0) return;

    setLoading(true);
    try {
      const response = await aiService.chat({
        prompt: `Help me prioritize these tasks. Return them in order of importance with brief reasoning:
        
        ${tasks.map(t => `- ${t.title}`).join('\n')}`,
      });

      // Show the prioritization advice (could also reorder tasks based on AI response)
      alert(response.content);
    } catch (error) {
      console.error('Error prioritizing:', error);
    } finally {
      setLoading(false);
    }
  };

  const filteredTasks = tasks.filter((task) => {
    switch (filter) {
      case 'active':
        return task.status !== 'completed';
      case 'completed':
        return task.status === 'completed';
      default:
        return true;
    }
  });

  const getPriorityColor = (priority: string) => {
    switch (priority) {
      case 'high':
        return '#d13438';
      case 'low':
        return '#8a8886';
      default:
        return '#0078d4';
    }
  };

  return (
    <div>
      <h2 className="section-title">✅ Tasks</h2>

      {/* Add Task */}
      <div className="card">
        <div className="task-input-container">
          <input
            type="text"
            className="chat-input input-rounded"
            placeholder="Add a new task..."
            value={newTaskTitle}
            onChange={(e) => setNewTaskTitle(e.target.value)}
            onKeyPress={(e) => e.key === 'Enter' && handleAddTask()}
          />
          <button className="action-button primary" onClick={handleAddTask}>
            Add
          </button>
        </div>
      </div>

      {/* Action Buttons */}
      <div className="task-actions">
        <button 
          className="action-button secondary" 
          onClick={handleExtractTasks}
          disabled={loading || !currentEmail}
        >
          📧 Extract from Email
        </button>
        <button 
          className="action-button secondary" 
          onClick={handlePrioritize}
          disabled={loading || tasks.length === 0}
        >
          🎯 AI Prioritize
        </button>
      </div>

      {/* Filter Tabs */}
      <div className="tab-container mb-12">
        {(['all', 'active', 'completed'] as const).map((f) => (
          <button
            key={f}
            className={`tab-button filter-tab ${filter === f ? 'active' : ''}`}
            onClick={() => setFilter(f)}
          >
            {f.charAt(0).toUpperCase() + f.slice(1)}
          </button>
        ))}
      </div>

      {/* Tasks List */}
      {loading ? (
        <div className="loading-spinner">
          <Spinner size="small" />
        </div>
      ) : filteredTasks.length > 0 ? (
        filteredTasks.map((task) => (
          <div key={task.id} className="rule-item">
            <div className="task-item-content">
              <input
                type="checkbox"
                checked={task.status === 'completed'}
                onChange={() => handleToggleComplete(task)}
                className="task-checkbox"
                aria-label={`Mark ${task.title} as ${task.status === 'completed' ? 'incomplete' : 'complete'}`}
              />
              <div className="rule-info">
                <div 
                  className={`rule-name ${task.status === 'completed' ? 'text-strike text-muted' : ''}`}
                >
                  {task.title}
                </div>
                {task.dueDate && (
                  <div className="rule-description">
                    Due: {new Date(task.dueDate).toLocaleDateString()}
                  </div>
                )}
                <div className="task-metadata">
                  <span className={`priority-badge priority-badge--${task.priority}`}>
                    {task.priority}
                  </span>
                  {task.sourceEmailId && (
                    <span className="source-badge">
                      📧 from email
                    </span>
                  )}
                </div>
              </div>
            </div>
            <button
              onClick={() => handleDeleteTask(task.id)}
              className="delete-btn"
              aria-label={`Delete ${task.title}`}
            >
              ×
            </button>
          </div>
        ))
      ) : (
        <div className="empty-state">
          <div className="empty-state-icon">📋</div>
          <p className="empty-state-text">
            No tasks yet. Add one above or extract from an email!
          </p>
        </div>
      )}

      {/* Summary */}
      {tasks.length > 0 && (
        <div className="task-summary">
          {tasks.filter(t => t.status === 'completed').length} of {tasks.length} tasks completed
        </div>
      )}
    </div>
  );
};

export default TasksPanel;
