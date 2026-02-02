import React from 'react';
import { useAppStore } from '../store/appStore';

const HistoryPanel: React.FC = () => {
  const { sessions, activeSessionId, loadSession, deleteSession, createNewSession } = useAppStore();

  return (
    <div style={{ padding: '12px' }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <h3 style={{ margin: 0 }}>Chat History</h3>
        <button
          onClick={createNewSession}
          style={{
            padding: '6px 10px',
            fontSize: '12px',
            borderRadius: '6px',
            border: '1px solid #d0d0d0',
            background: '#fff',
            cursor: 'pointer',
          }}
        >
          New Chat
        </button>
      </div>

      {sessions.length === 0 && (
        <p style={{ color: '#666', marginTop: '12px' }}>No saved sessions yet.</p>
      )}

      <ul style={{ listStyle: 'none', padding: 0, marginTop: '12px' }}>
        {sessions.map((session) => (
          <li
            key={session.id}
            style={{
              padding: '10px',
              border: '1px solid #e0e0e0',
              borderRadius: '8px',
              marginBottom: '8px',
              background: session.id === activeSessionId ? '#f3f9ff' : '#fff',
            }}
          >
            <div style={{ fontWeight: 600, marginBottom: '4px' }}>
              {session.title || 'Untitled'}
            </div>
            <div style={{ fontSize: '12px', color: '#666', marginBottom: '8px' }}>
              {new Date(session.updatedAt).toLocaleString()}
            </div>
            <div style={{ display: 'flex', gap: '8px' }}>
              <button
                onClick={() => loadSession(session.id)}
                style={{
                  padding: '6px 10px',
                  fontSize: '12px',
                  borderRadius: '6px',
                  border: '1px solid #d0d0d0',
                  background: '#fff',
                  cursor: 'pointer',
                }}
              >
                Open
              </button>
              <button
                onClick={() => deleteSession(session.id)}
                style={{
                  padding: '6px 10px',
                  fontSize: '12px',
                  borderRadius: '6px',
                  border: '1px solid #f2b8b5',
                  background: '#fff',
                  color: '#a4262c',
                  cursor: 'pointer',
                }}
              >
                Delete
              </button>
            </div>
          </li>
        ))}
      </ul>
    </div>
  );
};

export default HistoryPanel;
