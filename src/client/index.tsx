import React from 'react';
import { createRoot } from 'react-dom/client';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import App from './App';
import './styles/global.css';

// Function to render the app
const renderApp = () => {
  const container = document.getElementById('root');
  if (container) {
    const root = createRoot(container);
    root.render(
      <React.StrictMode>
        <FluentProvider theme={webLightTheme}>
          <App />
        </FluentProvider>
      </React.StrictMode>
    );
  }
};

// Initialize Office.js if available, otherwise render directly for browser testing
if (typeof Office !== 'undefined' && Office.onReady) {
  Office.onReady((info) => {
    // Render the app regardless of host type for development/testing
    // In production, Office.context.mailbox will handle Outlook-specific features
    renderApp();
  });
} else {
  // Office.js not loaded - render directly for standalone testing
  document.addEventListener('DOMContentLoaded', renderApp);
}
