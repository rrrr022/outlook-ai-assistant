import React from 'react';
import { createRoot } from 'react-dom/client';
import { FluentProvider, webDarkTheme } from '@fluentui/react-components';
import App from './App';
import './styles/global.css';
import './styles/retro-theme.css';

// Function to render the app
const renderApp = () => {
  const container = document.getElementById('root');
  if (container) {
    // Clear loading screen
    container.innerHTML = '';
    const root = createRoot(container);
    root.render(
      <React.StrictMode>
        <FluentProvider theme={webDarkTheme}>
          <App />
        </FluentProvider>
      </React.StrictMode>
    );
  }
};

// Track if app has been rendered
let appRendered = false;

// Safety timeout - render app after 5 seconds if Office.onReady hasn't fired
const safetyTimeout = setTimeout(() => {
  if (!appRendered) {
    console.log('Safety timeout: Rendering app without Office.onReady');
    appRendered = true;
    renderApp();
  }
}, 5000);

// Initialize Office.js if available, otherwise render directly
if (typeof Office !== 'undefined' && Office.onReady) {
  Office.onReady((info) => {
    if (!appRendered) {
      clearTimeout(safetyTimeout);
      appRendered = true;
      console.log('Office.onReady fired, host:', info.host);
      renderApp();
    }
  });
} else {
  // Office.js not loaded - render directly for standalone testing
  document.addEventListener('DOMContentLoaded', () => {
    if (!appRendered) {
      clearTimeout(safetyTimeout);
      appRendered = true;
      renderApp();
    }
  });
}
