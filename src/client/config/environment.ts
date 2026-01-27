/**
 * Environment configuration for frontend
 * Automatically detects production vs development
 */

const isLocalhost = typeof window !== 'undefined' && 
  (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1');

export const environment = {
  // API URL - uses Azure Function App in production, localhost in development
  apiUrl: isLocalhost
    ? 'https://localhost:3001'
    : 'https://outlook-ai-backend.azurewebsites.net',
  
  // Azure AD / MSAL configuration
  azure: {
    clientId: 'de86f8c9-815b-415c-94fc-b13163799862',
    authority: 'https://login.microsoftonline.com/common',
    redirectUri: isLocalhost
      ? 'https://localhost:8080/taskpane.html'
      : 'https://zealous-ground-01f6af20f.2.azurestaticapps.net/taskpane.html',
  },
  
  // Feature flags
  features: {
    enableGraphApi: true,
    enableVoiceInput: true,
    enableCalendar: true,
  },
  
  // Debugging
  debug: isLocalhost,
};

export default environment;
