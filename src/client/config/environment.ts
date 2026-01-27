/**
 * Environment configuration for frontend
 * Automatically detects production vs development
 */

const isProduction = process.env.NODE_ENV === 'production';
const isLocalhost = typeof window !== 'undefined' && window.location.hostname === 'localhost';

export const environment = {
  // API URL - uses environment variable in production, localhost in development
  apiUrl: isProduction && !isLocalhost
    ? (process.env.REACT_APP_API_URL || 'https://your-backend.azurewebsites.net')
    : 'https://localhost:3001',
  
  // Azure AD / MSAL configuration
  azure: {
    clientId: 'de86f8c9-815b-415c-94fc-b13163799862',
    authority: 'https://login.microsoftonline.com/common',
    redirectUri: isProduction && !isLocalhost
      ? (process.env.REACT_APP_REDIRECT_URI || 'https://your-app.azurestaticapps.net/taskpane.html')
      : 'https://localhost:8080/taskpane.html',
  },
  
  // Feature flags
  features: {
    enableGraphApi: true,
    enableVoiceInput: true,
    enableCalendar: true,
  },
  
  // Debugging
  debug: !isProduction || isLocalhost,
};

export default environment;
