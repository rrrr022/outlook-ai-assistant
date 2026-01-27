/**
 * Production configuration for Azure deployment
 */

export const productionConfig = {
  // Server settings
  port: process.env.PORT || 3001,
  
  // CORS settings - update these with your actual domains
  corsOrigins: [
    // Azure Static Web Apps domain (update after deployment)
    process.env.FRONTEND_URL || 'https://your-app.azurestaticapps.net',
    // Office Online domains
    'https://outlook.office.com',
    'https://outlook.office365.com',
    'https://outlook.live.com',
    // Local development
    'https://localhost:8080',
  ],
  
  // API settings
  apiPrefix: '/api',
  
  // Security
  trustProxy: true, // Required for Azure App Service
  
  // Logging
  logLevel: process.env.LOG_LEVEL || 'info',
};

export default productionConfig;
