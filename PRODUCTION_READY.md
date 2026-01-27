# Outlook AI Assistant - Production Deployment Complete! 🎉

## Your Production URLs

| Service | URL |
|---------|-----|
| **Frontend (Static Web App)** | https://zealous-ground-01f6af20f.2.azurestaticapps.net |
| **Taskpane** | https://zealous-ground-01f6af20f.2.azurestaticapps.net/taskpane.html |
| **Backend API** | https://outlook-ai-backend.azurewebsites.net |
| **Health Check** | https://outlook-ai-backend.azurewebsites.net/api/health |

## Azure Resources Created

| Resource | Name | Type |
|----------|------|------|
| Resource Group | outlook-ai-rg | Microsoft.Resources |
| Frontend | outlook-ai-frontend | Azure Static Web Apps |
| Backend | outlook-ai-backend | Azure Functions (Consumption) |
| Storage | outlookaist2026 | Storage Account |
| App Insights | outlook-ai-backend | Application Insights |

## Using the Production Add-in

### Option 1: Install in Outlook Web (Recommended for Testing)

1. Go to https://outlook.office.com
2. Click the **gear icon** → **View all Outlook settings**
3. Go to **Mail** → **Customize actions** → **Add-ins**
4. Click **My add-ins** → **Add a custom add-in** → **Add from URL**
5. Enter: `https://zealous-ground-01f6af20f.2.azurestaticapps.net/manifest.azure.xml`

### Option 2: Admin Deployment (for organization-wide use)

1. Go to [Microsoft 365 Admin Center](https://admin.microsoft.com)
2. Navigate to **Settings** → **Integrated apps**
3. Click **Upload custom apps**
4. Upload the `manifest.azure.xml` file from your project

## Production Manifest

The production manifest is at: `manifest.azure.xml`

This manifest points to the Azure-hosted URLs and can be used for both testing and production deployment.

## GitHub Actions (Auto-Deployment)

Your repository now has GitHub Actions configured for automatic deployments:

- **Frontend**: Automatically deploys when you push to `master` branch
- **Backend**: Deploy manually using Azure CLI (see below)

### Redeploy Backend
```bash
cd api
Compress-Archive -Path * -DestinationPath function-app.zip -Force
az functionapp deployment source config-zip --name outlook-ai-backend --resource-group outlook-ai-rg --src function-app.zip
```

### Redeploy Frontend (manual)
```bash
gh workflow run azure-static-web-apps.yml --repo rrrr022/outlook-ai-assistant
```

## Configuration

### Environment Variables (Backend)
These are already configured in Azure:
- `GITHUB_TOKEN` - Your GitHub Models API token
- `GITHUB_MODEL` - gpt-4o
- `FRONTEND_URL` - https://zealous-ground-01f6af20f.2.azurestaticapps.net
- `NODE_ENV` - production

### Azure AD App Registration
- **Client ID**: de86f8c9-815b-415c-94fc-b13163799862
- **Redirect URIs**: 
  - https://localhost:8080/taskpane.html (development)
  - https://zealous-ground-01f6af20f.2.azurestaticapps.net/taskpane.html (production)

## Full Inbox Access (Microsoft Graph)

⚠️ **Note**: Full inbox access requires admin consent for your organization.

If you need to access the entire inbox (not just the current email), your IT admin needs to approve these permissions:
- User.Read
- Mail.Read
- Mail.ReadWrite
- Calendars.Read

Until approved, the add-in works with the currently selected email using Office.js APIs.

## Costs

- **Azure Static Web Apps**: Free tier (100GB bandwidth/month)
- **Azure Functions**: Consumption plan (1 million executions/month free)
- **GitHub Models**: Rate limited but free for development

## Local Development

To continue developing locally:
```bash
npm run dev
```

This starts both frontend (https://localhost:8080) and backend (https://localhost:3001) with live reload.
