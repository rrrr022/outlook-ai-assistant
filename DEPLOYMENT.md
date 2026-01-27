# Azure Deployment Guide for Outlook AI Assistant

## Prerequisites
- Azure account (free tier is fine)
- GitHub account
- Node.js 18+ installed locally

---

## Step 1: Create Azure Static Web Apps (Frontend)

1. Go to [Azure Portal](https://portal.azure.com)
2. Search for "Static Web Apps" and click "Create"
3. Fill in:
   - **Subscription:** Your subscription
   - **Resource Group:** Create new → "outlook-ai-rg"
   - **Name:** "outlook-ai-frontend" (or your choice)
   - **Plan type:** Free
   - **Region:** Choose closest to you
   - **Source:** GitHub
4. Click "Sign in with GitHub" and authorize
5. Select:
   - **Organization:** Your GitHub username
   - **Repository:** Your repo name
   - **Branch:** main
6. Build Details:
   - **Build Presets:** Custom
   - **App location:** `/`
   - **Api location:** (leave empty)
   - **Output location:** `dist/client`
7. Click "Review + Create" → "Create"

**After creation:**
- Go to the resource
- Copy the URL (e.g., `https://xxxxx.azurestaticapps.net`)
- Go to "Configuration" → Copy the deployment token for GitHub secrets

---

## Step 2: Create Azure App Service (Backend)

1. In Azure Portal, search for "App Services" and click "Create"
2. Fill in:
   - **Subscription:** Your subscription
   - **Resource Group:** Select "outlook-ai-rg"
   - **Name:** "outlook-ai-backend" (must be globally unique)
   - **Publish:** Code
   - **Runtime stack:** Node 20 LTS
   - **Operating System:** Linux (cheaper) or Windows
   - **Region:** Same as frontend
   - **Pricing plan:** Free F1
3. Click "Review + Create" → "Create"

**After creation:**
- Go to the resource
- Copy the URL (e.g., `https://outlook-ai-backend.azurewebsites.net`)
- Go to "Deployment Center" → "Manage publish profile" → Download

---

## Step 3: Configure Environment Variables

### Backend (App Service):
1. Go to your App Service → "Configuration"
2. Add Application Settings:
   ```
   GITHUB_TOKEN = your_github_token
   GITHUB_MODEL = gpt-4o
   NODE_ENV = production
   FRONTEND_URL = https://your-frontend.azurestaticapps.net
   ```
3. Click "Save"

### Frontend (Static Web Apps):
1. Go to your Static Web App → "Configuration"
2. Add Application Settings:
   ```
   REACT_APP_API_URL = https://your-backend.azurewebsites.net
   ```
3. Click "Save"

---

## Step 4: Configure GitHub Secrets

1. Go to your GitHub repo → Settings → Secrets and variables → Actions
2. Add these secrets:
   - `AZURE_STATIC_WEB_APPS_API_TOKEN` = (from Static Web App)
   - `AZURE_APP_SERVICE_NAME` = outlook-ai-backend
   - `AZURE_APP_SERVICE_PUBLISH_PROFILE` = (paste entire XML from downloaded file)
   - `API_URL` = https://your-backend.azurewebsites.net

---

## Step 5: Update Production Manifest

1. Open `manifest.production.xml`
2. Replace all `YOUR-APP.azurestaticapps.net` with your actual frontend URL
3. Replace all `YOUR-BACKEND.azurewebsites.net` with your actual backend URL
4. Generate a new GUID for the `<Id>` field at https://www.guidgenerator.com

---

## Step 6: Update Azure AD App Registration

1. Go to [Azure Portal](https://portal.azure.com) → App registrations
2. Find your app "My AI Assistant"
3. Go to "Authentication" → Add a redirect URI:
   - `https://your-frontend.azurestaticapps.net/taskpane.html`
4. Click "Save"

---

## Step 7: Deploy

### Automatic (via GitHub Actions):
Just push to the `main` branch:
```bash
git add .
git commit -m "Deploy to Azure"
git push origin main
```

### Manual (first time):
```bash
# Build locally first to test
npm run build

# Push to GitHub to trigger deployment
git push origin main
```

---

## Step 8: Set Up Keep-Alive (Optional, for Free Tier)

To prevent the backend from sleeping:

1. Go to [UptimeRobot](https://uptimerobot.com) (free)
2. Create account and add a new monitor:
   - **Monitor Type:** HTTP(s)
   - **URL:** `https://your-backend.azurewebsites.net/health`
   - **Interval:** 5 minutes
3. This will ping your backend every 5 minutes to keep it warm

---

## Step 9: Publish the Add-in

### For personal/testing:
1. In Outlook Web, go to Settings → Manage Add-ins → Custom Add-ins
2. Upload `manifest.production.xml`

### For organization:
1. Go to Microsoft 365 Admin Center
2. Settings → Integrated apps → Upload custom apps
3. Upload `manifest.production.xml`

### For public (AppSource):
1. Go to [Partner Center](https://partner.microsoft.com)
2. Submit your add-in for review

---

## Troubleshooting

### Backend returns 500 error
- Check App Service logs: App Service → "Log stream"
- Verify environment variables are set correctly

### Frontend shows blank page
- Check browser console for errors
- Verify the API_URL is correct

### Auth popup fails
- Verify the redirect URI matches exactly in Azure AD
- Check the client ID is correct in the code

### CORS errors
- Update `productionConfig.ts` with your frontend URL
- Redeploy the backend

---

## Cost Summary (Free Tier)

| Service | Cost |
|---------|------|
| Static Web Apps | $0 |
| App Service (F1) | $0 |
| UptimeRobot | $0 |
| **Total** | **$0/month** |

**Limitations:**
- Backend sleeps after 20 min (UptimeRobot prevents this)
- 60 min CPU/day on backend
- 100 GB bandwidth/month on frontend

For production with more than ~50 users, consider upgrading to Basic ($13/month).
