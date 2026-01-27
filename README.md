# Outlook AI Assistant

An AI-powered Outlook Add-in for intelligent email management, calendar scheduling, and task automation.

## Features

- 🤖 **AI Chat** - Natural language interface to manage emails, calendar, and tasks
- 📧 **Smart Email** - AI-assisted email composition, replies, and summaries
- 📅 **Calendar Management** - Schedule meetings, manage appointments, find available time slots
- ✅ **Task Automation** - Create, track, and automate task management
- 📊 **Analytics Dashboard** - Track productivity metrics and email patterns
- 📝 **Email Templates** - Quick access to AI-generated email templates
- 🎤 **Voice Input** - Dictate commands and emails using voice
- 🛡️ **Approval System** - Review AI actions before execution

## Tech Stack

- **Frontend**: React 18, TypeScript, Fluent UI, Zustand
- **Backend**: Node.js, Express, TypeScript
- **Office Integration**: Office.js API
- **AI Providers**: GitHub Models API (primary), OpenAI, Anthropic

## Quick Start

### Prerequisites

- Node.js 18+ and npm
- Outlook Desktop (Windows or Mac) or Outlook Web
- A GitHub account (for GitHub Models API)

### Installation

1. **Clone and install dependencies:**
   ```bash
   git clone <repository-url>
   cd OutlookAi
   npm install
   ```

2. **Configure environment variables:**
   Create a `.env` file in the root directory:
   ```env
   GITHUB_TOKEN=your_github_personal_access_token
   OPENAI_API_KEY=your_openai_key_optional
   ANTHROPIC_API_KEY=your_anthropic_key_optional
   ```
   
   > **Note**: GitHub Models API is the primary AI provider. Generate a GitHub Personal Access Token with no special permissions required.

3. **Start the development server:**
   ```bash
   npm run dev
   ```
   
   This starts:
   - Frontend: https://localhost:8080
   - Backend: http://localhost:3001

## Sideloading the Add-in in Outlook

### Step 1: Trust the Self-Signed Certificate

Before sideloading, you need to trust the development certificate:

1. Open your browser and visit: https://localhost:8080
2. You'll see a security warning about the certificate
3. Click **Advanced** → **Proceed to localhost (unsafe)**
4. This allows Outlook to load the add-in from localhost

### Step 2: Sideload in Outlook Desktop (Windows)

#### Option A: Via Get Add-ins Dialog (Easiest)

1. Open **Outlook Desktop**
2. Click on any email to open it in the reading pane
3. In the ribbon, click **Home** → **Get Add-ins** (or **File** → **Get Add-ins**)
4. In the Add-ins dialog, click **My add-ins** on the left
5. Scroll down to the **Custom Addins** section
6. Click **+ Add a custom add-in** → **Add from file...**
7. Navigate to your project folder and select `manifest.xml`
8. Click **Install** when prompted

#### Option B: Via Centralized Deployment (Admin)

If you have Microsoft 365 admin access:
1. Go to Microsoft 365 admin center → **Settings** → **Integrated apps**
2. Click **Upload custom apps**
3. Upload your manifest.xml

### Step 3: Sideload in Outlook on the Web

1. Go to https://outlook.office.com
2. Open any email
3. Click the **...** (More actions) menu in the email toolbar
4. Select **Get Add-ins**
5. Click **My add-ins** in the sidebar
6. Scroll to **Custom Addins** section
7. Click **+ Add a custom add-in** → **Add from file...**
8. Upload your `manifest.xml`
9. Confirm the installation

### Step 4: Using the Add-in

Once installed:

1. Open any email in Outlook
2. Look for the **AI Assistant** button:
   - **Desktop**: In the ribbon under the Home tab
   - **Web**: In the email toolbar (may be under "..." menu)
3. Click to open the task pane on the right side
4. Use the tabs to access different features:
   - **Chat**: AI-powered conversation
   - **Email**: Compose and analyze emails
   - **Calendar**: View and manage events
   - **Tasks**: Track action items
   - **Analytics**: View productivity stats
   - **Settings**: Configure AI and approval settings

## Project Structure

```
OutlookAi/
├── manifest.xml              # Office Add-in manifest
├── src/
│   ├── client/               # React frontend
│   │   ├── components/       # UI components
│   │   │   ├── ChatPanel.tsx
│   │   │   ├── EmailPanel.tsx
│   │   │   ├── CalendarPanel.tsx
│   │   │   ├── TasksPanel.tsx
│   │   │   ├── AnalyticsPanel.tsx
│   │   │   ├── SettingsPanel.tsx
│   │   │   ├── ApprovalModal.tsx
│   │   │   └── TemplatesModal.tsx
│   │   ├── services/         # API & Office.js services
│   │   │   ├── aiService.ts
│   │   │   ├── outlookService.ts
│   │   │   └── approvalService.ts
│   │   ├── store/            # Zustand state management
│   │   └── styles/           # CSS stylesheets
│   ├── server/               # Express backend
│   │   └── services/ai/      # AI provider implementations
│   │       ├── openaiProvider.ts
│   │       ├── anthropicProvider.ts
│   │       └── githubModelsProvider.ts
│   └── shared/               # Shared TypeScript types
└── scripts/                  # Build utilities
```

## Development Scripts

| Command | Description |
|---------|-------------|
| `npm run dev` | Start both frontend and backend dev servers |
| `npm run dev:client` | Start only the frontend (webpack) |
| `npm run dev:server` | Start only the backend (Express) |
| `npm run build` | Build for production |
| `npm start` | Start production server |

## Configuration

### AI Providers

Configure in Settings → AI Provider:
- **GitHub Models** (Default): Uses your GitHub token - no additional cost
- **OpenAI**: Requires OPENAI_API_KEY in .env
- **Anthropic**: Requires ANTHROPIC_API_KEY in .env

### Approval Settings

For safety, the add-in can require approval before:
- Sending emails on your behalf
- Creating calendar meetings
- Executing automation rules

Configure in Settings → Approval Controls.

## Troubleshooting

### Add-in doesn't appear in Outlook

1. Verify servers are running:
   - Frontend: https://localhost:8080/taskpane.html
   - Backend: http://localhost:3001/health
2. Clear the Office cache:
   - Windows: Delete contents of `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
   - Then restart Outlook
3. Re-sideload the manifest

### "This add-in could not be started" error

1. Visit https://localhost:8080 in your browser and accept the certificate
2. Clear browser cookies for localhost
3. Make sure both servers are running

### AI not responding

1. Check that GITHUB_TOKEN is set in `.env`
2. Verify the token is valid on GitHub
3. Check browser console (F12) for errors
4. Verify backend is running: http://localhost:3001/health

### Manifest validation errors

1. Ensure all URLs in manifest.xml use `https://localhost:8080`
2. Check that icon files exist in `src/client/assets/`
3. Validate manifest at: https://aka.ms/validatemanifest

## Security Notes

- **API Keys**: Never commit `.env` files. The backend keeps keys server-side.
- **Approval System**: Enable all approval settings for production use
- **HTTPS**: Required for Office Add-ins. Dev server auto-generates certificates.

## License

MIT License - See LICENSE file for details.
