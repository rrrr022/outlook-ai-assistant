import environment from '../config/environment';

// Type definitions for MSAL (avoid importing the actual module at top level)
interface MsalConfiguration {
  auth: { clientId: string; authority: string; redirectUri: string };
  cache: { cacheLocation: string };
  system?: { loggerOptions?: { loggerCallback?: (level: number, message: string, containsPii: boolean) => void } };
}

// Declare Office types
declare const Office: any;

// Lazy load MSAL with webpack magic comment for true code splitting
let msalModule: any = null;
async function loadMsal() {
  if (!msalModule) {
    msalModule = await import(/* webpackChunkName: "msal-lib" */ '@azure/msal-browser');
  }
  return msalModule;
}

// Your Azure AD App Registration
const AZURE_CLIENT_ID = environment.azure.clientId;
const REDIRECT_URI = environment.azure.redirectUri;

// Auth dialog URL
const getAuthDialogUrl = () => {
  const origin = typeof window !== 'undefined' ? window.location.origin : '';
  return `${origin}/auth-dialog.html`;
};

const msalConfig: MsalConfiguration = {
  auth: {
    clientId: AZURE_CLIENT_ID,
    authority: environment.azure.authority,
    redirectUri: REDIRECT_URI,
  },
  cache: {
    cacheLocation: 'localStorage',
  },
  system: {
    loggerOptions: {
      loggerCallback: (level: number, message: string, containsPii: boolean) => {
        if (!containsPii && environment.debug) {
          console.log('[MSAL]', message);
        }
      },
    },
  },
};

// Microsoft Graph permissions
const graphScopes = ['User.Read', 'Mail.Read', 'Mail.ReadWrite', 'Calendars.Read'];

class MsalService {
  private msalInstance: any = null;
  private initialized: boolean = false;
  private initPromise: Promise<void> | null = null;

  /**
   * Initialize MSAL - must be called before any auth operations
   */
  async initialize(): Promise<void> {
    if (this.initialized) return;
    
    if (this.initPromise) {
      return this.initPromise;
    }

    this.initPromise = (async () => {
      const { PublicClientApplication } = await loadMsal();
      this.msalInstance = new PublicClientApplication(msalConfig);
      await this.msalInstance.initialize();
      this.initialized = true;
      console.log('✅ MSAL initialized successfully');
      
      // Handle redirect callback
      const response = await this.msalInstance.handleRedirectPromise();
      if (response) {
        console.log('✅ Login successful via redirect');
      }
    })();

    return this.initPromise;
  }

  /**
   * Check if user is signed in
   */
  isSignedIn(): boolean {
    if (!this.msalInstance) return false;
    const accounts = this.msalInstance.getAllAccounts();
    return accounts.length > 0;
  }

  /**
   * Get current account
   */
  getAccount() {
    if (!this.msalInstance) return null;
    const accounts = this.msalInstance.getAllAccounts();
    return accounts.length > 0 ? accounts[0] : null;
  }

  /**
   * Sign in using Office Dialog API (recommended for Office Add-ins)
   */
  async signIn(): Promise<any | null> {
    await this.initialize();
    if (!this.msalInstance) throw new Error('MSAL not initialized');

    // Check if we're in Office context and can use Dialog API
    if (typeof Office !== 'undefined' && Office.context && Office.context.ui) {
      return this.signInWithOfficeDialog();
    }
    
    // Fallback to popup for non-Office contexts (e.g., local testing)
    return this.signInWithPopup();
  }

  /**
   * Sign in using Office Dialog API
   */
  private async signInWithOfficeDialog(): Promise<any | null> {
    return new Promise((resolve, reject) => {
      const dialogUrl = getAuthDialogUrl();
      console.log('📤 Opening auth dialog:', dialogUrl);
      
      Office.context.ui.displayDialogAsync(
        dialogUrl,
        { 
          height: 60, 
          width: 40,
          promptBeforeOpen: false
        },
        (asyncResult: any) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error('❌ Failed to open dialog:', asyncResult.error.message);
            reject(new Error(asyncResult.error.message));
            return;
          }

          const dialog = asyncResult.value;
          console.log('📤 Auth dialog opened');

          // Handle messages from the dialog
          dialog.addEventHandler(
            Office.EventType.DialogMessageReceived,
            async (arg: any) => {
              dialog.close();
              
              try {
                const message = JSON.parse(arg.message);
                console.log('📨 Received message from dialog:', message.status);
                
                if (message.status === 'success') {
                  // Reinitialize MSAL to pick up the new cached tokens
                  this.msalInstance = null;
                  this.initialized = false;
                  this.initPromise = null;
                  await this.initialize();
                  
                  console.log('✅ Login successful:', message.account?.username);
                  resolve({
                    account: message.account,
                    accessToken: message.accessToken
                  });
                } else {
                  console.error('❌ Login failed:', message.error);
                  reject(new Error(message.error || 'Authentication failed'));
                }
              } catch (error) {
                console.error('❌ Error parsing dialog message:', error);
                reject(error);
              }
            }
          );

          // Handle dialog closed by user
          dialog.addEventHandler(
            Office.EventType.DialogEventReceived,
            (arg: any) => {
              console.log('📤 Dialog event:', arg.error);
              if (arg.error === 12006) {
                // Dialog closed by user
                reject(new Error('User cancelled the sign-in'));
              }
            }
          );
        }
      );
    });
  }

  /**
   * Sign in with popup (fallback for non-Office contexts)
   */
  private async signInWithPopup(): Promise<any | null> {
    const request = {
      scopes: graphScopes,
      prompt: 'select_account',
    };

    try {
      const response = await this.msalInstance.loginPopup(request);
      console.log('✅ Login successful:', response.account?.username);
      return response;
    } catch (error) {
      console.error('❌ Login failed:', error);
      throw error;
    }
  }

  /**
   * Sign out
   */
  async signOut(): Promise<void> {
    await this.initialize();
    if (!this.msalInstance) return;
    
    const account = this.getAccount();
    if (account) {
      await this.msalInstance.logoutPopup({
        account,
        postLogoutRedirectUri: REDIRECT_URI,
      });
    }
  }

  /**
   * Get access token for Microsoft Graph
   */
  async getAccessToken(): Promise<string | null> {
    await this.initialize();
    if (!this.msalInstance) return null;

    const account = this.getAccount();
    if (!account) {
      console.log('No account found, need to sign in first');
      return null;
    }

    const silentRequest = {
      scopes: graphScopes,
      account,
    };

    try {
      // Try silent token acquisition first
      const response = await this.msalInstance.acquireTokenSilent(silentRequest);
      console.log('✅ Got token silently');
      return response.accessToken;
    } catch (silentError) {
      console.log('Silent token acquisition failed, trying popup');
      
      try {
        // Fall back to popup
        const response = await this.msalInstance.acquireTokenPopup({
          scopes: graphScopes,
        });
        console.log('✅ Got token via popup');
        return response.accessToken;
      } catch (popupError) {
        console.error('❌ Token acquisition failed:', popupError);
        return null;
      }
    }
  }

  /**
   * Get user info
   */
  getUserInfo(): { name: string; email: string } | null {
    const account = this.getAccount();
    if (!account) return null;

    return {
      name: account.name || account.username || 'User',
      email: account.username || '',
    };
  }
}

// Export singleton instance
export const msalService = new MsalService();
