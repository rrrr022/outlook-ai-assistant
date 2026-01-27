import { PublicClientApplication, Configuration, AuthenticationResult, SilentRequest, PopupRequest } from '@azure/msal-browser';
import environment from '../config/environment';

// Your Azure AD App Registration
const AZURE_CLIENT_ID = environment.azure.clientId;

const msalConfig: Configuration = {
  auth: {
    clientId: AZURE_CLIENT_ID,
    authority: environment.azure.authority,
    redirectUri: environment.azure.redirectUri,
  },
  cache: {
    cacheLocation: 'localStorage',
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
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
  private msalInstance: PublicClientApplication;
  private initialized: boolean = false;
  private initPromise: Promise<void> | null = null;

  constructor() {
    this.msalInstance = new PublicClientApplication(msalConfig);
  }

  /**
   * Initialize MSAL - must be called before any auth operations
   */
  async initialize(): Promise<void> {
    if (this.initialized) return;
    
    if (this.initPromise) {
      return this.initPromise;
    }

    this.initPromise = this.msalInstance.initialize().then(() => {
      this.initialized = true;
      console.log('✅ MSAL initialized successfully');
      
      // Handle redirect callback
      return this.msalInstance.handleRedirectPromise().then((response) => {
        if (response) {
          console.log('✅ Login successful via redirect');
        }
      });
    });

    return this.initPromise;
  }

  /**
   * Check if user is signed in
   */
  isSignedIn(): boolean {
    const accounts = this.msalInstance.getAllAccounts();
    return accounts.length > 0;
  }

  /**
   * Get current account
   */
  getAccount() {
    const accounts = this.msalInstance.getAllAccounts();
    return accounts.length > 0 ? accounts[0] : null;
  }

  /**
   * Sign in with popup
   */
  async signIn(): Promise<AuthenticationResult | null> {
    await this.initialize();

    const request: PopupRequest = {
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
    
    const account = this.getAccount();
    if (account) {
      await this.msalInstance.logoutPopup({
        account,
        postLogoutRedirectUri: 'https://localhost:8080/taskpane.html',
      });
    }
  }

  /**
   * Get access token for Microsoft Graph
   */
  async getAccessToken(): Promise<string | null> {
    await this.initialize();

    const account = this.getAccount();
    if (!account) {
      console.log('No account found, need to sign in first');
      return null;
    }

    const silentRequest: SilentRequest = {
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
