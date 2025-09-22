import { PublicClientApplication, AccountInfo, InteractionRequiredAuthError } from '@azure/msal-browser';

class AuthService {
  private msalInstance: PublicClientApplication;
  private account: AccountInfo | null = null;
  private initialized: boolean = false;

  constructor() {
    this.msalInstance = new PublicClientApplication({
      auth: {
        clientId: process.env.REACT_APP_CLIENT_ID!,
        authority: `https://login.microsoftonline.com/${process.env.REACT_APP_TENANT_ID}`,
        redirectUri: window.location.origin
      },
      cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
      }
    });
  }

  async initialize(): Promise<void> {
    if (this.initialized) return;
    await this.msalInstance.initialize();
    this.initialized = true;
    const accounts = this.msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      this.account = accounts[0];
    }
  }

  async getAccessToken(): Promise<string> {
    if (!this.account) {
      throw new Error('No user logged in');
    }
    try {
      const response = await this.msalInstance.acquireTokenSilent({
        scopes: [`api://${process.env.REACT_APP_CLIENT_ID}/Container.Manage`],
        account: this.account
      });
      return response.accessToken;
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        const response = await this.msalInstance.acquireTokenPopup({
          scopes: [`api://${process.env.REACT_APP_CLIENT_ID}/Container.Manage`]
        });
        return response.accessToken;
      }
      throw error;
    }
  }

  async fetchWithAuth(url: string, options: RequestInit = {}): Promise<Response> {
    await this.initialize();
    const token = await this.getAccessToken();
    return fetch(url, {
      ...options,
      headers: {
        ...options.headers,
        'Authorization': `Bearer ${token}`
      }
    });
  }
}

export const authService = new AuthService();
