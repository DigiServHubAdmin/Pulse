// src/app/services/auth.service.ts
import { Injectable } from '@angular/core';
import { BehaviorSubject, Observable, from, throwError } from 'rxjs';
import { map, catchError } from 'rxjs/operators';
import { spfi, SPFI } from '@pnp/sp';
import { graphfi, GraphFI } from '@pnp/graph';
import { LogLevel, PnPLogging } from '@pnp/logging';
import { MSAL } from '@pnp/msaljsclient';
import {
  PublicClientApplication,
  AccountInfo,
  AuthenticationResult,
  InteractionRequiredAuthError,
  BrowserCacheLocation,
  IPublicClientApplication,
  Configuration,
  LogLevel as MsalLogLevel
} from '@azure/msal-browser';
import { environment } from '../../environments/environment';

// MSAL Configuration
export const msalConfig = {
  auth: {
    clientId: environment.msal.clientId,
    authority: environment.msal.authority,
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: 'localStorage',
    // cacheLocation: BrowserCacheLocation.LocalStorage,
    storeAuthStateInCookie: false,
  },
      
  // system: {
  //   loggerOptions: {
  //     loggerCallback: (level: any, message: string, containsPii: boolean) => {
  //       if (containsPii) return;
  //       console.log(`MSAL Log: ${message}`);
  //     },
  //     logLevel: LogLevel.Info,
  //     piiLoggingEnabled: false
  //   },
  //   windowHashTimeout: 30000,
  // }
};

// MSAL Request for SharePoint
export const loginRequest = {
  scopes: ['https://digiservhub.sharepoint.com/.default'],
  prompt: 'select_account' as const,
};

// MSAL Request for Microsoft Graph
export const graphRequest = {
  scopes: ['User.Read', 'Mail.Read', 'Sites.Read.All'],
};

@Injectable({
  providedIn: 'root'
})
export class AuthService {
  private sp!: SPFI;
  private graph!: GraphFI;
  private msalInstance: PublicClientApplication;
  private currentAccountSubject: BehaviorSubject<AccountInfo | null>;
  public currentAccount$: Observable<AccountInfo | null>;
  private isAuthenticatedSubject: BehaviorSubject<boolean>;
  public isAuthenticated$: Observable<boolean>;
  private initialized = false;

  constructor() {
    this.msalInstance = new PublicClientApplication(msalConfig);
    this.currentAccountSubject = new BehaviorSubject<AccountInfo | null>(null);
    this.currentAccount$ = this.currentAccountSubject.asObservable();
    this.isAuthenticatedSubject = new BehaviorSubject<boolean>(false);
    this.isAuthenticated$ = this.isAuthenticatedSubject.asObservable();

    this.initializeAuth();
  }

  private async initializeAuth(): Promise<void> {
    try {
      await this.msalInstance.initialize();
      this.initialized = true;
      
      // Handle redirect promise
      const response = await this.msalInstance.handleRedirectPromise();
      if (response) {
        console.log('Redirect response received:', response);
        await this.handleLoginResponse(response);
      }
      
      // Set active account
      const accounts = this.msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        const account = accounts[0];
        this.currentAccountSubject.next(account);
        this.isAuthenticatedSubject.next(true);
        await this.setActiveAccount(account);
      }
      
      // Add event callback for login success
      this.msalInstance.addEventCallback(async (event) => {
        if (event.eventType === 'msal:loginSuccess' && event.payload) {
          console.log('Login success event received');
          const payload = event.payload as AuthenticationResult;
          if (payload.account) {
            await this.handleLoginResponse(payload);
          }
        }
        
        // if (event.eventType === 'msal:loginFailure') {
        //   console.error('Login failure event:', event.error);
        // }
      });
      
    } catch (error) {
      console.error('Error initializing MSAL:', error);
    }
  }

  private async handleLoginResponse(response: AuthenticationResult): Promise<void> {
    if (response && response.account) {
      this.currentAccountSubject.next(response.account);
      this.isAuthenticatedSubject.next(true);
      await this.setActiveAccount(response.account);
    }
  }

  private async setActiveAccount(account: AccountInfo): Promise<void> {
    this.msalInstance.setActiveAccount(account);
    await this.initializePnPWithAuth(account);
  }

  private async initializePnPWithAuth(account: AccountInfo): Promise<void> {
    try {
      // Create token acquisition function
      const getToken = async (scopes: string[]): Promise<string> => {
        const request = {
          scopes: scopes,
          account: account,
        };
        
        try {
          console.log('Acquiring token silently for scopes:', scopes);
          const response = await this.msalInstance.acquireTokenSilent(request);
          console.log('Token acquired successfully');
          return response.accessToken;
        } catch (error) {
          console.error('Silent token acquisition failed:', error);
          if (error instanceof InteractionRequiredAuthError) {
            console.log('Interaction required, trying popup...');
            try {
              const response = await this.msalInstance.acquireTokenPopup(request);
              console.log('Token acquired via popup');
              return response.accessToken;
            } catch (popupError) {
              console.error('Popup token acquisition failed:', popupError);
              throw popupError;
            }
          }
          throw error;
        }
      };

      // Create MSAL client for PnPjs - CORRECT WAY
      // const msalClient = new MSAL({
      //   auth: {
      //     clientId: msalConfig.auth.clientId,
      //     authority: msalConfig.auth.authority,
      //   },
      //   getToken: getToken
      // });
      const msalClient = MSAL({
        configuration: msalConfig,
        authParams: loginRequest,
      });

      // const msalClient = MSAL({
      //   auth: {
      //     clientId: msalConfig.auth.clientId,
      //     authority: msalConfig.auth.authority,
      //   },
      //   getToken: getToken  // Function that returns token
      // });
            

      // Initialize SharePoint client with MSAL and logging
      this.sp = spfi()
        .using(msalClient)
        .using(PnPLogging(LogLevel.Info));
      
      // Initialize Graph client with MSAL and logging
      this.graph = graphfi()
        .using(msalClient)
        .using(PnPLogging(LogLevel.Info));
      
      console.log('PnPjs initialized successfully with authentication');
      
    } catch (error) {
      console.error('Error initializing PnPjs with auth:', error);
      throw error;
    }
  }

  // FIXED: Login with popup - properly handles closing
  login(): Observable<AuthenticationResult> {
    return new Observable<AuthenticationResult>((observer) => {
      // Check if already authenticated
      const accounts = this.msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        console.log('Already authenticated');
        const account = accounts[0];
        this.currentAccountSubject.next(account);
        this.isAuthenticatedSubject.next(true);
        this.setActiveAccount(account).then(() => {
          observer.next({ account } as AuthenticationResult);
          observer.complete();
        }).catch((error) => {
          observer.error(error);
        });
        return;
      }

      // Start login popup
      console.log('Starting login popup...');
      this.msalInstance.loginPopup(loginRequest)
        .then(async (response) => {
          console.log('Login popup completed successfully');
          if (response && response.account) {
            this.currentAccountSubject.next(response.account);
            this.isAuthenticatedSubject.next(true);
            await this.setActiveAccount(response.account);
            observer.next(response);
            observer.complete();
          } else {
            observer.error(new Error('No account in login response'));
          }
        })
        .catch((error) => {
          console.error('Login popup error:', error);
          // Handle specific error cases
          if (error.errorCode === 'user_cancelled') {
            console.log('User cancelled login');
            observer.error(new Error('Login cancelled by user'));
          } else if (error.errorCode === 'popup_window_error') {
            console.error('Popup window error - possibly blocked');
            observer.error(new Error('Popup blocked. Please allow popups for this site.'));
          } else {
            observer.error(error);
          }
        });
    }).pipe(
      catchError((error) => {
        console.error('Login error:', error);
        return throwError(() => error);
      })
    );
  }

  // Alternative: Login with redirect (more reliable)
  loginWithRedirect(): Observable<void> {
    return from(
      new Promise<void>((resolve, reject) => {
        const accounts = this.msalInstance.getAllAccounts();
        if (accounts.length > 0) {
          console.log('Already authenticated');
          const account = accounts[0];
          this.currentAccountSubject.next(account);
          this.isAuthenticatedSubject.next(true);
          this.setActiveAccount(account).then(() => resolve()).catch(reject);
          return;
        }

        console.log('Starting login redirect...');
        this.msalInstance.loginRedirect(loginRequest).catch((error) => {
          console.error('Login redirect error:', error);
          reject(error);
        });
        
        // Note: The page will redirect, so we don't resolve here
      })
    ).pipe(
      catchError((error) => {
        console.error('Login redirect error:', error);
        return throwError(() => error);
      })
    );
  }

  logout(): Observable<void> {
    return from(
      this.msalInstance.logoutPopup({
        postLogoutRedirectUri: msalConfig.auth.postLogoutRedirectUri,
        mainWindowRedirectUri: msalConfig.auth.redirectUri,
      }).then(() => {
        console.log('Logout successful');
        this.currentAccountSubject.next(null);
        this.isAuthenticatedSubject.next(false);
      })
    ).pipe(
      catchError((error) => {
        console.error('Logout error:', error);
        return throwError(() => error);
      })
    );
  }

  // FIXED: Get current user from Graph
  // getCurrentUser(): Observable<any> {
  //   if (!this.graph) {
  //     return throwError(() => new Error('Graph client not initialized. Please authenticate first.'));
  //   }
    
  //   return from(this.graph.me()).pipe(
  //     map((user) => user),
  //     catchError((error) => {
  //       console.error('Error getting current user:', error);
  //       return throwError(() => error);
  //     })
  //   );
  // }

  getSharePoint(): SPFI {
    if (!this.sp) {
      throw new Error('SharePoint client not initialized. Please authenticate first.');
    }
    return this.sp;
  }

  getGraph(): GraphFI {
    if (!this.graph) {
      throw new Error('Graph client not initialized. Please authenticate first.');
    }
    return this.graph;
  }

  getAccessToken(scopes?: string[]): Observable<string> {
    const account = this.msalInstance.getActiveAccount();
    if (!account) {
      return throwError(() => new Error('No active account found'));
    }

    const tokenScopes = scopes || loginRequest.scopes;
    
    return from(
      this.msalInstance.acquireTokenSilent({
        scopes: tokenScopes,
        account: account,
      }).then((result) => result.accessToken)
    ).pipe(
      catchError((error) => {
        console.error('Error getting access token:', error);
        return throwError(() => error);
      })
    );
  }

  isInitialized(): boolean {
    return this.initialized;
  }

  isAuthenticated(): boolean {
    return this.isAuthenticatedSubject.value;
  }

  getCurrentAccount(): AccountInfo | null {
    return this.currentAccountSubject.value;
  }
}