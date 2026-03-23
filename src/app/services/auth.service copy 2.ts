// src/app/services/auth.service.ts
import { Injectable } from '@angular/core';
import { BehaviorSubject, Observable, from, throwError } from 'rxjs';
import { map, catchError } from 'rxjs/operators';
import { spfi, SPFI, SPFx } from '@pnp/sp';
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
import { Router } from '@angular/router';
import { environment } from '../../environments/environment';
import type { MSALOptions } from "@pnp/msaljsclient";
import { SPBrowser } from "@pnp/sp";
import { getMSAL } from "@pnp/msaljsclient";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";

const options: MSALOptions = {
  configuration: {
    auth: {
      authority: `https://login.microsoftonline.com/${environment.tenantId}/`,
      clientId: `${environment.clientId}`,
    },
    cache: {
      claimsBasedCachingEnabled: true // in order to avoid network call to refresh a token every time claims are requested
    }
  },
  authParams: {
    forceRefresh: false,
    scopes: ["https://digiservhub.sharepoint.com/.default"],
  }
};

// MSAL Configuration
export const msalConfig = {
  auth: {
    clientId: environment.clientId, // Replace with your Azure AD App Registration Client ID
    authority: 'https://login.microsoftonline.com/' + environment.tenantId, // Replace with your tenant ID
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: 'localStorage',
    storeAuthStateInCookie: false,
  }
};

export const loginRequest = {
  scopes: ['https://digiservhub.sharepoint.com/.default'],
  prompt: 'select_account' as const,
};

@Injectable({
  providedIn: 'root'
})
export class AuthService {
  private sp!: SPFI;
  private graph!: GraphFI;
  private msalInstance: IPublicClientApplication;
  private currentAccountSubject: BehaviorSubject<AccountInfo | null>;
  public currentAccount$: Observable<AccountInfo | null>;
  private isAuthenticatedSubject: BehaviorSubject<boolean>;
  public isAuthenticated$: Observable<boolean>;
  private initialized = false;
  // private msalClient: MSAL | null = null;

  constructor(private router: Router) {
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

      const response = await this.msalInstance.handleRedirectPromise();

      if (response) {
        this.handleLoginResponse(response);
      }

      const accounts = this.msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        const account = accounts[0];
        this.currentAccountSubject.next(account);
        this.isAuthenticatedSubject.next(true);
        await this.setActiveAccount(account);
      }

      this.msalInstance.addEventCallback((event) => {
        if (event.eventType === 'msal:loginSuccess' && event.payload) {
          const payload = event.payload as AuthenticationResult;
          if (payload.account) {
            this.handleLoginResponse(payload);
          }
        }
      });

    } catch (error) {
      console.error('Error initializing MSAL:', error);
    }
  }

  private handleLoginResponse(response: AuthenticationResult): void {
    if (response && response.account) {
      this.currentAccountSubject.next(response.account);
      this.isAuthenticatedSubject.next(true);
      this.setActiveAccount(response.account);
    }
  }

  private async setActiveAccount(account: AccountInfo): Promise<void> {
    this.msalInstance.setActiveAccount(account);
    await this.initializePnPWithAuth(account);
  }

  // CORRECT WAY: Initialize SP and Graph with both MSAL and Logging
  private async initializePnPWithAuth(account: AccountInfo): Promise<void> {
    // try {
    // Create MSAL middleware for token acquisition
    try {
      // Create token acquisition function
      const getToken = async (scopes: string[]): Promise<string> => {
        const request = {
          scopes: scopes,
          account: account,
        };

        try {
          const response = await this.msalInstance.acquireTokenSilent(request);
          return response.accessToken;
        } catch (error) {
          if (error instanceof InteractionRequiredAuthError) {
            const response = await this.msalInstance.acquireTokenPopup(request);
            return response.accessToken;
          }
          throw error;
        }
      };

      // Create MSAL instance for PnPjs - MSAL is a class, not an interface
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

      // const sp = spfi("https://digiservhub.sharepoint.com/sites/Pulse").using(SPBrowser(), MSAL(options));
      // const user = await sp.web.currentUser();
      // console.log("Current user:", user);
      // Initialize SP with BOTH MSAL and Logging middlewares
      // Order matters: MSAL should come first to add auth, then logging
      this.sp = spfi()
        .using(msalClient)  // First: Authentication
        .using(PnPLogging(LogLevel.Info));  // Second: Logging

      // Initialize Graph with BOTH MSAL and Logging middlewares
      this.graph = graphfi()
        .using(msalClient)  // First: Authentication
        .using(PnPLogging(LogLevel.Info));  // Second: Logging

      console.log('PnPjs initialized with MSAL auth and logging');
      // console.log(this.sp.web.getContextInfo());

    } catch (error) {
      console.error('Error initializing PnPjs with auth:', error);
    }
  }

  // Alternative: If you need to add logging to existing instances
  private addLoggingToExisting(): void {
    if (this.sp) {
      // You can add logging middleware to existing instance
      this.sp = this.sp.using(PnPLogging(LogLevel.Info));
    }
    if (this.graph) {
      this.graph = this.graph.using(PnPLogging(LogLevel.Info));
    }
  }

  login(): Observable<void> {
    if (!this.initialized) {
      return throwError(() => new Error('Auth service not initialized'));
    }

    return from(
      new Promise<void>((resolve, reject) => {
        const accounts = this.msalInstance.getAllAccounts();
        if (accounts.length > 0) {
          const account = accounts[0];
          this.currentAccountSubject.next(account);
          this.isAuthenticatedSubject.next(true);
          this.setActiveAccount(account);
          resolve();
          return;
        }

        this.msalInstance.loginRedirect(loginRequest).catch((error) => {
          console.error('Login redirect error:', error);
          reject(error);
        });
      })
    ).pipe(
      catchError((error) => {
        console.error('Login error:', error);
        return throwError(() => error);
      })
    );
  }

  logout(): Observable<void> {
    return from(
      this.msalInstance.logoutRedirect({
        postLogoutRedirectUri: msalConfig.auth.postLogoutRedirectUri,
      }).then(() => {
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

  // getCurrentUser(): Observable<any> {
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

  isInitialized(): boolean {
    return this.initialized;
  }

  isAuthenticated(): boolean {
    return this.isAuthenticatedSubject.value;
  }
}