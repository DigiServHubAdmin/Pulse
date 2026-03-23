// src/app/services/auth.service.ts
import { Injectable, Inject } from '@angular/core';
import { BehaviorSubject, Observable, from, throwError } from 'rxjs';
import { map, catchError, switchMap } from 'rxjs/operators';
import { spfi, SPFI, SPFx } from '@pnp/sp';
import { graphfi, GraphFI } from '@pnp/graph';
import { LogLevel, PnPLogging } from '@pnp/logging';
import { MSAL } from '@pnp/msaljsclient';
import { PublicClientApplication, InteractionType, AccountInfo, AuthenticationResult } from '@azure/msal-browser';
import "@pnp/sp/webs";
import { environment } from '../../environments/environment';
// import "@pnp/graph/me";

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

// MSAL Request for SharePoint
export const loginRequest = {
  scopes: ['https://digiservhub.sharepoint.com/.default'], // Replace with your SharePoint URL
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

  constructor() {
    this.msalInstance = new PublicClientApplication(msalConfig);
    this.currentAccountSubject = new BehaviorSubject<AccountInfo | null>(null);
    this.currentAccount$ = this.currentAccountSubject.asObservable();
    this.isAuthenticatedSubject = new BehaviorSubject<boolean>(false);
    this.isAuthenticated$ = this.isAuthenticatedSubject.asObservable();

    this.initializeAuth();
    this.initializePnP();
  }

  private initializePnP(): void {
    // Initialize SharePoint client
    this.sp = spfi().using(PnPLogging(LogLevel.Info));
    
    // Initialize Graph client
    this.graph = graphfi().using(PnPLogging(LogLevel.Info));
  }

  private async initializeAuth(): Promise<void> {
    await this.msalInstance.initialize();
    
    // Handle redirect promise
    const response = await this.msalInstance.handleRedirectPromise();
    if (response) {
      this.handleLoginResponse(response);
    }
    
    // Set active account
    const accounts = this.msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      this.currentAccountSubject.next(accounts[0]);
      this.isAuthenticatedSubject.next(true);
      this.setActiveAccount(accounts[0]);
    }
  }

  private setActiveAccount(account: AccountInfo): void {
    this.msalInstance.setActiveAccount(account);
    this.setPnPToken(account);
  }

  private async setPnPToken(account: AccountInfo): Promise<void> {
    try {
      // Configure PnPjs with MSAL token provider
      const msalClient = MSAL({
        configuration: msalConfig,
        authParams: loginRequest,
      });
      
      this.sp = spfi().using(msalClient);
      this.graph = graphfi().using(msalClient);
    } catch (error) {
      console.error('Error setting PnP token:', error);
    }
  }

  private async acquireTokenSilent(scopes: string[]): Promise<AuthenticationResult> {
    const account = this.msalInstance.getActiveAccount();
    if (!account) {
      throw new Error('No active account found');
    }

    const request = {
      scopes: scopes,
      account: account,
    };

    try {
      return await this.msalInstance.acquireTokenSilent(request);
    } catch (error) {
      // If silent acquisition fails, try popup
      return await this.msalInstance.acquireTokenPopup(request);
    }
  }

  private handleLoginResponse(response: AuthenticationResult): void {
    if (response && response.account) {
      this.currentAccountSubject.next(response.account);
      this.isAuthenticatedSubject.next(true);
      this.setActiveAccount(response.account);
    }
  }

  login(): Observable<AuthenticationResult> {
    return from(
      this.msalInstance.loginPopup(loginRequest).then((response) => {
        if (response && response.account) {
          this.currentAccountSubject.next(response.account);
          this.isAuthenticatedSubject.next(true);
          this.setActiveAccount(response.account);
        }
        return response;
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
      this.msalInstance.logoutPopup({
        postLogoutRedirectUri: msalConfig.auth.postLogoutRedirectUri,
        mainWindowRedirectUri: msalConfig.auth.redirectUri,
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

  // async getCurrentUser(): Promise<Observable<any>> {
  //   const currentUser = await this.graph.using;
  //   return from(this.graph.me()).pipe(
  //     map((user) => user),
  //     catchError((error) => {
  //       console.error('Error getting current user:', error);
  //       return throwError(() => error);
  //     })
  //   );
  // }

  getSharePoint(): SPFI {
    return this.sp;
  }

  getGraph() {
    return this.graph;
  }

  getAccessToken(scopes?: string[]): Observable<string> {
    const tokenScopes = scopes || loginRequest.scopes;
    console.log(tokenScopes);
    
    return from(this.acquireTokenSilent(tokenScopes)).pipe(
      map((result) => result.accessToken),
      catchError((error) => {
        console.error('Error getting access token:', error);
        return throwError(() => error);
      })
    );
  }
}