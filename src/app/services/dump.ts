// import { Injectable } from '@angular/core';
// import { spfi, SPFI, SPBrowser } from "@pnp/sp";
// import { graphfi, GraphFI, GraphBrowser } from '@pnp/graph';
// import { getMSAL, MSAL, MSALOptions } from "@pnp/msaljsclient";
// import { Configuration, PublicClientApplication, AccountInfo, EventType, AuthenticationResult } from "@azure/msal-browser";
// import { environment } from '../../environments/environment';

// // Import required SharePoint sub-modules
// import "@pnp/sp/webs";
// import "@pnp/sp/lists";
// import "@pnp/sp/items";
// import "@pnp/sp/site-users/web";
// import "@pnp/graph/users";
// import "@pnp/graph/mail";
// import "@pnp/graph/mail/messages";
// import { BehaviorSubject, Observable } from 'rxjs';

// const options: MSALOptions = {
//   configuration: {
//     auth: {
//       authority: environment.msal.authority,
//       clientId: environment.msal.clientId,
//     },
//     cache: {
//       claimsBasedCachingEnabled: true // in order to avoid network call to refresh a token every time claims are requested
//     }
//   },
//   authParams: {
//     forceRefresh: false,
//     scopes: ["https://{tenant}.sharepoint.com/.default"],
//   }
// };

// @Injectable({
//   providedIn: 'root'
// })
// export class PnPjs {
//   private _sp!: SPFI;
//   private _graph!: GraphFI;
//   private msalInstance: PublicClientApplication;
  
//   // Reactive State
//   private currentAccountSubject = new BehaviorSubject<AccountInfo | null>(null);
//   public currentAccount$ = this.currentAccountSubject.asObservable();
  
//   private isAuthenticatedSubject = new BehaviorSubject<boolean>(false);
//   public isAuthenticated$ = this.isAuthenticatedSubject.asObservable();
    
//   private readonly msalConfig = {
//     auth: {
//       clientId: environment.msal.clientId,
//       authority: environment.msal.authority,
//       redirectUri: window.location.origin,
//       allowRedirectInIframe: true,
//     },
//     cache: {
//       cacheLocation: "sessionStorage",
//       claimsBasedCachingEnabled: true
//     },
//     system: {
//       allowNativeBroker: false 
//     }    
//   };



//   constructor() {
//     this.msalInstance = new PublicClientApplication(this.msalConfig as Configuration);
//     this.setupMsal();
//     // this.initializeMsal();
//     // this.initPnPjs();
//   }
//   private async setupMsal() {
//     await this.msalInstance.initialize();

//     // Check if a user is already signed in from a previous session
//     const accounts = await this.msalInstancePublic.getAllAccounts();
//     if (accounts.length > 0) {
//       this.setAuthenticatedUser(accounts[0]);
//     }

//     // Initialize PnPjs engines
//     this.initPnPjs();
//   }
//   private setAuthenticatedUser(account: AccountInfo) {
//     this.currentAccountSubject.next(account);
//     this.isAuthenticatedSubject.next(true);
//   }

//   private async initPnPjs() {
//     const msalOptions = {
//       configuration: this.msalConfig,
//       authParams: { forceRefresh: false }
//     };

//     this._sp = spfi(environment.sharePoint.siteUrl).using(
//       SPBrowser(),
//       MSAL({ ...msalOptions, authParams: { scopes: [`${environment.sharePoint.baseUrl}/.default`] } })
//     );

//     this._graph = graphfi().using(
//       GraphBrowser(),
//       MSAL({ ...msalOptions, authParams: { scopes: ["https://graph.microsoft.com/.default"] } })
//     );
//     const accounts = await this.msalInstancePublic.getAllAccounts();
//     console.log(accounts);
    
//   }
//   public async login(): Promise<void> {
//     try {
//       // Ensure instance is ready
//       if (!this.msalInstance) {
//         await this.setupMsal();
//       }

//       const loginRequest = {
//         scopes: ["https://graph.microsoft.com/.default"],
//         prompt: "select_account"
//       };

//       const result: AuthenticationResult = await this.msalInstance.loginPopup(loginRequest);

//       if (result.account) {
//         this.setAuthenticatedUser(result.account);
//         // Explicitly notify MSAL that the interaction is complete 
//         // if the popup doesn't close automatically
//         this.msalInstance.setActiveAccount(result.account);
//       }
//     } catch (error) {
//       console.error("Login failed:", error);
//       throw error;
//     }
//   }
  
//   public async checkAuthStatus(): Promise<boolean> {
//     // If we already have a value in the subject, return it
//     if (this.isAuthenticatedSubject.value) return true;
    
//     // Otherwise, wait briefly for setupMsal to complete
//     const accounts = this.msalInstance.getAllAccounts();
//     return accounts.length > 0;
//   }


//   private async initializeMsal() {
//     // Required for Redirect flows
//     await this.msalInstance.initialize();
//     console.log("MSAL Initiated");
    
//     // Listen for login/logout events to update the app state reactively
//     this.msalInstance.addEventCallback((message) => {
//       if (message.eventType === EventType.LOGIN_SUCCESS && message.payload) {
//         const payload = message.payload as any;
//         this.currentAccountSubject.next(payload.account);
//         this.isAuthenticatedSubject.next(true);
//       }
//     });

//     const accounts = this.msalInstance.getAllAccounts();
//     if (accounts.length > 0) {
//       this.currentAccountSubject.next(accounts[0]);
//       this.isAuthenticatedSubject.next(true);
//     }
//   }
//   private initPnPjs2() {
//     const commonMsalParams = {
//       configuration: this.msalConfig,
//       authParams: { forceRefresh: false }
//     };
//     // Initialize SP
//     this._sp = spfi(environment.sharePoint.siteUrl)
//       .using(
//         SPBrowser(),
//         MSAL({
//           ...commonMsalParams,
//           authParams: { 
//             ...commonMsalParams.authParams, 
//             scopes: [`${environment.sharePoint.baseUrl}/.default`] 
//           }
//         })
//       );

//     // Initialize Graph
//     this._graph = graphfi()
//       .using(
//         GraphBrowser(),
//         MSAL({
//           ...commonMsalParams,
//           authParams: { 
//             ...commonMsalParams.authParams, 
//             scopes: ["https://graph.microsoft.com/.default"] 
//           }
//         })
//       );
//   }
//   // Exposed Getters
//   public get sp(): SPFI { return this._sp; }
//   public get graph(): GraphFI { return this._graph; }

//   async getMyProfile() {
//     const me = await this.graph.me();
//     console.log(this.msalInstancePublic);
    
//     return me;
//   }
//   async logout() {
//     const msalInstance = getMSAL();
//     msalInstance.logoutRedirect();
//     this.currentAccountSubject.next(null);
//     this.isAuthenticatedSubject.next(false);
//   }

// //   get msalInstancePublic(): PublicClientApplication {
// //     return getMSAL();
// //   }






//   private initPnPjs1() {
//     this._sp = spfi("https://digiservhub.sharepoint.com/sites/Pulse")
//       .using(
//         SPBrowser(),
//         MSAL({
//           configuration: this.msalConfig,
//           authParams: {
//             forceRefresh: false,
//             scopes: ["https://digiservhub.sharepoint.com/.default"]
//           }
//         })
//       );
//   }
//   private initGraph1() {
//     this._graph = graphfi()
//       .using(
//         GraphBrowser(),
//         MSAL({
//           configuration: this.msalConfig,
//           authParams: {
//             forceRefresh: false,
//             scopes: ["https://graph.microsoft.com/.default"]
//           }
//         })
//       );
//   }

//   // public get sp(): SPFI {
//   //   if (!this._sp) {
//   //     throw new Error("PnPjs not initialized.");
//   //   }
//   //   return this._sp;
//   // }
//   // public get graph(): GraphFI {
//   //   if (!this._graph) {
//   //     throw new Error("Graph not initialized.");
//   //   }
//   //   return this._graph;
//   // }

//   spCurrentUser1() {
//     return this.sp.web.currentUser();
//   }

//   // public async isAuthenticated(): Promise<boolean> {
//   //   try {
//   //     const msalInstance = getMSAL();
//   //     const accounts = msalInstance.getAllAccounts();
//   //     return accounts.length > 0;
//   //   } catch {
//   //     return false;
//   //   }
//   // }
// }