import { Injectable } from '@angular/core';
import { spfi, SPFI, SPBrowser } from "@pnp/sp";
import { graphfi, GraphFI, GraphBrowser } from '@pnp/graph';
import { getMSAL, MSAL, MSALOptions } from "@pnp/msaljsclient";
import { Configuration, PublicClientApplication, AccountInfo, EventType, AuthenticationResult } from "@azure/msal-browser";
import { environment } from '../../environments/environment';

// Import required SharePoint sub-modules
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/graph/users";
import "@pnp/graph/mail";
import "@pnp/graph/mail/messages";
import { BehaviorSubject, Observable } from 'rxjs';
import { get } from '@pnp/queryable';

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

const msalConfig = {
  auth: {
    authority: environment.msal.authority,
    clientId: environment.msal.clientId,
    redirectUri: window.location.origin,
    allowRedirectInIframe: true,
  },
  cache: {
    cacheLocation: "sessionStorage",
    claimsBasedCachingEnabled: true // in order to avoid network call to refresh a token every time claims are requested
  },
  system: {
    allowNativeBroker: false
  }
}
const spOptions: MSALOptions = {
  configuration: msalConfig,
  authParams: {
    forceRefresh: false,
    scopes: [`${environment.sharePoint.baseUrl}/.default`],
  }
};
const graphOptions: MSALOptions = {
  configuration: msalConfig,
  authParams: {
    forceRefresh: false,
    scopes: ["https://graph.microsoft.com/.default"],
  }
};

@Injectable({
  providedIn: 'root'
})
export class PnPjs {
  private _sp!: SPFI;
  private _graph!: GraphFI;
  // private msalInstance: PublicClientApplication = new PublicClientApplication(spOptions.configuration);
  private msalInstance!: PublicClientApplication ;

  // Reactive State
  private currentAccountSubject = new BehaviorSubject<AccountInfo | null>(null);
  public currentAccount$ = this.currentAccountSubject.asObservable();

  private isAuthenticatedSubject = new BehaviorSubject<boolean>(false);
  public isAuthenticated$ = this.isAuthenticatedSubject.asObservable();

  constructor() {
    this.initPnPjs();
    // this.msalInstance = getMSAL();
  }

  private setAuthenticatedUser(account: AccountInfo) {
    this.currentAccountSubject.next(account);
    this.isAuthenticatedSubject.next(true);
  }

  async initPnPjs() {
    console.log("PnPJs Initiated");
    this._sp = spfi(environment.sharePoint.siteUrl).using(SPBrowser(),MSAL(spOptions));
    this._graph = graphfi().using(GraphBrowser(),MSAL(graphOptions));
    await this.getMyProfile();
  }

  public get sp(): SPFI { return this._sp; }
  public get graph(): GraphFI { return this._graph; }
  
  async getMyProfile() {
    const me = await this.graph.me();
    this.msalInstance = getMSAL() as unknown as PublicClientApplication;
    const allAccounts = await this.msalInstance.getAllAccounts();
    const currentAccount = await this.msalInstance.getActiveAccount();
    console.log(currentAccount);
    if (currentAccount) {
      this.setAuthenticatedUser(currentAccount as AccountInfo);
    }
    return me;
  }
  async logout() {
    // const msalInstance = getMSAL();
    // const currentAccount = msalInstance.getActiveAccount();
    this.msalInstance.logoutPopup({ account: this.msalInstance.getActiveAccount() });
    this.currentAccountSubject.next(null);
    this.isAuthenticatedSubject.next(false);
  }

  allListItems(listname: string) {
    return this._sp.web.lists.getByTitle(listname).items();
  }
  updateListItem(listname: string, id: number, data: any) {
    return this._sp.web.lists.getByTitle(listname).items.getById(id).update(data);
  }
  addListItem(listname: string, data: any) {
    return this._sp.web.lists.getByTitle(listname).items.add(data);
  }
}