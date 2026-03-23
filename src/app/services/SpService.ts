import { Injectable } from '@angular/core';
import { spfi, SPFI, SPBrowser } from "@pnp/sp";
import { graphfi, GraphFI, GraphBrowser } from '@pnp/graph';
import { MSAL, MSALOptions } from "@pnp/msaljsclient";
import { Configuration, PublicClientApplication, AccountInfo } from "@azure/msal-browser";
import { environment } from '../../environments/environment';

// Import required SharePoint sub-modules
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/graph/users";
import "@pnp/graph/mail";
import "@pnp/graph/mail/messages";

@Injectable({
  providedIn: 'root'
})
export class SpService {
  private _sp: SPFI | null = null;
  private _graph: GraphFI | null = null;
  private msalInstance: PublicClientApplication;
  // MSAL Configuration
  private msalConfig  = {
    auth: {
      clientId: environment.msal.clientId,
      authority: environment.msal.authority,
      redirectUri: window.location.origin,
    //   allowRedirectInIframe: true,
    },
    cache: {
      cacheLocation: "sessionStorage",
      claimsBasedCachingEnabled: true // in order to avoid network call to refresh a token every time claims are requested
    },
    system: {
      // Optional: If you still see the error, ensuring the interaction is 
      // not blocked by aggressive browser policies
      allowNativeBroker: false 
    }    
  };

  constructor() {
    this.msalInstance = new PublicClientApplication(this.msalConfig as Configuration);
    this.initPnPjs();
    this.initGraph();
  }

  private initPnPjs() {
    // In PnPjs v4, MSAL takes a single options object
    this._sp = spfi("https://digiservhub.sharepoint.com/sites/Pulse")
      .using(
        SPBrowser(), 
        MSAL({
          configuration: this.msalConfig,
          authParams: {
            forceRefresh: false,
            scopes: ["https://digiservhub.sharepoint.com/.default"]
          }
        })
      );
      
  }
  private initGraph() {
    // Correct implementation for GraphFI
    this._graph = graphfi()
      .using(
        GraphBrowser(), 
        MSAL({
          configuration: this.msalConfig,
          authParams: {
            // Use Graph scopes, NOT SharePoint scopes
            forceRefresh: false,
            scopes: ["https://graph.microsoft.com/.default"]
          }
        })
      );
  }

  public get sp(): SPFI {
    if (!this._sp) {
      throw new Error("PnPjs not initialized.");
    }
    return this._sp;
  }
  public get graph(): GraphFI {
    if (!this._graph) {
      throw new Error("Graph not initialized.");
    }
    return this._graph;
  }

  async getCurrentUser1() {
    return await this.sp.web.currentUser();
  }
  async getCurrentUser() {
    return await this.graph.me();
  }

  async getListItems(listName: string) {
    return await this.sp.web.lists.getByTitle(listName).items();
  }
    // This method is used by the Auth Guard
  public async isAuthenticated(): Promise<boolean> {
    try {
      const accounts = this.msalInstance.getAllAccounts();
      return accounts.length > 0;
    } catch {
      return false;
    }
  }


  async sendEmail(to: string, subject: string, body: string) {
    try {
      // PnPjs sendMail expects the Message properties directly.
      // contentType must be lowercase "html" or "text" per TypeScript definitions.
      await this.graph.me.sendMail({
        subject: subject,
        body: {
          contentType: "html",
          content: body,
        },
        toRecipients: [
          {
            emailAddress: {
              address: to,
            },
          },
        ]
      }, true); // Second parameter is 'saveToSentItems'
      
      return { success: true };
    } catch (error) {
      console.error("Error sending email via Graph:", error);
      throw error;
    }
  }
}