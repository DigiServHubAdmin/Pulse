export const environment = {
    production: true,
    clientId: 'abb1441a-0c00-425a-b6ea-d087d543cb62',
    tenantId: '89ac7044-3e92-403a-b2ab-968be1d01657',
    msal: {
        clientId: 'abb1441a-0c00-425a-b6ea-d087d543cb62',
        authority: 'https://login.microsoftonline.com/89ac7044-3e92-403a-b2ab-968be1d01657',
        redirectUri: window.location.origin,
        postLogoutRedirectUri: window.location.origin
    },
    sharePoint: {
        baseUrl: 'https://digiservhub.sharepoint.com',
        siteUrl: 'https://digiservhub.sharepoint.com/sites/pulse'
    },
    graphScopes: ['User.Read', 'Mail.Read', 'Sites.Read.All']
};
