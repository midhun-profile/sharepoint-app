// ==============================================================================
// MICROSOFT ENTRA ID (AZURE AD) MSAL AUTHENTICATION CONFIGURATION
// ==============================================================================
// Strict Security Notice:
// - Uses OAuth 2.0 Authorization Code Flow with PKCE.
// - Tokens are cached securely using MSAL's built-in memory/session mechanics.
// - No client secrets exist in client code.
// ==============================================================================

import { Configuration, PopupRequest, PublicClientApplication } from '@azure/msal-browser';

const clientId = (import.meta as any).env?.VITE_ENTRA_CLIENT_ID || '00000000-0000-0000-0000-000000000000';
const tenantId = (import.meta as any).env?.VITE_ENTRA_TENANT_ID || 'common';
const redirectUri = (import.meta as any).env?.VITE_ENTRA_REDIRECT_URI || (typeof window !== 'undefined' ? window.location.origin : '');

export const msalConfig: Configuration = {
  auth: {
    clientId: clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri: redirectUri,
    postLogoutRedirectUri: redirectUri,
  },
  cache: {
    cacheLocation: 'sessionStorage',
  },
};

export const graphScopes = {
  scopes: [
    'User.Read',
    'Sites.Read.All',
    'Sites.ReadWrite.All',
    'Directory.Read.All',
  ],
};

export const loginRequest: PopupRequest = {
  scopes: graphScopes.scopes,
};

export const msalInstance = new PublicClientApplication(msalConfig);
