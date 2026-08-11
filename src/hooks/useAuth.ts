import { useCallback } from 'react';
import { msalInstance, graphScopes } from '../config/msalConfig';

/**
 * Enterprise Authentication Hook
 *
 * Provides silent token acquisition for Microsoft Graph API requests with
 * automated fallback for sandbox demo engine execution.
 */
export function useAuth() {
  const getAccessToken = useCallback(async (): Promise<string | null> => {
    const mockEnv = (import.meta as any).env?.VITE_ENABLE_MOCK_ENGINE;
    const account = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];

    if (!account) {
      // Return mock token if running in demo/mock mode without active MSAL session
      if (mockEnv !== 'false') {
        return 'mock-demo-bearer-token';
      }
      return null;
    }

    try {
      const response = await msalInstance.acquireTokenSilent({
        account,
        scopes: graphScopes.scopes,
      });
      return response.accessToken;
    } catch (error) {
      console.warn('Silent token acquisition failed in useAuth:', error);
      if (mockEnv !== 'false') {
        return 'mock-demo-bearer-token';
      }
      return null;
    }
  }, []);

  const activeAccount = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];

  return {
    getAccessToken,
    account: activeAccount,
    isAuthenticated: Boolean(activeAccount),
  };
}
