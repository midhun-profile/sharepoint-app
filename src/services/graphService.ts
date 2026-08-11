// ==============================================================================
// DIRECT-TO-GRAPH SHAREPOINT SERVICE LAYER
// ==============================================================================
// Interacts directly with Microsoft Graph API using delegated OAuth tokens.
// Features transparent demo engine fallback for instant sandbox preview.
// ==============================================================================

import { graphScopes, msalInstance } from '../config/msalConfig';
import { SharePointColumnDefinition, SharePointListItem } from '../types';
import { MOCK_ITEMS, MOCK_SCHEMAS } from './mockData';

class GraphService {
  private isMockEngine(): boolean {
    const mockEnv = (import.meta as any).env?.VITE_ENABLE_MOCK_ENGINE;
    if (mockEnv === 'false') return false;
    // Default to true if no active MSAL session exists
    const activeAccount = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
    return !activeAccount;
  }

  private async getAccessToken(): Promise<string | null> {
    const account = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
    if (!account) return null;

    try {
      const response = await msalInstance.acquireTokenSilent({
        account,
        scopes: graphScopes.scopes,
      });
      return response.accessToken;
    } catch (error) {
      console.warn('Silent token acquisition failed, attempting interactive login', error);
      return null;
    }
  }

  /**
   * Reads SharePoint List Column Schema Metadata
   */
  async getListColumns(siteId: string, listId: string): Promise<SharePointColumnDefinition[]> {
    if (this.isMockEngine()) {
      return MOCK_SCHEMAS[listId] || [
        { id: '1', name: 'Title', displayName: 'Title', type: 'Text', required: true },
        { id: '2', name: 'Description', displayName: 'Description', type: 'Note' },
        { id: '3', name: 'Status', displayName: 'Status', type: 'Choice', choices: ['Active', 'Pending', 'Closed'] },
        { id: '4', name: 'CreatedDate', displayName: 'Created Date', type: 'DateTime' },
      ];
    }

    const token = await this.getAccessToken();
    if (!token) throw new Error('Authentication required. Please log in with Microsoft Entra ID.');

    const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/columns`;
    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });

    if (!response.ok) {
      throw new Error(`Graph API error (${response.status}): ${await response.text()}`);
    }

    const data = await response.json();
    return data.value.map((col: any) => ({
      id: col.id,
      name: col.name,
      displayName: col.displayName,
      type: this.mapGraphTypeToColumnType(col),
      required: col.required || false,
      readOnly: col.readOnly || false,
      choices: col.choice ? col.choice.choices : undefined,
    }));
  }

  /**
   * Fetches List Items with expanded fields
   */
  async getListItems(siteId: string, listId: string): Promise<SharePointListItem[]> {
    if (this.isMockEngine()) {
      const items = MOCK_ITEMS[listId] || [];
      return [...items];
    }

    const token = await this.getAccessToken();
    if (!token) throw new Error('Authentication required');

    const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?expand=fields`;
    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });

    if (!response.ok) {
      throw new Error(`Graph API error: ${await response.text()}`);
    }

    const data = await response.json();
    return data.value.map((item: any) => ({
      id: item.id,
      created: item.createdDateTime,
      createdBy: item.createdBy?.user ? {
        displayName: item.createdBy.user.displayName,
        email: item.createdBy.user.email || item.createdBy.user.userPrincipalName,
      } : undefined,
      modified: item.lastModifiedDateTime,
      fields: item.fields || {},
    }));
  }

  /**
   * Fetches a single SharePoint List Item by ID
   */
  async getListItem(siteId: string, listId: string, itemId: string): Promise<SharePointListItem> {
    if (this.isMockEngine()) {
      const items = MOCK_ITEMS[listId] || [];
      const found = items.find(i => String(i.id) === String(itemId));
      if (found) return found;
      return {
        id: itemId,
        created: new Date().toISOString(),
        modified: new Date().toISOString(),
        fields: { Title: `Item #${itemId}` },
      };
    }

    const token = await this.getAccessToken();
    if (!token) throw new Error('Authentication required');

    const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${itemId}?expand=fields`;
    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });

    if (!response.ok) {
      throw new Error(`Graph API error (${response.status}): ${await response.text()}`);
    }

    const item = await response.json();
    return {
      id: item.id,
      created: item.createdDateTime,
      createdBy: item.createdBy?.user ? {
        displayName: item.createdBy.user.displayName,
        email: item.createdBy.user.email || item.createdBy.user.userPrincipalName,
      } : undefined,
      modified: item.lastModifiedDateTime,
      fields: item.fields || {},
    };
  }

  /**
   * Creates a new SharePoint List Item
   */
  async createListItem(siteId: string, listId: string, fields: Record<string, any>): Promise<SharePointListItem> {
    if (this.isMockEngine()) {
      const newItem: SharePointListItem = {
        id: `item-${Date.now()}`,
        created: new Date().toISOString(),
        createdBy: { displayName: 'Current Admin', email: 'admin@contoso.com' },
        modified: new Date().toISOString(),
        fields: { ...fields },
      };

      if (!MOCK_ITEMS[listId]) {
        MOCK_ITEMS[listId] = [];
      }
      MOCK_ITEMS[listId].unshift(newItem);
      return newItem;
    }

    const token = await this.getAccessToken();
    if (!token) throw new Error('Authentication required');

    const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`;
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ fields }),
    });

    if (!response.ok) {
      throw new Error(`Failed to create list item: ${await response.text()}`);
    }

    const item = await response.json();
    return {
      id: item.id,
      created: item.createdDateTime,
      fields: item.fields || fields,
    };
  }

  /**
   * Updates an existing SharePoint List Item
   */
  async updateListItem(siteId: string, listId: string, itemId: string, fields: Record<string, any>): Promise<SharePointListItem> {
    if (this.isMockEngine()) {
      const list = MOCK_ITEMS[listId] || [];
      const itemIndex = list.findIndex(i => i.id === itemId);
      if (itemIndex >= 0) {
        list[itemIndex] = {
          ...list[itemIndex],
          modified: new Date().toISOString(),
          fields: { ...list[itemIndex].fields, ...fields },
        };
        return list[itemIndex];
      }
      throw new Error('Item not found');
    }

    const token = await this.getAccessToken();
    if (!token) throw new Error('Authentication required');

    const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${itemId}/fields`;
    const response = await fetch(url, {
      method: 'PATCH',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(fields),
    });

    if (!response.ok) {
      throw new Error(`Failed to update list item: ${await response.text()}`);
    }

    const updatedFields = await response.json();
    return {
      id: itemId,
      modified: new Date().toISOString(),
      fields: updatedFields,
    };
  }

  /**
   * Deletes a SharePoint List Item
   */
  async deleteListItem(siteId: string, listId: string, itemId: string): Promise<boolean> {
    if (this.isMockEngine()) {
      if (MOCK_ITEMS[listId]) {
        MOCK_ITEMS[listId] = MOCK_ITEMS[listId].filter(i => i.id !== itemId);
      }
      return true;
    }

    const token = await this.getAccessToken();
    if (!token) throw new Error('Authentication required');

    const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${itemId}`;
    const response = await fetch(url, {
      method: 'DELETE',
      headers: { Authorization: `Bearer ${token}` },
    });

    if (!response.ok) {
      throw new Error(`Failed to delete item: ${await response.text()}`);
    }

    return true;
  }

  private mapGraphTypeToColumnType(col: any): any {
    if (col.choice) return 'Choice';
    if (col.number) return 'Number';
    if (col.currency) return 'Currency';
    if (col.dateTime) return 'DateTime';
    if (col.boolean) return 'Boolean';
    if (col.personOrGroup) return 'Person';
    if (col.lookup) return 'Lookup';
    if (col.text && col.text.allowMultipleLines) return 'Note';
    return 'Text';
  }
}

export const graphService = new GraphService();
