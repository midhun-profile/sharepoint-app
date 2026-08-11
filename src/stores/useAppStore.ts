// ==============================================================================
// GLOBAL APPLICATION STATE ENGINE (ZUSTAND)
// ==============================================================================

import { create } from 'zustand';
import { AuditLog, MenuConfig, ToastMessage, UserProfile, UserRole, WorkspaceConfig } from '../types';
import { INITIAL_MENUS, INITIAL_WORKSPACES } from '../services/mockData';

interface AppState {
  // Theme & Environment
  isDarkMode: boolean;
  isLiveEntraMode: boolean;
  toggleDarkMode: () => void;
  setLiveEntraMode: (enabled: boolean) => void;

  // Navigation
  activeWorkspaceId: string | null;
  activeMenuId: string | null;
  setActiveWorkspace: (id: string | null) => void;
  setActiveMenu: (id: string | null) => void;

  // Workspaces & Menus Configuration
  workspaces: WorkspaceConfig[];
  menus: MenuConfig[];
  addWorkspace: (workspace: Omit<WorkspaceConfig, 'id' | 'createdAt' | 'updatedAt'>) => WorkspaceConfig;
  updateWorkspace: (id: string, workspace: Partial<WorkspaceConfig>) => void;
  deleteWorkspace: (id: string) => void;

  addMenu: (menu: Omit<MenuConfig, 'id'>) => MenuConfig;
  updateMenu: (id: string, menu: Partial<MenuConfig>) => void;
  deleteMenu: (id: string) => void;

  // User & Security Context
  currentUser: UserProfile;
  setUserRole: (role: UserRole) => void;

  // Audit Logging
  auditLogs: AuditLog[];
  logAction: (action: AuditLog['action'], details: string, workspaceName?: string, menuName?: string) => void;

  // Feedback Notifications
  toasts: ToastMessage[];
  addToast: (type: ToastMessage['type'], title: string, message: string) => void;
  removeToast: (id: string) => void;
}

export const useAppStore = create<AppState>((set, get) => ({
  isDarkMode: false,
  isLiveEntraMode: false,
  toggleDarkMode: () => {
    const next = !get().isDarkMode;
    if (next) {
      document.documentElement.classList.add('dark');
    } else {
      document.documentElement.classList.remove('dark');
    }
    set({ isDarkMode: next });
  },
  setLiveEntraMode: (enabled) => set({ isLiveEntraMode: enabled }),

  activeWorkspaceId: 'ws-projects',
  activeMenuId: 'menu-active-projects',
  setActiveWorkspace: (id) => {
    const state = get();
    // Auto-select first menu in workspace if available
    const firstMenu = state.menus.find((m) => m.workspaceId === id && m.visible);
    set({
      activeWorkspaceId: id,
      activeMenuId: firstMenu ? firstMenu.id : null,
    });
  },
  setActiveMenu: (id) => set({ activeMenuId: id }),

  workspaces: INITIAL_WORKSPACES,
  menus: INITIAL_MENUS,

  addWorkspace: (wsData) => {
    const newWs: WorkspaceConfig = {
      ...wsData,
      id: `ws-${Date.now()}`,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
    set((state) => ({
      workspaces: [...state.workspaces, newWs],
    }));
    get().logAction('CONFIG_CHANGE', `Created new workspace: "${newWs.name}"`);
    get().addToast('success', 'Workspace Created', `Workspace "${newWs.name}" is now available.`);
    return newWs;
  },

  updateWorkspace: (id, wsData) => {
    set((state) => ({
      workspaces: state.workspaces.map((w) =>
        w.id === id ? { ...w, ...wsData, updatedAt: new Date().toISOString() } : w
      ),
    }));
    get().logAction('CONFIG_CHANGE', `Updated workspace settings for ID: ${id}`);
    get().addToast('info', 'Workspace Updated', 'Workspace configuration saved successfully.');
  },

  deleteWorkspace: (id) => {
    const ws = get().workspaces.find((w) => w.id === id);
    set((state) => ({
      workspaces: state.workspaces.filter((w) => w.id !== id),
      menus: state.menus.filter((m) => m.workspaceId !== id),
      activeWorkspaceId: state.activeWorkspaceId === id ? null : state.activeWorkspaceId,
      activeMenuId: state.activeWorkspaceId === id ? null : state.activeMenuId,
    }));
    if (ws) {
      get().logAction('CONFIG_CHANGE', `Deleted workspace "${ws.name}"`);
      get().addToast('warning', 'Workspace Removed', `Workspace "${ws.name}" and its associated menus were deleted.`);
    }
  },

  addMenu: (menuData) => {
    const newMenu: MenuConfig = {
      ...menuData,
      id: `menu-${Date.now()}`,
    };
    set((state) => ({
      menus: [...state.menus, newMenu],
    }));
    get().logAction('CONFIG_CHANGE', `Created menu: "${newMenu.name}"`, undefined, newMenu.name);
    get().addToast('success', 'Menu Configured', `Menu "${newMenu.name}" created and linked to SharePoint List.`);
    return newMenu;
  },

  updateMenu: (id, menuData) => {
    set((state) => ({
      menus: state.menus.map((m) => (m.id === id ? { ...m, ...menuData } : m)),
    }));
    get().addToast('info', 'Menu Updated', 'SharePoint menu configuration saved successfully.');
  },

  deleteMenu: (id) => {
    const menu = get().menus.find((m) => m.id === id);
    set((state) => ({
      menus: state.menus.filter((m) => m.id !== id),
      activeMenuId: state.activeMenuId === id ? null : state.activeMenuId,
    }));
    if (menu) {
      get().addToast('warning', 'Menu Deleted', `Menu "${menu.name}" has been removed.`);
    }
  },

  currentUser: {
    id: 'usr-1001',
    displayName: 'Alex Rivera (Admin)',
    email: 'alex.rivera@contoso.com',
    userPrincipalName: 'alex.rivera@contoso.com',
    role: 'Administrator',
    jobTitle: 'Enterprise IT Platform Architect',
    department: 'Corporate IT Operations',
  },

  setUserRole: (role) => {
    set((state) => ({
      currentUser: { ...state.currentUser, role },
    }));
    get().addToast('info', 'Role Updated', `Switched active user permission role to ${role}.`);
  },

  auditLogs: [
    {
      id: 'log-1',
      timestamp: new Date(Date.now() - 3600000).toISOString(),
      userId: 'usr-1001',
      userName: 'Alex Rivera (Admin)',
      action: 'LOGIN',
      details: 'Logged in via Microsoft Entra ID OAuth 2.0 PKCE',
    },
    {
      id: 'log-2',
      timestamp: new Date(Date.now() - 1800000).toISOString(),
      userId: 'usr-1001',
      userName: 'Alex Rivera (Admin)',
      action: 'CREATE',
      workspaceName: 'Projects',
      menuName: 'Active Projects',
      details: 'Created list item: "SharePoint Modern Intranet & Document Workflows"',
    },
  ],

  logAction: (action, details, workspaceName, menuName) => {
    const user = get().currentUser;
    const newLog: AuditLog = {
      id: `log-${Date.now()}`,
      timestamp: new Date().toISOString(),
      userId: user.id,
      userName: user.displayName,
      action,
      workspaceName,
      menuName,
      details,
    };
    set((state) => ({
      auditLogs: [newLog, ...state.auditLogs.slice(0, 99)], // keep last 100
    }));
  },

  toasts: [],
  addToast: (type, title, message) => {
    const id = `toast-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`;
    const newToast: ToastMessage = { id, type, title, message, timestamp: Date.now() };
    set((state) => ({ toasts: [newToast, ...state.toasts] }));

    setTimeout(() => {
      get().removeToast(id);
    }, 5000);
  },
  removeToast: (id) => {
    set((state) => ({
      toasts: state.toasts.filter((t) => t.id !== id),
    }));
  },
}));
