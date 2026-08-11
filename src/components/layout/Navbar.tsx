import React, { useState } from 'react';
import { useAppStore } from '../../stores/useAppStore';
import { UserRole } from '../../types';
import {
  Layers,
  Search,
  Sun,
  Moon,
  Shield,
  Activity,
  ChevronDown,
  Database,
  Building2,
  Lock,
  X,
  History,
} from 'lucide-react';

interface NavbarProps {
  onOpenAdmin: () => void;
  onOpenAudit: () => void;
  globalSearch: string;
  onGlobalSearchChange: (value: string) => void;
}

export const Navbar: React.FC<NavbarProps> = ({
  onOpenAdmin,
  onOpenAudit,
  globalSearch,
  onGlobalSearchChange,
}) => {
  const {
    isDarkMode,
    toggleDarkMode,
    currentUser,
    setUserRole,
    isLiveEntraMode,
    setLiveEntraMode,
    activeWorkspaceId,
    workspaces,
    setActiveWorkspace,
  } = useAppStore();

  const [isRoleDropdownOpen, setIsRoleDropdownOpen] = useState(false);

  const activeWs = workspaces.find((w) => w.id === activeWorkspaceId);

  return (
    <header className="sticky top-0 z-30 h-16 bg-white dark:bg-slate-900 border-b border-slate-200 dark:border-slate-800 px-4 md:px-6 flex items-center justify-between gap-4 transition-colors">
      {/* Left section: App Brand Logo & Workspace Switcher */}
      <div className="flex items-center gap-3 md:gap-5">
        <button
          onClick={() => setActiveWorkspace(null)}
          className="flex items-center gap-2.5 text-slate-900 dark:text-slate-100 font-bold text-base hover:opacity-90 transition-opacity"
        >
          <div className="w-9 h-9 rounded-lg bg-gradient-to-br from-brand-600 to-brand-800 text-white flex items-center justify-center shadow-md">
            <Layers className="w-5 h-5" />
          </div>
          <div className="hidden sm:flex flex-col text-left">
            <span className="leading-tight text-sm font-semibold tracking-tight">SharePoint Platform</span>
            <span className="text-[10px] text-slate-500 dark:text-slate-400 font-medium tracking-wide uppercase">Contoso Enterprise</span>
          </div>
        </button>

        {activeWs && (
          <div className="hidden md:flex items-center gap-2 pl-4 border-l border-slate-200 dark:border-slate-800">
            <span className="text-xs font-medium text-slate-400">Workspace:</span>
            <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md text-xs font-semibold bg-brand-50 text-brand-700 dark:bg-brand-950/60 dark:text-brand-300 border border-brand-200/60 dark:border-brand-800/60">
              <Building2 className="w-3.5 h-3.5" />
              {activeWs.name}
            </span>
          </div>
        )}
      </div>

      {/* Middle section: Global Search */}
      <div className="flex-1 max-w-md hidden lg:block">
        <div className="relative">
          <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
          <input
            type="text"
            value={globalSearch}
            onChange={(e) => onGlobalSearchChange(e.target.value)}
            placeholder="Global search lists, columns, items..."
            className="w-full pl-9 pr-8 py-1.5 text-xs rounded-lg border border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-950 text-slate-900 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-brand-500/50"
          />
          {globalSearch && (
            <button
              onClick={() => onGlobalSearchChange('')}
              className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
            >
              <X className="w-3.5 h-3.5" />
            </button>
          )}
        </div>
      </div>

      {/* Right section: System controls, Engine toggle, Role selector, Admin & User Profile */}
      <div className="flex items-center gap-2 sm:gap-3">
        {/* Live Entra vs Demo Engine Switcher */}
        <button
          onClick={() => setLiveEntraMode(!isLiveEntraMode)}
          className={`px-2.5 py-1.5 rounded-md text-xs font-medium border flex items-center gap-1.5 transition-colors ${
            isLiveEntraMode
              ? 'bg-emerald-50 text-emerald-700 border-emerald-300 dark:bg-emerald-950/50 dark:text-emerald-300 dark:border-emerald-800'
              : 'bg-amber-50 text-amber-700 border-amber-300 dark:bg-amber-950/50 dark:text-amber-300 dark:border-amber-800'
          }`}
          title="Toggle between Live Entra ID Graph connection and Demo engine"
        >
          <Database className="w-3.5 h-3.5 shrink-0" />
          <span className="hidden sm:inline">
            {isLiveEntraMode ? 'Live Graph API' : 'Demo Engine'}
          </span>
        </button>

        {/* Role Switcher */}
        <div className="relative">
          <button
            onClick={() => setIsRoleDropdownOpen(!isRoleDropdownOpen)}
            className="px-2.5 py-1.5 rounded-md text-xs font-medium border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 text-slate-700 dark:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-800 flex items-center gap-1.5 transition-colors"
          >
            <Shield className="w-3.5 h-3.5 text-brand-600 dark:text-brand-400" />
            <span className="hidden sm:inline">{currentUser.role}</span>
            <ChevronDown className="w-3 h-3 opacity-60" />
          </button>

          {isRoleDropdownOpen && (
            <div className="absolute right-0 mt-1.5 w-44 rounded-lg bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 shadow-xl py-1 z-50 animate-fade-in text-xs">
              <div className="px-3 py-1.5 text-[10px] font-semibold text-slate-400 uppercase tracking-wider border-b border-slate-100 dark:border-slate-800">
                Switch Security Role
              </div>
              {(['Administrator', 'Manager', 'Employee'] as UserRole[]).map((r) => (
                <button
                  key={r}
                  onClick={() => {
                    setUserRole(r);
                    setIsRoleDropdownOpen(false);
                  }}
                  className={`w-full text-left px-3 py-2 flex items-center justify-between hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors ${
                    currentUser.role === r ? 'font-bold text-brand-600 dark:text-brand-400 bg-brand-50/50 dark:bg-brand-950/30' : 'text-slate-700 dark:text-slate-300'
                  }`}
                >
                  {r}
                  {currentUser.role === r && <Lock className="w-3 h-3 text-brand-600" />}
                </button>
              ))}
            </div>
          )}
        </div>

        {/* Audit Log Trigger */}
        <button
          onClick={onOpenAudit}
          className="p-2 rounded-md text-slate-500 hover:text-slate-700 dark:text-slate-400 dark:hover:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors"
          title="Audit Logs"
        >
          <History className="w-4 h-4" />
        </button>

        {/* Dark Mode Toggle */}
        <button
          onClick={toggleDarkMode}
          className="p-2 rounded-md text-slate-500 hover:text-slate-700 dark:text-slate-400 dark:hover:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors"
          title={isDarkMode ? 'Switch to Light Mode' : 'Switch to Dark Mode'}
        >
          {isDarkMode ? <Sun className="w-4 h-4 text-amber-400" /> : <Moon className="w-4 h-4" />}
        </button>

        {/* Admin Configuration Button */}
        {currentUser.role === 'Administrator' && (
          <button
            onClick={onOpenAdmin}
            className="px-3 py-1.5 rounded-md text-xs font-semibold bg-brand-600 hover:bg-brand-700 text-white shadow-sm transition-colors flex items-center gap-1.5"
          >
            <Activity className="w-3.5 h-3.5" />
            <span className="hidden md:inline">Admin Panel</span>
          </button>
        )}

        {/* User Avatar Badge */}
        <div className="flex items-center gap-2 pl-2 border-l border-slate-200 dark:border-slate-800">
          <div className="w-7 h-7 rounded-full bg-slate-800 dark:bg-slate-200 text-white dark:text-slate-900 flex items-center justify-center font-bold text-xs shadow-sm">
            {currentUser.displayName.charAt(0)}
          </div>
        </div>
      </div>
    </header>
  );
};
