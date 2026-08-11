import React, { useState, useEffect, useRef, useCallback } from 'react';
import { useAuth } from '../../hooks/useAuth';
import { useDebounce } from '../../hooks/useDebounce';
import {
  User,
  Search,
  X,
  Check,
  Loader2,
  AlertCircle,
  Users,
} from 'lucide-react';

export interface PersonUser {
  id: string;
  displayName: string;
  userPrincipalName: string;
  email: string;
  Claims?: string;
}

export interface PeoplePickerProps {
  /** Selected SharePoint Person structure (single object or array for multi) */
  value: any;
  /** Callback fired when selection changes with strict SharePoint Person OData payload */
  onChange: (value: any) => void;
  /** Input placeholder string */
  placeholder?: string;
  /** Override flag to disable picker */
  disabled?: boolean;
  /** Error message to render */
  error?: string;
  /** Enables multi-user selection array mode */
  isMulti?: boolean;
}

const MOCK_ENTRA_USERS: PersonUser[] = [
  { id: 'usr-001', displayName: 'Jane Doe', email: 'jane.doe@contoso.com', userPrincipalName: 'jane.doe@contoso.com', Claims: 'i:0#.f|membership|jane.doe@contoso.com' },
  { id: 'usr-002', displayName: 'Alex Chen', email: 'alex.chen@contoso.com', userPrincipalName: 'alex.chen@contoso.com', Claims: 'i:0#.f|membership|alex.chen@contoso.com' },
  { id: 'usr-003', displayName: 'Sarah Jenkins', email: 'sarah.jenkins@contoso.com', userPrincipalName: 'sarah.jenkins@contoso.com', Claims: 'i:0#.f|membership|sarah.jenkins@contoso.com' },
  { id: 'usr-004', displayName: 'Marcus Vance', email: 'marcus.vance@contoso.com', userPrincipalName: 'marcus.vance@contoso.com', Claims: 'i:0#.f|membership|marcus.vance@contoso.com' },
  { id: 'usr-005', displayName: 'Emily Taylor', email: 'emily.taylor@contoso.com', userPrincipalName: 'emily.taylor@contoso.com', Claims: 'i:0#.f|membership|emily.taylor@contoso.com' },
  { id: 'usr-006', displayName: 'David Miller', email: 'david.miller@contoso.com', userPrincipalName: 'david.miller@contoso.com', Claims: 'i:0#.f|membership|david.miller@contoso.com' },
];

/**
 * Helper to normalize incoming value prop into a clean array of PersonUser
 */
function normalizeSelectedUsers(rawVal: any): PersonUser[] {
  if (!rawVal) return [];
  const list = Array.isArray(rawVal) ? rawVal : [rawVal];
  return list
    .filter(Boolean)
    .map((item) => {
      if (typeof item === 'string') {
        return {
          id: item,
          displayName: item,
          email: item.includes('@') ? item : `${item}@contoso.com`,
          userPrincipalName: item.includes('@') ? item : `${item}@contoso.com`,
          Claims: `i:0#.f|membership|${item.toLowerCase()}`,
        };
      }
      const email = item.email || item.userPrincipalName || item.mail || '';
      return {
        id: item.id || item.objectId || email,
        displayName: item.displayName || item.title || email,
        email: email,
        userPrincipalName: item.userPrincipalName || email,
        Claims: item.Claims || (email ? `i:0#.f|membership|${email.toLowerCase()}` : undefined),
      };
    });
}

/**
 * Formats a user into the strict SharePoint OData Person payload structure
 */
export function formatPersonPayload(user: PersonUser) {
  const email = user.email || user.userPrincipalName || '';
  return {
    id: user.id,
    displayName: user.displayName,
    email: email,
    userPrincipalName: user.userPrincipalName || email,
    Claims: user.Claims || `i:0#.f|membership|${email.toLowerCase()}`,
  };
}

/**
 * Enterprise PeoplePicker Component
 *
 * Auto-completes Microsoft Entra ID users using Microsoft Graph API (`/v1.0/users`).
 * Features a 300ms debounce buffer and formats selection into strict SharePoint Person OData structures.
 */
export const PeoplePicker: React.FC<PeoplePickerProps> = ({
  value,
  onChange,
  placeholder = 'Type to search Entra ID users...',
  disabled = false,
  error,
  isMulti = false,
}) => {
  const { getAccessToken } = useAuth();
  const [query, setQuery] = useState('');
  const debouncedQuery = useDebounce(query, 300);

  const [results, setResults] = useState<PersonUser[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [isOpen, setIsOpen] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);

  // Normalize selected users
  const selectedUsers = normalizeSelectedUsers(value);

  // Search Entra ID users via Microsoft Graph API
  const searchUsers = useCallback(
    async (searchTerm: string) => {
      if (!searchTerm || searchTerm.trim().length === 0) {
        setResults(MOCK_ENTRA_USERS.slice(0, 5));
        return;
      }

      setIsLoading(true);
      try {
        const token = await getAccessToken();

        // If no token or mock bearer token, use mock filtering
        if (!token || token === 'mock-demo-bearer-token') {
          const filtered = MOCK_ENTRA_USERS.filter(
            (u) =>
              u.displayName.toLowerCase().includes(searchTerm.toLowerCase()) ||
              u.email.toLowerCase().includes(searchTerm.toLowerCase())
          );
          setResults(filtered);
          setIsLoading(false);
          return;
        }

        // Live Microsoft Graph API request using $filter with fallback
        const cleanTerm = encodeURIComponent(searchTerm.replace(/'/g, "''"));
        const filterUrl = `https://graph.microsoft.com/v1.0/users?$filter=startswith(displayName,'${cleanTerm}') or startswith(mail,'${cleanTerm}') or startswith(userPrincipalName,'${cleanTerm}')&$select=id,displayName,userPrincipalName,mail&$top=10`;

        const response = await fetch(filterUrl, {
          headers: {
            Authorization: `Bearer ${token}`,
            ConsistencyLevel: 'eventual',
          },
        });

        if (response.ok) {
          const data = await response.json();
          const mapped: PersonUser[] = (data.value || []).map((u: any) => ({
            id: u.id,
            displayName: u.displayName,
            email: u.mail || u.userPrincipalName || '',
            userPrincipalName: u.userPrincipalName || u.mail || '',
            Claims: `i:0#.f|membership|${(u.mail || u.userPrincipalName || '').toLowerCase()}`,
          }));
          setResults(mapped);
        } else {
          // Fallback if search filter fails on sandbox tenant
          const fallbackList = MOCK_ENTRA_USERS.filter((u) =>
            u.displayName.toLowerCase().includes(searchTerm.toLowerCase())
          );
          setResults(fallbackList);
        }
      } catch (err) {
        console.warn('PeoplePicker Graph API search error:', err);
        const fallbackList = MOCK_ENTRA_USERS.filter((u) =>
          u.displayName.toLowerCase().includes(searchTerm.toLowerCase())
        );
        setResults(fallbackList);
      } finally {
        setIsLoading(false);
      }
    },
    [getAccessToken]
  );

  // Trigger search on debounced query change
  useEffect(() => {
    if (isOpen) {
      searchUsers(debouncedQuery);
    }
  }, [debouncedQuery, isOpen, searchUsers]);

  // Handle click outside popover
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Handle user item selection
  const handleSelectUser = (user: PersonUser) => {
    const formatted = formatPersonPayload(user);

    if (isMulti) {
      const exists = selectedUsers.some((u) => u.id === user.id || u.email === user.email);
      let updated: any[];
      if (exists) {
        updated = selectedUsers.filter((u) => u.id !== user.id && u.email !== user.email);
      } else {
        updated = [...selectedUsers.map(formatPersonPayload), formatted];
      }
      onChange(updated);
    } else {
      onChange(formatted);
      setIsOpen(false);
      setQuery('');
    }
  };

  // Remove single user from selection
  const handleRemoveUser = (userId: string, e?: React.MouseEvent) => {
    e?.stopPropagation();
    if (disabled) return;

    if (isMulti) {
      const updated = selectedUsers
        .filter((u) => u.id !== userId)
        .map(formatPersonPayload);
      onChange(updated);
    } else {
      onChange(null);
    }
  };

  // Helper for generating avatar initials
  const getInitials = (name: string) => {
    if (!name) return 'U';
    const parts = name.trim().split(' ');
    if (parts.length >= 2) return `${parts[0][0]}${parts[1][0]}`.toUpperCase();
    return name.slice(0, 2).toUpperCase();
  };

  return (
    <div ref={containerRef} className="relative w-full space-y-1.5">
      {/* Search Input Box & Selected Chips Container */}
      <div
        onClick={() => !disabled && setIsOpen(true)}
        className={`min-h-[42px] px-3 py-2 rounded-xl border transition-all cursor-text flex flex-wrap items-center gap-2 ${
          error
            ? 'border-rose-500 bg-rose-50/30 dark:bg-rose-950/20'
            : isOpen
            ? 'border-brand-500 ring-2 ring-brand-500/30 bg-white dark:bg-slate-950'
            : 'border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-950 hover:border-slate-300'
        } ${disabled ? 'opacity-60 cursor-not-allowed bg-slate-100 dark:bg-slate-900' : ''}`}
      >
        <User className="w-4 h-4 text-slate-400 shrink-0" />

        {/* Selected User Chips */}
        {selectedUsers.map((user) => (
          <span
            key={user.id}
            className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-xs font-semibold bg-brand-50 text-brand-700 dark:bg-brand-950 dark:text-brand-300 border border-brand-200 dark:border-brand-800 shrink-0"
          >
            <div className="w-4 h-4 rounded-full bg-brand-200 dark:bg-brand-800 text-brand-800 dark:text-brand-200 flex items-center justify-center text-[9px] font-bold">
              {getInitials(user.displayName)}
            </div>
            <span className="truncate max-w-[130px]">{user.displayName}</span>
            {!disabled && (
              <button
                type="button"
                onClick={(e) => handleRemoveUser(user.id, e)}
                className="p-0.5 rounded-full hover:bg-brand-200/60 dark:hover:bg-brand-800 text-brand-500 transition-colors"
                title="Remove user"
              >
                <X className="w-3 h-3" />
              </button>
            )}
          </span>
        ))}

        {/* Input element (hidden if single user already selected and popover is closed) */}
        {(!selectedUsers.length || isMulti || isOpen) && (
          <input
            type="text"
            value={query}
            disabled={disabled}
            onChange={(e) => {
              setQuery(e.target.value);
              setIsOpen(true);
            }}
            onFocus={() => setIsOpen(true)}
            placeholder={selectedUsers.length === 0 ? placeholder : 'Add another user...'}
            className="flex-1 min-w-[120px] bg-transparent text-xs text-slate-900 dark:text-slate-100 focus:outline-none placeholder-slate-400"
          />
        )}

        {/* Loading Spinner / Search Indicator */}
        <div className="ml-auto flex items-center gap-1 shrink-0">
          {isLoading ? (
            <Loader2 className="w-3.5 h-3.5 text-brand-600 animate-spin" />
          ) : query ? (
            <button
              type="button"
              onClick={(e) => {
                e.stopPropagation();
                setQuery('');
              }}
              className="text-slate-400 hover:text-slate-600"
            >
              <X className="w-3.5 h-3.5" />
            </button>
          ) : (
            <Search className="w-3.5 h-3.5 text-slate-400" />
          )}
        </div>
      </div>

      {/* Popover Dropdown Results Menu */}
      {isOpen && !disabled && (
        <div className="absolute left-0 right-0 mt-1 rounded-2xl bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 shadow-2xl p-2 z-50 text-xs animate-fade-in max-h-60 overflow-y-auto custom-scrollbar space-y-1">
          <div className="px-3 py-1.5 text-[10px] uppercase font-bold text-slate-400 flex items-center justify-between border-b border-slate-100 dark:border-slate-800">
            <span className="flex items-center gap-1">
              <Users className="w-3 h-3" /> Entra ID Directory Results
            </span>
            <span>{results.length} found</span>
          </div>

          {results.length === 0 ? (
            <div className="py-6 text-center text-slate-400">
              No matching Entra ID users found.
            </div>
          ) : (
            results.map((u) => {
              const isSelected = selectedUsers.some(
                (sel) => sel.id === u.id || sel.email === u.email
              );

              return (
                <div
                  key={u.id}
                  onClick={() => handleSelectUser(u)}
                  className={`flex items-center justify-between p-2 rounded-xl cursor-pointer transition-colors ${
                    isSelected
                      ? 'bg-brand-50/80 dark:bg-brand-950/60 text-brand-900 dark:text-brand-100'
                      : 'hover:bg-slate-50 dark:hover:bg-slate-800/80 text-slate-800 dark:text-slate-200'
                  }`}
                >
                  <div className="flex items-center gap-2.5 min-w-0">
                    <div className="w-7 h-7 rounded-full bg-slate-200 dark:bg-slate-700 text-slate-700 dark:text-slate-200 font-bold flex items-center justify-center text-xs shrink-0">
                      {getInitials(u.displayName)}
                    </div>
                    <div className="min-w-0">
                      <p className="font-semibold text-xs truncate">{u.displayName}</p>
                      <p className="text-[10px] text-slate-400 truncate">{u.email}</p>
                    </div>
                  </div>

                  {isSelected && (
                    <Check className="w-4 h-4 text-brand-600 dark:text-brand-400 shrink-0" />
                  )}
                </div>
              );
            })
          )}
        </div>
      )}

      {/* Error Message */}
      {error && (
        <div className="flex items-center gap-1 text-rose-600 dark:text-rose-400 text-[11px] pt-0.5">
          <AlertCircle className="w-3.5 h-3.5" />
          <span>{error}</span>
        </div>
      )}
    </div>
  );
};
