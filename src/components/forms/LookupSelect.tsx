import React, { useState, useRef, useEffect, useMemo } from 'react';
import { useQuery } from '@tanstack/react-query';
import { useAuth } from '../../hooks/useAuth';
import { graphService } from '../../services/graphService';
import { MOCK_ITEMS } from '../../services/mockData';
import {
  Database,
  Search,
  X,
  Check,
  ChevronDown,
  Loader2,
  AlertCircle,
  ExternalLink,
} from 'lucide-react';

export interface LookupItem {
  LookupId: string | number;
  LookupValue: string;
}

export interface LookupSelectProps {
  /** Target SharePoint site ID (defaults to root site if absent) */
  siteId?: string;
  /** Relational parent SharePoint List ID to query items from */
  targetListId: string;
  /** Primary display field key in target list (defaults to 'Title') */
  lookupDisplayField?: string;
  /** Selected Lookup value structure or array */
  value: any;
  /** Callback fired on selection change returning formatted SharePoint Lookup OData structure */
  onChange: (value: any) => void;
  /** Input placeholder string */
  placeholder?: string;
  /** Override flag to disable select */
  disabled?: boolean;
  /** Error message string */
  error?: string;
  /** Enables multi-lookup selection mode */
  isMulti?: boolean;
}

/**
 * Normalizes incoming value prop into an array of LookupItem
 */
function normalizeLookupValue(val: any): LookupItem[] {
  if (!val) return [];
  const list = Array.isArray(val) ? val : [val];
  return list
    .filter(Boolean)
    .map((item, idx) => {
      if (typeof item === 'string' || typeof item === 'number') {
        return { LookupId: item, LookupValue: String(item) };
      }
      return {
        LookupId: item.LookupId ?? item.id ?? item.Id ?? idx,
        LookupValue: String(
          item.LookupValue ??
            item.title ??
            item.Title ??
            item.name ??
            item.SKU ??
            item.FullName ??
            item.LookupId ??
            'Selected Item'
        ),
      };
    });
}

/**
 * Enterprise LookupSelect Component
 *
 * Resolves relational items from a parent SharePoint list using Microsoft Graph OData queries
 * (`$select=id,fields&$expand=fields($select=Id,${lookupDisplayField})&$top=100`).
 * Powered by TanStack Query for caching and optimal client performance.
 */
export const LookupSelect: React.FC<LookupSelectProps> = ({
  siteId,
  targetListId,
  lookupDisplayField = 'Title',
  value,
  onChange,
  placeholder = 'Select from parent list...',
  disabled = false,
  error,
  isMulti = false,
}) => {
  const { getAccessToken } = useAuth();
  const [isOpen, setIsOpen] = useState(false);
  const [filterText, setFilterText] = useState('');
  const containerRef = useRef<HTMLDivElement>(null);

  const selectedLookups = useMemo(() => normalizeLookupValue(value), [value]);

  // Fetch parent SharePoint List items via TanStack Query
  const {
    data: lookupOptions = [],
    isLoading,
    isError,
    error: queryError,
  } = useQuery<LookupItem[]>({
    queryKey: ['lookupListItems', siteId, targetListId, lookupDisplayField],
    queryFn: async () => {
      if (!targetListId) return [];

      const token = await getAccessToken();

      // If no valid live token or mock mode, fall back to graphService or MOCK_ITEMS
      if (!token || token === 'mock-demo-bearer-token' || !siteId) {
        const mockData = MOCK_ITEMS[targetListId];
        if (mockData && mockData.length > 0) {
          return mockData.map((item) => {
            const displayVal =
              item.fields[lookupDisplayField] ||
              item.fields.Title ||
              item.fields.SKU ||
              item.fields.FullName ||
              item.title ||
              `Item #${item.id}`;
            return {
              LookupId: item.id,
              LookupValue: String(displayVal),
            };
          });
        }

        // Fallback demo options if specific mock key is missing
        return [
          { LookupId: '101', LookupValue: 'Enterprise Intranet Overhaul (Project Alpha)' },
          { LookupId: '102', LookupValue: 'Supply Chain Barcode Scanner (Project Beta)' },
          { LookupId: '103', LookupValue: 'Global Azure Security Compliance (Project Gamma)' },
          { LookupId: '104', LookupValue: 'High-Speed Fiber Transceiver SFP+ (SKU-88210)' },
          { LookupId: '105', LookupValue: 'Industrial Gigabit Ethernet Switch (SKU-99042)' },
        ];
      }

      // Live Microsoft Graph API OData Query
      const effectiveSite = siteId || 'root';
      const cleanField = encodeURIComponent(lookupDisplayField);
      const url = `https://graph.microsoft.com/v1.0/sites/${effectiveSite}/lists/${targetListId}/items?expand=fields($select=id,${cleanField},Title)&$top=100`;

      const response = await fetch(url, {
        headers: { Authorization: `Bearer ${token}` },
      });

      if (!response.ok) {
        console.warn(`Graph API Lookup fetch failed (${response.status}), falling back to service layer`);
        const fallbackItems = await graphService.getListItems(effectiveSite, targetListId);
        return fallbackItems.map((item) => {
          const displayVal =
            item.fields[lookupDisplayField] || item.fields.Title || item.title || `Record #${item.id}`;
          return {
            LookupId: item.id,
            LookupValue: String(displayVal),
          };
        });
      }

      const data = await response.json();
      return (data.value || []).map((item: any) => {
        const fields = item.fields || {};
        const displayVal = fields[lookupDisplayField] || fields.Title || item.id;
        return {
          LookupId: item.id,
          LookupValue: String(displayVal),
        };
      });
    },
    staleTime: 5 * 60 * 1000, // Cache lookup results for 5 minutes
    retry: 1,
  });

  // Dismiss dropdown on click outside
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Filter available options
  const filteredOptions = useMemo(() => {
    if (!filterText.trim()) return lookupOptions;
    return lookupOptions.filter((opt) =>
      opt.LookupValue.toLowerCase().includes(filterText.toLowerCase())
    );
  }, [lookupOptions, filterText]);

  // Handle selecting an option
  const handleSelectOption = (item: LookupItem) => {
    if (isMulti) {
      const exists = selectedLookups.some((s) => String(s.LookupId) === String(item.LookupId));
      let updated: LookupItem[];
      if (exists) {
        updated = selectedLookups.filter((s) => String(s.LookupId) !== String(item.LookupId));
      } else {
        updated = [...selectedLookups, item];
      }
      onChange(updated);
    } else {
      onChange(item);
      setIsOpen(false);
      setFilterText('');
    }
  };

  // Remove single item from selection
  const handleRemoveOption = (lookupId: string | number, e?: React.MouseEvent) => {
    e?.stopPropagation();
    if (disabled) return;

    if (isMulti) {
      const updated = selectedLookups.filter((s) => String(s.LookupId) !== String(lookupId));
      onChange(updated);
    } else {
      onChange(null);
    }
  };

  return (
    <div ref={containerRef} className="relative w-full space-y-1.5">
      {/* Trigger Button & Selected Pills Container */}
      <div
        onClick={() => !disabled && setIsOpen(!isOpen)}
        className={`min-h-[42px] px-3 py-2 rounded-xl border transition-all cursor-pointer flex flex-wrap items-center gap-2 select-none ${
          error
            ? 'border-rose-500 bg-rose-50/30 dark:bg-rose-950/20'
            : isOpen
            ? 'border-brand-500 ring-2 ring-brand-500/30 bg-white dark:bg-slate-950'
            : 'border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-950 hover:border-slate-300'
        } ${disabled ? 'opacity-60 cursor-not-allowed bg-slate-100 dark:bg-slate-900' : ''}`}
      >
        <Database className="w-4 h-4 text-slate-400 shrink-0" />

        {/* Render Selected Items */}
        {selectedLookups.length > 0 ? (
          selectedLookups.map((item) => (
            <span
              key={item.LookupId}
              className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-xs font-semibold bg-emerald-50 text-emerald-700 dark:bg-emerald-950 dark:text-emerald-300 border border-emerald-200 dark:border-emerald-800 shrink-0"
            >
              <span className="truncate max-w-[200px]">{item.LookupValue}</span>
              {!disabled && (
                <button
                  type="button"
                  onClick={(e) => handleRemoveOption(item.LookupId, e)}
                  className="p-0.5 rounded-full hover:bg-emerald-200/60 dark:hover:bg-emerald-800 text-emerald-600 transition-colors"
                  title="Deselect item"
                >
                  <X className="w-3 h-3" />
                </button>
              )}
            </span>
          ))
        ) : (
          <span className="text-xs text-slate-400 font-normal flex-1">
            {placeholder}
          </span>
        )}

        {/* Dropdown State Icons */}
        <div className="ml-auto flex items-center gap-1 shrink-0">
          {isLoading ? (
            <Loader2 className="w-3.5 h-3.5 text-brand-600 animate-spin" />
          ) : (
            <ChevronDown
              className={`w-3.5 h-3.5 text-slate-400 transition-transform ${
                isOpen ? 'rotate-180 text-brand-600' : ''
              }`}
            />
          )}
        </div>
      </div>

      {/* Dropdown Options Popover Menu */}
      {isOpen && !disabled && (
        <div className="absolute left-0 right-0 mt-1 rounded-2xl bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 shadow-2xl p-2 z-50 text-xs animate-fade-in space-y-1.5">
          {/* Inner Search Filter Input */}
          <div className="relative p-1">
            <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
            <input
              type="text"
              value={filterText}
              onChange={(e) => setFilterText(e.target.value)}
              placeholder="Filter list options..."
              className="w-full pl-8 pr-3 py-1.5 text-xs rounded-xl bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 focus:outline-none focus:border-brand-500 text-slate-900 dark:text-slate-100"
              autoFocus
            />
          </div>

          {/* List Options Menu */}
          <div className="max-h-52 overflow-y-auto custom-scrollbar space-y-1 pt-1">
            {isLoading ? (
              <div className="py-6 flex items-center justify-center gap-2 text-slate-400">
                <Loader2 className="w-4 h-4 animate-spin text-brand-600" />
                <span>Loading relational SharePoint items...</span>
              </div>
            ) : isError ? (
              <div className="p-3 rounded-xl bg-rose-50 dark:bg-rose-950/40 text-rose-600 dark:text-rose-400 text-center flex items-center justify-center gap-2">
                <AlertCircle className="w-4 h-4 shrink-0" />
                <span>Failed to query lookup list items.</span>
              </div>
            ) : filteredOptions.length === 0 ? (
              <div className="py-6 text-center text-slate-400">
                No lookup records match your filter.
              </div>
            ) : (
              filteredOptions.map((opt) => {
                const isSelected = selectedLookups.some(
                  (sel) => String(sel.LookupId) === String(opt.LookupId)
                );

                return (
                  <div
                    key={opt.LookupId}
                    onClick={() => handleSelectOption(opt)}
                    className={`flex items-center justify-between px-3 py-2 rounded-xl cursor-pointer transition-colors ${
                      isSelected
                        ? 'bg-emerald-50 dark:bg-emerald-950/60 text-emerald-900 dark:text-emerald-100 font-semibold'
                        : 'hover:bg-slate-50 dark:hover:bg-slate-800/80 text-slate-800 dark:text-slate-200'
                    }`}
                  >
                    <div className="flex items-center gap-2 truncate">
                      <ExternalLink className="w-3 h-3 text-slate-400 shrink-0" />
                      <span className="truncate">{opt.LookupValue}</span>
                    </div>

                    {isSelected && (
                      <Check className="w-4 h-4 text-emerald-600 dark:text-emerald-400 shrink-0" />
                    )}
                  </div>
                );
              })
            )}
          </div>
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
