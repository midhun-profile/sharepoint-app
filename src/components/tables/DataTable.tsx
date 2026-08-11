import React, { useState, useRef, useEffect, useCallback } from 'react';
import { ColumnConfig, FieldType } from '../../types';
import { ImageCell, ImageItem } from './cells/ImageCell';
import { ImagePreviewModal } from '../dialogs/ImagePreviewModal';
import {
  ArrowUpDown,
  ArrowUp,
  ArrowDown,
  Edit2,
  Trash2,
  Eye,
  Check,
  X,
  Paperclip,
  Image as ImageIcon,
  User,
  Calendar,
  Layers,
} from 'lucide-react';
import './styles.css';

/**
 * Generic DataTable Props Contract
 */
export interface DataTableProps<TData extends Record<string, any>> {
  /** Configured columns mapping metadata to fields */
  columns: ColumnConfig[];
  /** Generic transactional dataset array */
  data: TData[];
  /** Loading state indicator */
  isLoading?: boolean;
  /** Set of selected row primary keys */
  selectedIds?: Set<string>;
  /** Selection toggle all callback */
  onSelectAll?: (selected: boolean) => void;
  /** Single row selection toggle callback */
  onSelectRow?: (id: string, selected: boolean) => void;
  /** Row edit action handler */
  onEditRow?: (item: TData) => void;
  /** Row delete action handler */
  onDeleteRow?: (item: TData) => void;
  /** Row detail view action handler */
  onViewRow?: (item: TData) => void;
  /** Column header sort handler */
  onSort?: (columnKey: string) => void;
  /** Active sorted column key */
  sortField?: string;
  /** Active sort direction ('asc' | 'desc') */
  sortDirection?: 'asc' | 'desc';
  /** Primary identifier key field name in dataset (defaults to 'id') */
  primaryKey?: string;
  /** Enables/disables multi-selection checkbox column */
  enableSelection?: boolean;
  /** Enables/disables row action buttons column */
  enableActions?: boolean;
  /** Empty state message */
  emptyMessage?: string;
}

/**
 * Staff-Architect Grade Generic DataTable Component
 *
 * Implements a 3-section CSS Sticky Grid Architecture:
 * - Section A (Left Sticky Column, z-index: 10 / z-index: 20 header): Multi-selection checkboxes
 * - Section B (Middle Scrollable Columns): Flexible dynamic columns scrolling horizontally
 * - Section C (Right Sticky Column, z-index: 10 / z-index: 20 header): CRUD Action controls
 *
 * Intersection headers (Section A & C top) use z-index: 20 to ensure scrolling body content
 * passes under both horizontal and vertical sticky boundaries seamlessly.
 */
export function DataTable<TData extends Record<string, any>>({
  columns,
  data,
  isLoading = false,
  selectedIds = new Set(),
  onSelectAll,
  onSelectRow,
  onEditRow,
  onDeleteRow,
  onViewRow,
  onSort,
  sortField,
  sortDirection,
  primaryKey = 'id',
  enableSelection = true,
  enableActions = true,
  emptyMessage = 'No records found matching your query.',
}: DataTableProps<TData>): React.ReactElement {
  const containerRef = useRef<HTMLDivElement>(null);

  // Dynamic horizontal scroll indicators state (boundary shadows)
  const [scrolledLeft, setScrolledLeft] = useState(false);
  const [scrolledRight, setScrolledRight] = useState(false);

  // Check scroll position to dynamically apply boundary drop-shadows on sticky edges
  const checkScrollBoundaries = useCallback(() => {
    const container = containerRef.current;
    if (!container) return;

    const { scrollLeft, scrollWidth, clientWidth } = container;
    setScrolledLeft(scrollLeft > 2);
    setScrolledRight(scrollLeft + clientWidth < scrollWidth - 2);
  }, []);

  useEffect(() => {
    const container = containerRef.current;
    if (!container) return;

    checkScrollBoundaries();

    // Attach resize observer to recalculate scroll shadows on window or layout resize
    const resizeObserver = new ResizeObserver(() => {
      checkScrollBoundaries();
    });

    resizeObserver.observe(container);
    return () => resizeObserver.disconnect();
  }, [checkScrollBoundaries, data, columns]);

  // Filter visible columns configured in SharePoint metadata
  const visibleColumns = columns.filter((col) => col.visible !== false);

  // Derive selection state helpers
  const allRowKeys = data.map((item) => String(item[primaryKey] ?? ''));
  const allSelected =
    allRowKeys.length > 0 && allRowKeys.every((key) => selectedIds.has(key));
  const isIndeterminate =
    !allSelected && allRowKeys.some((key) => selectedIds.has(key));

  // Helper renderer for dynamic field data types
  const renderCellContent = (item: TData, col: ColumnConfig) => {
    const rawVal = item[col.key];

    if (rawVal === null || rawVal === undefined || rawVal === '') {
      return <span className="text-slate-300 dark:text-slate-600 font-mono">—</span>;
    }

    switch (col.type) {
      case 'boolean':
        return Boolean(rawVal) ? (
          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold bg-emerald-50 text-emerald-700 dark:bg-emerald-950/60 dark:text-emerald-300 border border-emerald-200 dark:border-emerald-800">
            <Check className="w-3 h-3" /> True
          </span>
        ) : (
          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold bg-slate-100 text-slate-500 dark:bg-slate-800 dark:text-slate-400">
            <X className="w-3 h-3" /> False
          </span>
        );

      case 'choice':
        return (
          <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-brand-50 text-brand-700 dark:bg-brand-950/60 dark:text-brand-300 border border-brand-200/80 dark:border-brand-800/80">
            {String(rawVal)}
          </span>
        );

      case 'multichoice':
        const choices = Array.isArray(rawVal) ? rawVal : [String(rawVal)];
        return (
          <div className="flex flex-wrap gap-1">
            {choices.map((c, idx) => (
              <span
                key={idx}
                className="inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-medium bg-slate-100 text-slate-700 dark:bg-slate-800 dark:text-slate-300"
              >
                {c}
              </span>
            ))}
          </div>
        );

      case 'date':
      case 'datetime':
        try {
          const d = new Date(rawVal);
          return (
            <span className="inline-flex items-center gap-1.5 text-slate-600 dark:text-slate-300 font-mono text-[11px]">
              <Calendar className="w-3 h-3 text-slate-400 shrink-0" />
              {d.toLocaleDateString(undefined, {
                year: 'numeric',
                month: 'short',
                day: 'numeric',
              })}
            </span>
          );
        } catch {
          return String(rawVal);
        }

      case 'person':
        const name = typeof rawVal === 'object' ? rawVal?.displayName || rawVal?.title : String(rawVal);
        return (
          <span className="inline-flex items-center gap-1.5 font-medium text-slate-800 dark:text-slate-200">
            <div className="w-5 h-5 rounded-full bg-slate-200 dark:bg-slate-700 text-slate-600 dark:text-slate-300 flex items-center justify-center text-[10px] font-bold">
              <User className="w-3 h-3" />
            </div>
            <span className="truncate max-w-[140px]">{name}</span>
          </span>
        );

      case 'attachment':
        const count = Array.isArray(rawVal) ? rawVal.length : 1;
        return (
          <span className="inline-flex items-center gap-1 text-slate-600 dark:text-slate-300 font-medium text-[11px]">
            <Paperclip className="w-3.5 h-3.5 text-slate-400" />
            {count} file{count > 1 ? 's' : ''}
          </span>
        );

      case 'image':
        return (
          <div className="w-7 h-7 rounded-lg bg-slate-100 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 overflow-hidden flex items-center justify-center">
            {typeof rawVal === 'string' && rawVal.startsWith('http') ? (
              <img src={rawVal} alt="Thumbnail" className="w-full h-full object-cover" />
            ) : (
              <ImageIcon className="w-3.5 h-3.5 text-slate-400" />
            )}
          </div>
        );

      case 'number':
        return <span className="font-mono text-slate-800 dark:text-slate-200">{Number(rawVal).toLocaleString()}</span>;

      default:
        return <span className="truncate max-w-[240px] inline-block">{String(rawVal)}</span>;
    }
  };

  return (
    <div
      ref={containerRef}
      onScroll={checkScrollBoundaries}
      className="data-table-container overflow-x-auto w-full relative border border-slate-200 dark:border-slate-800 rounded-2xl bg-white dark:bg-slate-900 shadow-xs max-h-[72vh] custom-scrollbar"
    >
      <table className="w-full text-left border-collapse border-spacing-0 text-xs">
        {/* ======================================================================== */}
        {/* TABLE HEAD - STICKY TOP HEADER ROW WITH CORNER Z-INDEX BOOST (Z-20)     */}
        {/* ======================================================================== */}
        <thead className="bg-slate-50 dark:bg-slate-800/90 text-slate-500 dark:text-slate-400 uppercase font-bold tracking-wider select-none sticky top-0 z-20 shadow-xs">
          <tr>
            {/* SECTION A: Left Sticky Checkbox Column Header (z-index: 20) */}
            {enableSelection && (
              <th
                className={`sticky left-0 top-0 z-20 bg-slate-50 dark:bg-slate-800 border-b border-slate-200 dark:border-slate-800 px-4 py-3.5 w-12 text-center transition-shadow ${
                  scrolledLeft ? 'sticky-left-shadow border-r border-slate-200/80 dark:border-slate-700/80' : ''
                }`}
              >
                <input
                  type="checkbox"
                  checked={allSelected}
                  ref={(el) => {
                    if (el) el.indeterminate = isIndeterminate;
                  }}
                  onChange={(e) => onSelectAll && onSelectAll(e.target.checked)}
                  className="w-4 h-4 rounded border-slate-300 dark:border-slate-700 text-brand-600 focus:ring-brand-500/30 cursor-pointer accent-brand-600"
                  aria-label="Select all rows"
                />
              </th>
            )}

            {/* SECTION B: Middle Flexible Scrollable Column Headers (z-index: 10) */}
            {visibleColumns.map((col) => {
              const isSorted = sortField === col.key;

              return (
                <th
                  key={col.key}
                  onClick={() => col.sortable && onSort && onSort(col.key)}
                  style={{ width: col.width ? `${col.width}px` : 'auto', minWidth: '120px' }}
                  className={`sticky top-0 z-10 bg-slate-50 dark:bg-slate-800 border-b border-slate-200 dark:border-slate-800 px-4 py-3.5 transition-colors ${
                    col.sortable ? 'cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-700/80 hover:text-slate-900 dark:hover:text-slate-100' : ''
                  }`}
                >
                  <div className="flex items-center gap-1.5">
                    <span className="truncate">{col.label}</span>
                    {col.sortable && (
                      <span className="shrink-0 text-slate-400">
                        {isSorted ? (
                          sortDirection === 'asc' ? (
                            <ArrowUp className="w-3.5 h-3.5 text-brand-600" />
                          ) : (
                            <ArrowDown className="w-3.5 h-3.5 text-brand-600" />
                          )
                        ) : (
                          <ArrowUpDown className="w-3 h-3 opacity-40 hover:opacity-100" />
                        )}
                      </span>
                    )}
                  </div>
                </th>
              );
            })}

            {/* SECTION C: Right Sticky Action Controls Column Header (z-index: 20) */}
            {enableActions && (
              <th
                className={`sticky right-0 top-0 z-20 bg-slate-50 dark:bg-slate-800 border-b border-slate-200 dark:border-slate-800 px-4 py-3.5 w-24 text-right transition-shadow ${
                  scrolledRight ? 'sticky-right-shadow border-l border-slate-200/80 dark:border-slate-700/80' : ''
                }`}
              >
                Actions
              </th>
            )}
          </tr>
        </thead>

        {/* ======================================================================== */}
        {/* TABLE BODY - STICKY LEFT/RIGHT CELL COVERS WITH OPAQUE BACKGROUNDS     */}
        {/* ======================================================================== */}
        <tbody className="divide-y divide-slate-100 dark:divide-slate-800/80 bg-white dark:bg-slate-900">
          {isLoading ? (
            /* Pulsing Loading Skeleton Rows */
            Array.from({ length: 5 }).map((_, rIdx) => (
              <tr key={rIdx} className="animate-pulse">
                {enableSelection && (
                  <td className="sticky left-0 bg-white dark:bg-slate-900 px-4 py-3.5 w-12 text-center">
                    <div className="w-4 h-4 bg-slate-200 dark:bg-slate-800 rounded mx-auto" />
                  </td>
                )}
                {visibleColumns.map((col) => (
                  <td key={col.key} className="px-4 py-3.5">
                    <div className="h-4 bg-slate-200 dark:bg-slate-800 rounded w-3/4" />
                  </td>
                ))}
                {enableActions && (
                  <td className="sticky right-0 bg-white dark:bg-slate-900 px-4 py-3.5 w-24 text-right">
                    <div className="h-6 w-16 bg-slate-200 dark:bg-slate-800 rounded ml-auto" />
                  </td>
                )}
              </tr>
            ))
          ) : data.length === 0 ? (
            /* Empty State Container */
            <tr>
              <td
                colSpan={visibleColumns.length + (enableSelection ? 1 : 0) + (enableActions ? 1 : 0)}
                className="py-12 px-4 text-center text-slate-400 dark:text-slate-500"
              >
                <div className="max-w-xs mx-auto space-y-2">
                  <Layers className="w-8 h-8 mx-auto text-slate-300 dark:text-slate-600 stroke-1" />
                  <p className="text-xs font-semibold text-slate-600 dark:text-slate-400">{emptyMessage}</p>
                </div>
              </td>
            </tr>
          ) : (
            /* Transactional Data Mapping */
            data.map((item, rowIdx) => {
              const rowKey = String(item[primaryKey] ?? rowIdx);
              const isSelected = selectedIds.has(rowKey);

              // Cell background MUST be strictly opaque to cover middle columns scrolling underneath
              const cellBg = isSelected
                ? 'bg-blue-50 dark:bg-slate-800'
                : 'bg-white dark:bg-slate-900 group-hover:bg-slate-50 dark:group-hover:bg-slate-800/80';

              return (
                <tr key={rowKey} className="group transition-colors">
                  {/* SECTION A: Left Sticky Checkbox Cell (z-index: 10) */}
                  {enableSelection && (
                    <td
                      className={`sticky left-0 z-10 ${cellBg} border-b border-slate-100 dark:border-slate-800 px-4 py-3 w-12 text-center transition-shadow ${
                        scrolledLeft ? 'sticky-left-shadow border-r border-slate-200/80 dark:border-slate-800' : ''
                      }`}
                    >
                      <input
                        type="checkbox"
                        checked={isSelected}
                        onChange={(e) => onSelectRow && onSelectRow(rowKey, e.target.checked)}
                        className="w-4 h-4 rounded border-slate-300 dark:border-slate-700 text-brand-600 focus:ring-brand-500/30 cursor-pointer accent-brand-600"
                        aria-label={`Select row ${rowKey}`}
                      />
                    </td>
                  )}

                  {/* SECTION B: Middle Scrollable Data Cells */}
                  {visibleColumns.map((col) => (
                    <td
                      key={col.key}
                      style={{ width: col.width ? `${col.width}px` : 'auto' }}
                      className="px-4 py-3 whitespace-nowrap text-slate-700 dark:text-slate-200 font-normal border-b border-slate-100 dark:border-slate-800/80"
                    >
                      {renderCellContent(item, col)}
                    </td>
                  ))}

                  {/* SECTION C: Right Sticky CRUD Action Cell (z-index: 10) */}
                  {enableActions && (
                    <td
                      className={`sticky right-0 z-10 ${cellBg} border-b border-slate-100 dark:border-slate-800 px-4 py-3 w-24 text-right transition-shadow ${
                        scrolledRight ? 'sticky-right-shadow border-l border-slate-200/80 dark:border-slate-800' : ''
                      }`}
                    >
                      <div className="flex items-center justify-end gap-1 opacity-90 group-hover:opacity-100">
                        {onViewRow && (
                          <button
                            onClick={() => onViewRow(item)}
                            className="p-1.5 rounded-lg text-slate-400 hover:text-brand-600 dark:hover:text-brand-400 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors"
                            title="View Details"
                          >
                            <Eye className="w-3.5 h-3.5" />
                          </button>
                        )}
                        {onEditRow && (
                          <button
                            onClick={() => onEditRow(item)}
                            className="p-1.5 rounded-lg text-slate-400 hover:text-amber-600 dark:hover:text-amber-400 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors"
                            title="Edit Item"
                          >
                            <Edit2 className="w-3.5 h-3.5" />
                          </button>
                        )}
                        {onDeleteRow && (
                          <button
                            onClick={() => onDeleteRow(item)}
                            className="p-1.5 rounded-lg text-slate-400 hover:text-rose-600 dark:hover:text-rose-400 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors"
                            title="Delete Item"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        )}
                      </div>
                    </td>
                  )}
                </tr>
              );
            })
          )}
        </tbody>
      </table>
    </div>
  );
}
