import React, { useState, useMemo } from 'react';
import { MenuConfig, SharePointColumnDefinition, SharePointListItem } from '../../types';
import {
  Search,
  ArrowUpDown,
  Plus,
  Download,
  Trash2,
  Edit,
  Eye,
  Columns,
  ChevronLeft,
  ChevronRight,
  Copy,
  CheckCircle,
  XCircle,
  FileSpreadsheet,
  X,
  Image as ImageIcon,
} from 'lucide-react';

interface DynamicTableProps {
  menu: MenuConfig;
  columns: SharePointColumnDefinition[];
  items: SharePointListItem[];
  isLoading: boolean;
  onAddItem: () => void;
  onEditItem: (item: SharePointListItem) => void;
  onDeleteItem: (itemId: string) => void;
  onBulkDeleteItems?: (itemIds: string[]) => void;
  onBulkEditItems?: (itemIds: string[], updateFields: Record<string, any>) => void;
  onDuplicateItem: (item: SharePointListItem) => void;
  onSelectItem: (item: SharePointListItem) => void;
  canCreate: boolean;
  canUpdate: boolean;
  canDelete: boolean;
  canExport: boolean;
}

export const DynamicTable: React.FC<DynamicTableProps> = ({
  menu,
  columns,
  items,
  isLoading,
  onAddItem,
  onEditItem,
  onDeleteItem,
  onBulkDeleteItems,
  onBulkEditItems,
  onDuplicateItem,
  onSelectItem,
  canCreate,
  canUpdate,
  canDelete,
  canExport,
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [sortColumn, setSortColumn] = useState<string>(menu.defaultSortColumn || 'Title');
  const [sortDirection, setSortDirection] = useState<'asc' | 'desc'>(menu.defaultSortDirection || 'asc');
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [currentPage, setCurrentPage] = useState(1);
  const [visibleColNames, setVisibleColNames] = useState<string[]>(menu.visibleColumns);
  const [isColumnPickerOpen, setIsColumnPickerOpen] = useState(false);
  const [previewImage, setPreviewImage] = useState<string | null>(null);

  // Bulk Edit Modal state
  const [isBulkEditModalOpen, setIsBulkEditModalOpen] = useState(false);
  const [enabledBulkFields, setEnabledBulkFields] = useState<Set<string>>(new Set());
  const [bulkFieldValues, setBulkFieldValues] = useState<Record<string, any>>({});

  // Filter columns to display (up to 20 columns supported)
  const activeColumns = useMemo(() => {
    return columns.filter((col) => visibleColNames.includes(col.name));
  }, [columns, visibleColNames]);

  // Search & Filter Items
  const filteredItems = useMemo(() => {
    let result = [...items];

    if (searchTerm.trim()) {
      const term = searchTerm.toLowerCase();
      result = result.filter((item) => {
        return activeColumns.some((col) => {
          const val = item.fields[col.name];
          if (val === undefined || val === null) return false;
          return String(val).toLowerCase().includes(term);
        });
      });
    }

    // Sort
    result.sort((a, b) => {
      const valA = a.fields[sortColumn] ?? '';
      const valB = b.fields[sortColumn] ?? '';

      if (valA < valB) return sortDirection === 'asc' ? -1 : 1;
      if (valA > valB) return sortDirection === 'asc' ? 1 : -1;
      return 0;
    });

    return result;
  }, [items, searchTerm, sortColumn, sortDirection, activeColumns]);

  // Pagination calculations
  const pageSize = menu.pageSize || 10;
  const totalPages = Math.ceil(filteredItems.length / pageSize) || 1;
  const paginatedItems = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return filteredItems.slice(start, start + pageSize);
  }, [filteredItems, currentPage, pageSize]);

  // Handle Sort Toggle
  const handleSort = (colName: string) => {
    if (sortColumn === colName) {
      setSortDirection(sortDirection === 'asc' ? 'desc' : 'asc');
    } else {
      setSortColumn(colName);
      setSortDirection('asc');
    }
  };

  // Bulk Selection Handlers
  const toggleSelectAll = () => {
    if (selectedIds.size === paginatedItems.length) {
      setSelectedIds(new Set());
    } else {
      setSelectedIds(new Set(paginatedItems.map((i) => i.id)));
    }
  };

  const toggleSelectOne = (id: string) => {
    const next = new Set(selectedIds);
    if (next.has(id)) {
      next.delete(id);
    } else {
      next.add(id);
    }
    setSelectedIds(next);
  };

  // Export to CSV
  const handleExportCSV = () => {
    const exportData = filteredItems.map((item) => {
      const row: Record<string, any> = {};
      activeColumns.forEach((col) => {
        row[col.displayName] = item.fields[col.name] ?? '';
      });
      return row;
    });

    if (exportData.length === 0) return;

    const headers = Object.keys(exportData[0]).join(',');
    const rows = exportData.map((r) => Object.values(r).map((v) => `"${String(v).replace(/"/g, '""')}"`).join(','));
    const csvContent = 'data:text/csv;charset=utf-8,' + [headers, ...rows].join('\n');

    const encodedUri = encodeURI(csvContent);
    const link = document.createElement('a');
    link.setAttribute('href', encodedUri);
    link.setAttribute('download', `${menu.name.replace(/\s+/g, '_')}_Export.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  // Bulk Delete
  const handleBulkDelete = () => {
    if (selectedIds.size === 0) return;
    const idsArray = Array.from(selectedIds);
    if (onBulkDeleteItems) {
      onBulkDeleteItems(idsArray);
    } else {
      idsArray.forEach((id) => onDeleteItem(id));
    }
    setSelectedIds(new Set());
  };

  // Bulk Edit Handlers
  const handleOpenBulkEdit = () => {
    setEnabledBulkFields(new Set());
    setBulkFieldValues({});
    setIsBulkEditModalOpen(true);
  };

  const handleApplyBulkEdit = () => {
    if (enabledBulkFields.size === 0 || selectedIds.size === 0) return;
    const updates: Record<string, any> = {};
    enabledBulkFields.forEach((fieldName) => {
      updates[fieldName] = bulkFieldValues[fieldName] ?? '';
    });

    if (onBulkEditItems) {
      onBulkEditItems(Array.from(selectedIds), updates);
    }
    setIsBulkEditModalOpen(false);
    setSelectedIds(new Set());
  };

  return (
    <div className="flex flex-col h-full bg-white dark:bg-slate-900 rounded-xl border border-slate-200 dark:border-slate-800 shadow-enterprise overflow-hidden">
      {/* Table Toolbar Header */}
      <div className="p-4 border-b border-slate-200 dark:border-slate-800 flex flex-col md:flex-row md:items-center justify-between gap-3 bg-slate-50/50 dark:bg-slate-950/40 relative z-40">
        {/* Left: Search & Filters */}
        <div className="flex items-center gap-2 flex-1 max-w-lg">
          <div className="relative flex-1">
            <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
            <input
              type="text"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              placeholder={`Search in ${menu.name}...`}
              className="w-full pl-9 pr-3 py-1.5 text-xs rounded-lg border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 text-slate-900 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-brand-500/50"
            />
          </div>

          {/* Column Display Config Modal Toggle */}
          <div className="relative">
            <button
              onClick={() => setIsColumnPickerOpen(!isColumnPickerOpen)}
              className="px-2.5 py-1.5 rounded-lg border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-800 text-xs font-semibold flex items-center gap-1.5 transition-colors"
              title="Customize visible columns"
            >
              <Columns className="w-3.5 h-3.5" />
              <span className="hidden sm:inline">Columns</span>
            </button>

            {isColumnPickerOpen && (
              <div className="absolute left-0 md:right-0 md:left-auto mt-2 w-56 p-3 rounded-xl bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 shadow-xl z-50 animate-fade-in text-xs space-y-2">
                <div className="font-bold text-slate-500 uppercase text-[10px] pb-1 border-b border-slate-100 dark:border-slate-800">
                  Visible Table Columns
                </div>
                <div className="max-h-48 overflow-y-auto space-y-1">
                  {columns.map((col) => (
                    <label
                      key={col.id}
                      className="flex items-center gap-2 px-1.5 py-1 rounded hover:bg-slate-50 dark:hover:bg-slate-800 cursor-pointer"
                    >
                      <input
                        type="checkbox"
                        checked={visibleColNames.includes(col.name)}
                        onChange={(e) => {
                          if (e.target.checked) {
                            setVisibleColNames([...visibleColNames, col.name]);
                          } else {
                            setVisibleColNames(visibleColNames.filter((n) => n !== col.name));
                          }
                        }}
                        className="w-3.5 h-3.5 rounded border-slate-300 text-brand-600 focus:ring-brand-500"
                      />
                      <span className="text-slate-700 dark:text-slate-300 truncate">{col.displayName}</span>
                    </label>
                  ))}
                </div>
              </div>
            )}
          </div>
        </div>

        {/* Right: Actions */}
        <div className="flex items-center gap-2 self-end md:self-auto">
          {selectedIds.size > 0 && canUpdate && (
            <button
              onClick={handleOpenBulkEdit}
              className="px-3 py-1.5 rounded-lg bg-amber-50 text-amber-700 dark:bg-amber-950/60 dark:text-amber-300 border border-amber-200 dark:border-amber-900 text-xs font-semibold flex items-center gap-1.5 hover:bg-amber-100 transition-colors"
            >
              <Edit className="w-3.5 h-3.5" />
              <span>Bulk Edit ({selectedIds.size})</span>
            </button>
          )}

          {selectedIds.size > 0 && canDelete && (
            <button
              onClick={handleBulkDelete}
              className="px-3 py-1.5 rounded-lg bg-red-50 text-red-600 dark:bg-red-950/60 dark:text-red-300 border border-red-200 dark:border-red-900 text-xs font-semibold flex items-center gap-1.5 hover:bg-red-100 transition-colors"
            >
              <Trash2 className="w-3.5 h-3.5" />
              <span>Delete ({selectedIds.size})</span>
            </button>
          )}

          {canExport && (
            <button
              onClick={handleExportCSV}
              className="px-3 py-1.5 rounded-lg border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-800 text-xs font-semibold flex items-center gap-1.5 transition-colors"
            >
              <Download className="w-3.5 h-3.5" />
              <span className="hidden sm:inline">Export CSV</span>
            </button>
          )}

          {canCreate && (
            <button
              onClick={onAddItem}
              className="px-3.5 py-1.5 rounded-lg bg-brand-600 hover:bg-brand-700 text-white shadow-sm text-xs font-bold flex items-center gap-1.5 transition-colors"
            >
              <Plus className="w-4 h-4" />
              <span>New Record</span>
            </button>
          )}
        </div>
      </div>

      {/* Table Body Container */}
      <div className="flex-1 overflow-x-auto overflow-y-auto min-h-[300px]">
        {activeColumns.length === 0 ? (
          <div className="flex flex-col items-center justify-center p-12 text-center text-slate-400 min-h-[300px]">
            <Columns className="w-10 h-10 mx-auto mb-2 opacity-40" />
            <p className="text-sm font-semibold text-slate-600 dark:text-slate-300">
              No columns added or visible
            </p>
            <p className="text-xs text-slate-400 max-w-sm mt-1">
              Select columns from the Column Selector or configure columns in Menu Settings.
            </p>
          </div>
        ) : (
          <table className="w-full text-left border-separate border-spacing-0 text-xs min-w-full">
          {/* Sticky Table Header */}
          <thead className="bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-300 font-bold">
            <tr>
              <th className="sticky top-0 left-0 z-30 p-3 w-12 min-w-[48px] max-w-[48px] text-center bg-slate-100 dark:bg-slate-800 border-b border-r border-slate-200 dark:border-slate-800">
                <input
                  type="checkbox"
                  checked={selectedIds.size > 0 && selectedIds.size === paginatedItems.length}
                  onChange={toggleSelectAll}
                  className="w-3.5 h-3.5 rounded border-slate-300 text-brand-600 focus:ring-brand-500"
                />
              </th>
              {activeColumns.map((col, colIdx) => (
                <th
                  key={col.id}
                  onClick={() => handleSort(col.name)}
                  className={`sticky top-0 p-3 cursor-pointer hover:bg-slate-200/50 dark:hover:bg-slate-700/50 transition-colors whitespace-nowrap min-w-[150px] max-w-[240px] border-b border-slate-200 dark:border-slate-800 ${
                    colIdx === 0
                      ? 'left-[48px] z-30 bg-slate-100 dark:bg-slate-800 border-r border-slate-200 dark:border-slate-800 shadow-[2px_0_5px_-2px_rgba(0,0,0,0.08)]'
                      : 'z-10 bg-slate-100 dark:bg-slate-800'
                  }`}
                >
                  <div className="flex items-center gap-1.5">
                    <span>{col.displayName}</span>
                    <ArrowUpDown
                      className={`w-3 h-3 ${
                        sortColumn === col.name ? 'text-brand-600 font-bold' : 'opacity-40'
                      }`}
                    />
                  </div>
                </th>
              ))}
              <th className="sticky top-0 right-0 z-30 p-3 w-32 text-right whitespace-nowrap bg-slate-100 dark:bg-slate-800 border-b border-l border-slate-200 dark:border-slate-800 shadow-[-2px_0_5px_-2px_rgba(0,0,0,0.08)]">
                Actions
              </th>
            </tr>
          </thead>

          {/* Table Data Rows */}
          <tbody className="divide-y divide-slate-100 dark:divide-slate-800/80">
            {isLoading ? (
              // Loading skeletons
              Array.from({ length: 5 }).map((_, idx) => (
                <tr key={idx} className="animate-pulse">
                  <td className="sticky left-0 z-20 p-3 w-12 min-w-[48px] max-w-[48px] text-center bg-white dark:bg-slate-900 border-r border-b border-slate-200 dark:border-slate-800">
                    <div className="w-3.5 h-3.5 bg-slate-200 dark:bg-slate-800 rounded mx-auto" />
                  </td>
                  {activeColumns.map((col, colIdx) => (
                    <td
                      key={col.id}
                      className={`p-3 border-b border-slate-100 dark:border-slate-800 ${
                        colIdx === 0
                          ? 'sticky left-[48px] z-20 bg-white dark:bg-slate-900 border-r border-slate-200 dark:border-slate-800'
                          : ''
                      }`}
                    >
                      <div className="h-3 bg-slate-200 dark:bg-slate-800 rounded w-24" />
                    </td>
                  ))}
                  <td className="sticky right-0 z-20 p-3 bg-white dark:bg-slate-900 border-l border-b border-slate-200 dark:border-slate-800">
                    <div className="h-3 bg-slate-200 dark:bg-slate-800 rounded w-12 ml-auto" />
                  </td>
                </tr>
              ))
            ) : paginatedItems.length === 0 ? (
              // Empty State
              <tr>
                <td colSpan={activeColumns.length + 2} className="p-12 text-center text-slate-400">
                  <FileSpreadsheet className="w-10 h-10 mx-auto mb-2 opacity-40" />
                  <p className="text-sm font-semibold text-slate-600 dark:text-slate-300">
                    No records found
                  </p>
                  <p className="text-xs text-slate-400">
                    Try adjusting your search query or add a new record.
                  </p>
                </td>
              </tr>
            ) : (
              paginatedItems.map((item) => {
                const isSelected = selectedIds.has(item.id);
                const cellBg = isSelected
                  ? 'bg-blue-50 dark:bg-slate-800'
                  : 'bg-white dark:bg-slate-900 group-hover:bg-slate-50 dark:group-hover:bg-slate-800';

                return (
                  <tr key={item.id} className="group transition-colors">
                    {/* Sticky Checkbox Column */}
                    <td
                      className={`sticky left-0 z-20 p-3 w-12 min-w-[48px] max-w-[48px] text-center border-b border-r border-slate-200 dark:border-slate-800/80 ${cellBg}`}
                    >
                      <input
                        type="checkbox"
                        checked={isSelected}
                        onChange={() => toggleSelectOne(item.id)}
                        className="w-3.5 h-3.5 rounded border-slate-300 text-brand-600 focus:ring-brand-500"
                      />
                    </td>

                    {/* Data Columns (Supporting up to 20 columns + Images) */}
                    {activeColumns.map((col, colIdx) => {
                      const val = item.fields[col.name];

                      // Detect image column or image URL
                      const isImageCol = col.type === 'Image' || col.type === 'image';
                      const isImageUrl =
                        typeof val === 'string' &&
                        (val.startsWith('http://') ||
                          val.startsWith('https://') ||
                          val.startsWith('data:image/') ||
                          /\.(jpg|jpeg|png|webp|gif|svg)(\?.*)?$/i.test(val));

                      const isStickyFirst = colIdx === 0;

                      return (
                        <td
                          key={col.id}
                          onClick={() => onSelectItem(item)}
                          className={`p-3 cursor-pointer max-w-xs truncate text-slate-800 dark:text-slate-200 font-medium min-w-[150px] border-b border-slate-100 dark:border-slate-800/80 ${cellBg} ${
                            isStickyFirst
                              ? 'sticky left-[48px] z-20 border-r border-slate-200 dark:border-slate-800/80 shadow-[2px_0_5px_-2px_rgba(0,0,0,0.08)]'
                              : ''
                          }`}
                        >
                          {val === undefined || val === null || val === '' ? (
                            <span className="text-slate-400 font-normal italic">-</span>
                          ) : isImageCol || isImageUrl ? (
                            <div
                              className="flex items-center gap-2 group/img cursor-pointer"
                              onClick={(e) => {
                                e.stopPropagation();
                                if (typeof val === 'string') setPreviewImage(val);
                              }}
                            >
                              <img
                                src={String(val)}
                                alt={col.displayName}
                                className="w-8 h-8 rounded-lg object-cover border border-slate-200 dark:border-slate-700 shadow-2xs group-hover/img:scale-105 transition-transform shrink-0"
                                onError={(e) => {
                                  (e.currentTarget as HTMLElement).style.display = 'none';
                                }}
                              />
                              <span className="truncate max-w-[120px] text-xs font-medium text-slate-700 dark:text-slate-300">
                                {typeof val === 'string'
                                  ? val.split('/').pop()?.substring(0, 15) || 'Image'
                                  : 'Image'}
                              </span>
                            </div>
                          ) : col.type === 'Choice' ? (
                            <span className="inline-flex px-2 py-0.5 rounded-full text-[11px] font-semibold bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 border border-slate-200 dark:border-slate-700">
                              {String(val)}
                            </span>
                          ) : String(col.type).toLowerCase() === 'datetime' || String(col.type).toLowerCase() === 'date' ? (
                            <span className="font-mono text-slate-700 dark:text-slate-300 text-xs whitespace-nowrap">
                              {(() => {
                                if (!val) return '-';
                                try {
                                  const str = String(val);
                                  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
                                    const [y, m, d] = str.split('-').map(Number);
                                    return new Date(y, m - 1, d).toLocaleDateString('en-US', {
                                      month: 'short',
                                      day: 'numeric',
                                      year: 'numeric',
                                    });
                                  }
                                  const d = new Date(val as string | number | Date);
                                  return isNaN(d.getTime())
                                    ? str
                                    : d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
                                } catch {
                                  return String(val);
                                }
                              })()}
                            </span>
                          ) : col.type === 'Currency' ? (
                            <span className="font-mono font-bold text-emerald-600 dark:text-emerald-400">
                              ${Number(val).toLocaleString(undefined, { minimumFractionDigits: 2 })}
                            </span>
                          ) : typeof val === 'boolean' ? (
                            val ? (
                              <CheckCircle className="w-4 h-4 text-emerald-500" />
                            ) : (
                              <XCircle className="w-4 h-4 text-slate-300" />
                            )
                          ) : (
                            String(val)
                          )}
                        </td>
                      );
                    })}

                    {/* Sticky Actions Column */}
                    <td
                      className={`sticky right-0 z-20 p-3 text-right whitespace-nowrap border-b border-l border-slate-200 dark:border-slate-800/80 shadow-[-2px_0_5px_-2px_rgba(0,0,0,0.08)] ${cellBg}`}
                    >
                      <div className="flex items-center justify-end gap-1">
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            onSelectItem(item);
                          }}
                          className="p-1 rounded text-slate-400 hover:text-brand-600 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors"
                          title="View Details"
                        >
                          <Eye className="w-3.5 h-3.5" />
                        </button>

                        {canUpdate && (
                          <button
                            onClick={(e) => {
                              e.stopPropagation();
                              onEditItem(item);
                            }}
                            className="p-1 rounded text-slate-400 hover:text-amber-600 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors"
                            title="Edit Record"
                          >
                            <Edit className="w-3.5 h-3.5" />
                          </button>
                        )}

                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            onDuplicateItem(item);
                          }}
                          className="p-1 rounded text-slate-400 hover:text-purple-600 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors"
                          title="Duplicate Record"
                        >
                          <Copy className="w-3.5 h-3.5" />
                        </button>

                        {canDelete && (
                          <button
                            onClick={(e) => {
                              e.stopPropagation();
                              onDeleteItem(item.id);
                            }}
                            className="p-1 rounded text-slate-400 hover:text-red-600 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors"
                            title="Delete Record"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        )}
                      </div>
                    </td>
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
        )}
      </div>

      {/* Pagination Footer */}
      <div className="p-3 border-t border-slate-200 dark:border-slate-800 bg-slate-50/50 dark:bg-slate-950/40 flex items-center justify-between text-xs text-slate-500">
        <div>
          Showing {filteredItems.length === 0 ? 0 : (currentPage - 1) * pageSize + 1} to{' '}
          {Math.min(currentPage * pageSize, filteredItems.length)} of {filteredItems.length} records
        </div>

        <div className="flex items-center gap-2">
          <button
            disabled={currentPage === 1}
            onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
            className="p-1.5 rounded-lg border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 disabled:opacity-40 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors"
          >
            <ChevronLeft className="w-4 h-4" />
          </button>
          <span className="font-semibold text-slate-700 dark:text-slate-300">
            Page {currentPage} of {totalPages}
          </span>
          <button
            disabled={currentPage >= totalPages}
            onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
            className="p-1.5 rounded-lg border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 disabled:opacity-40 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors"
          >
            <ChevronRight className="w-4 h-4" />
          </button>
        </div>
      </div>
      {/* Image Lightbox Preview Modal */}
      {previewImage && (
        <div
          className="fixed inset-0 z-50 bg-slate-900/80 backdrop-blur-xs flex items-center justify-center p-4 animate-fade-in"
          onClick={() => setPreviewImage(null)}
        >
          <div
            className="relative max-w-3xl max-h-[85vh] bg-white dark:bg-slate-900 rounded-2xl p-3 border border-slate-200 dark:border-slate-800 shadow-2xl overflow-hidden flex flex-col items-center"
            onClick={(e) => e.stopPropagation()}
          >
            <button
              onClick={() => setPreviewImage(null)}
              className="absolute top-3 right-3 p-1.5 rounded-full bg-slate-900/70 hover:bg-slate-900 text-white z-10 transition-colors"
              title="Close Preview"
            >
              <X className="w-5 h-5" />
            </button>
            <img
              src={previewImage}
              alt="Asset Preview"
              className="max-w-full max-h-[75vh] object-contain rounded-xl"
            />
            <div className="mt-2 text-xs text-slate-500 dark:text-slate-400 font-mono truncate max-w-full px-2">
              {previewImage}
            </div>
          </div>
        </div>
      )}

      {/* Bulk Edit Modal */}
      {isBulkEditModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-xs animate-fade-in">
          <div className="bg-white dark:bg-slate-900 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-enterprise w-full max-w-xl max-h-[85vh] flex flex-col animate-scale-up">
            {/* Header */}
            <div className="p-4 md:p-6 border-b border-slate-200 dark:border-slate-800 flex items-center justify-between bg-slate-50/50 dark:bg-slate-950/50">
              <div className="flex items-center gap-2">
                <div className="p-2 rounded-lg bg-amber-50 text-amber-600 dark:bg-amber-950 dark:text-amber-400">
                  <Edit className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-slate-900 dark:text-slate-100">
                    Bulk Edit ({selectedIds.size} Selected Records)
                  </h3>
                  <p className="text-xs text-slate-500">Select fields to update across all chosen records.</p>
                </div>
              </div>
              <button
                onClick={() => setIsBulkEditModalOpen(false)}
                className="p-1.5 rounded-lg text-slate-400 hover:text-slate-600 dark:hover:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Form list */}
            <div className="flex-1 overflow-y-auto p-4 md:p-6 space-y-4">
              {columns
                .filter((col) => !col.readOnly && col.name !== 'ID')
                .map((col) => {
                  const isEnabled = enabledBulkFields.has(col.name);
                  return (
                    <div
                      key={col.id}
                      className={`p-3 rounded-xl border transition-colors ${
                        isEnabled
                          ? 'border-amber-300 dark:border-amber-800 bg-amber-50/30 dark:bg-amber-950/20'
                          : 'border-slate-200 dark:border-slate-800 bg-slate-50/40 dark:bg-slate-950/30'
                      }`}
                    >
                      <div className="flex items-center justify-between mb-2">
                        <label className="flex items-center gap-2 cursor-pointer text-xs font-bold text-slate-800 dark:text-slate-200">
                          <input
                            type="checkbox"
                            checked={isEnabled}
                            onChange={(e) => {
                              const next = new Set(enabledBulkFields);
                              if (e.target.checked) {
                                next.add(col.name);
                              } else {
                                next.delete(col.name);
                              }
                              setEnabledBulkFields(next);
                            }}
                            className="w-4 h-4 rounded border-slate-300 text-amber-600 focus:ring-amber-500"
                          />
                          <span>{col.displayName}</span>
                        </label>
                        <span className="text-[10px] text-slate-400 font-mono uppercase">{col.type}</span>
                      </div>

                      {isEnabled && (
                        <div className="pl-6">
                          {col.type === 'Choice' ? (
                            <select
                              value={bulkFieldValues[col.name] || ''}
                              onChange={(e) =>
                                setBulkFieldValues({ ...bulkFieldValues, [col.name]: e.target.value })
                              }
                              className="w-full px-3 py-1.5 text-xs rounded-lg border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 text-slate-900 dark:text-slate-100 focus:ring-2 focus:ring-amber-500/50 outline-none"
                            >
                              <option value="">-- Select {col.displayName} --</option>
                              {col.choices?.map((c) => (
                                <option key={c} value={c}>
                                  {c}
                                </option>
                              ))}
                            </select>
                          ) : String(col.type).toLowerCase() === 'datetime' || String(col.type).toLowerCase() === 'date' ? (
                            <input
                              type="date"
                              value={bulkFieldValues[col.name] || ''}
                              onChange={(e) =>
                                setBulkFieldValues({ ...bulkFieldValues, [col.name]: e.target.value })
                              }
                              className="w-full px-3 py-1.5 text-xs rounded-lg border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 text-slate-900 dark:text-slate-100 focus:ring-2 focus:ring-amber-500/50 outline-none"
                            />
                          ) : col.type === 'Currency' || col.type === 'Number' ? (
                            <input
                              type="number"
                              value={bulkFieldValues[col.name] ?? ''}
                              onChange={(e) =>
                                setBulkFieldValues({
                                  ...bulkFieldValues,
                                  [col.name]: e.target.value === '' ? '' : Number(e.target.value),
                                })
                              }
                              placeholder={`Enter ${col.displayName}...`}
                              className="w-full px-3 py-1.5 text-xs rounded-lg border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 text-slate-900 dark:text-slate-100 focus:ring-2 focus:ring-amber-500/50 outline-none"
                            />
                          ) : col.type === 'Boolean' ? (
                            <select
                              value={bulkFieldValues[col.name] === undefined ? '' : String(bulkFieldValues[col.name])}
                              onChange={(e) =>
                                setBulkFieldValues({
                                  ...bulkFieldValues,
                                  [col.name]: e.target.value === 'true',
                                })
                              }
                              className="w-full px-3 py-1.5 text-xs rounded-lg border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 text-slate-900 dark:text-slate-100 focus:ring-2 focus:ring-amber-500/50 outline-none"
                            >
                              <option value="true">Yes / True</option>
                              <option value="false">No / False</option>
                            </select>
                          ) : (
                            <input
                              type="text"
                              value={bulkFieldValues[col.name] || ''}
                              onChange={(e) =>
                                setBulkFieldValues({ ...bulkFieldValues, [col.name]: e.target.value })
                              }
                              placeholder={`Enter new ${col.displayName}...`}
                              className="w-full px-3 py-1.5 text-xs rounded-lg border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 text-slate-900 dark:text-slate-100 focus:ring-2 focus:ring-amber-500/50 outline-none"
                            />
                          )}
                        </div>
                      )}
                    </div>
                  );
                })}
            </div>

            {/* Footer */}
            <div className="p-4 border-t border-slate-200 dark:border-slate-800 bg-slate-50/50 dark:bg-slate-950/50 flex items-center justify-between">
              <span className="text-xs text-slate-500">
                {enabledBulkFields.size} field(s) selected for update
              </span>
              <div className="flex items-center gap-2">
                <button
                  onClick={() => setIsBulkEditModalOpen(false)}
                  className="px-4 py-2 rounded-xl text-xs font-semibold text-slate-600 dark:text-slate-400 hover:bg-slate-200/60 dark:hover:bg-slate-800 transition-colors"
                >
                  Cancel
                </button>
                <button
                  disabled={enabledBulkFields.size === 0}
                  onClick={handleApplyBulkEdit}
                  className="px-4 py-2 rounded-xl text-xs font-bold text-white bg-amber-600 hover:bg-amber-700 disabled:opacity-50 shadow-sm transition-colors"
                >
                  Apply to {selectedIds.size} Records
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
