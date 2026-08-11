import React, { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { useAppStore } from '../../stores/useAppStore';
import { SharePointColumnDefinition, SharePointListItem } from '../../types';
import { graphService } from '../../services/graphService';
import { DynamicTable } from '../crud/DynamicTable';
import { DynamicForm } from '../crud/DynamicForm';
import { ItemDetailDrawer } from '../crud/ItemDetailDrawer';
import { IconResolver } from '../common/IconResolver';
import {
  ListFilter,
  Plus,
  RefreshCw,
  FolderKanban,
  CheckCircle2,
  AlertCircle,
  Database,
  Building2,
  X,
} from 'lucide-react';

interface WorkspaceViewProps {
  onOpenCreateMenu: () => void;
}

export const WorkspaceView: React.FC<WorkspaceViewProps> = ({ onOpenCreateMenu }) => {
  const navigate = useNavigate();
  const {
    activeWorkspaceId,
    activeMenuId,
    setActiveMenu,
    workspaces,
    menus,
    currentUser,
    logAction,
    addToast,
  } = useAppStore();

  const activeWs = workspaces.find((w) => w.id === activeWorkspaceId);
  const activeWsMenus = menus
    .filter((m) => m.workspaceId === activeWorkspaceId && m.visible)
    .sort((a, b) => a.displayOrder - b.displayOrder);

  const activeMenu = menus.find((m) => m.id === activeMenuId);

  // State for data and CRUD dialogs
  const [columns, setColumns] = useState<SharePointColumnDefinition[]>([]);
  const [items, setItems] = useState<SharePointListItem[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);

  // Modal / Drawer state
  const [isFormModalOpen, setIsFormModalOpen] = useState(false);
  const [editingItem, setEditingItem] = useState<SharePointListItem | null>(null);
  const [selectedItem, setSelectedItem] = useState<SharePointListItem | null>(null);
  const [isSaving, setIsSaving] = useState(false);

  // Load SharePoint Column Definitions and List Items
  const loadListData = async () => {
    if (!activeMenu) return;
    setIsLoading(true);
    setErrorMsg(null);

    try {
      const [colsData, itemsData] = await Promise.all([
        graphService.getListColumns(activeMenu.sharePointSiteId, activeMenu.sharePointListId),
        graphService.getListItems(activeMenu.sharePointSiteId, activeMenu.sharePointListId),
      ]);

      setColumns(colsData);
      setItems(itemsData);
    } catch (err: any) {
      console.error('Failed to load SharePoint List Data:', err);
      setErrorMsg(err.message || 'Error communicating with Microsoft Graph API.');
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    if (activeMenu) {
      loadListData();
    }
  }, [activeMenuId]);

  if (!activeWs) return null;

  // Role CRUD Permissions Check for Active Menu
  const userRole = currentUser.role;
  const menuPermissions = activeMenu?.permissions[userRole] || {
    create: true,
    read: true,
    update: true,
    delete: true,
    export: true,
  };

  // Form Submit Handler (Create or Update)
  const handleFormSubmit = async (formData: Record<string, any>) => {
    if (!activeMenu) return;
    setIsSaving(true);

    try {
      if (editingItem) {
        // Update
        const updated = await graphService.updateListItem(
          activeMenu.sharePointSiteId,
          activeMenu.sharePointListId,
          editingItem.id,
          formData
        );

        setItems((prev) => prev.map((i) => (i.id === editingItem.id ? updated : i)));
        logAction('UPDATE', `Updated item ID: ${editingItem.id}`, activeWs.name, activeMenu.name);
        addToast('success', 'Record Saved', 'SharePoint list item updated successfully.');
      } else {
        // Create
        const created = await graphService.createListItem(
          activeMenu.sharePointSiteId,
          activeMenu.sharePointListId,
          formData
        );

        setItems((prev) => [created, ...prev]);
        logAction('CREATE', `Created item "${formData.Title || 'New Record'}"`, activeWs.name, activeMenu.name);
        addToast('success', 'Record Created', 'New item added to SharePoint list.');
      }

      setIsFormModalOpen(false);
      setEditingItem(null);
    } catch (err: any) {
      addToast('error', 'Operation Failed', err.message || 'Could not save record to SharePoint.');
    } finally {
      setIsSaving(false);
    }
  };

  // Delete Handler
  const handleDeleteItem = async (itemId: string) => {
    if (!activeMenu) return;
    if (!window.confirm('Are you sure you want to delete this SharePoint record?')) return;

    try {
      await graphService.deleteListItem(activeMenu.sharePointSiteId, activeMenu.sharePointListId, itemId);
      setItems((prev) => prev.filter((i) => i.id !== itemId));
      setSelectedItem(null);
      logAction('DELETE', `Deleted item ID: ${itemId}`, activeWs.name, activeMenu.name);
      addToast('warning', 'Record Deleted', 'Item permanently removed from SharePoint.');
    } catch (err: any) {
      addToast('error', 'Delete Error', err.message || 'Failed to delete record.');
    }
  };

  // Bulk Delete Handler
  const handleBulkDeleteItems = async (itemIds: string[]) => {
    if (!activeMenu || itemIds.length === 0) return;
    const confirmMsg =
      itemIds.length === 1
        ? 'Are you sure you want to delete this SharePoint record?'
        : `Are you sure you want to delete ${itemIds.length} selected SharePoint records?`;
    if (!window.confirm(confirmMsg)) return;

    try {
      await Promise.all(
        itemIds.map((id) =>
          graphService.deleteListItem(activeMenu.sharePointSiteId, activeMenu.sharePointListId, id)
        )
      );
      setItems((prev) => prev.filter((i) => !itemIds.includes(i.id)));
      if (selectedItem && itemIds.includes(selectedItem.id)) {
        setSelectedItem(null);
      }
      logAction('DELETE', `Bulk deleted ${itemIds.length} items`, activeWs.name, activeMenu.name);
      addToast('warning', 'Records Deleted', `${itemIds.length} record(s) permanently removed from SharePoint.`);
    } catch (err: any) {
      addToast('error', 'Delete Error', err.message || 'Failed to delete selected records.');
    }
  };

  // Bulk Edit Handler
  const handleBulkEditItems = async (itemIds: string[], updateFields: Record<string, any>) => {
    if (!activeMenu || itemIds.length === 0 || Object.keys(updateFields).length === 0) return;
    setIsSaving(true);

    try {
      const updatedResults = await Promise.all(
        itemIds.map((id) =>
          graphService.updateListItem(
            activeMenu.sharePointSiteId,
            activeMenu.sharePointListId,
            id,
            updateFields
          )
        )
      );

      const updatedMap = new Map(updatedResults.map((item) => [item.id, item]));
      setItems((prev) => prev.map((i) => (updatedMap.has(i.id) ? updatedMap.get(i.id)! : i)));

      logAction('UPDATE', `Bulk updated ${itemIds.length} items`, activeWs.name, activeMenu.name);
      addToast('success', 'Records Updated', `Successfully updated ${itemIds.length} selected items.`);
    } catch (err: any) {
      addToast('error', 'Bulk Edit Error', err.message || 'Failed to update selected records.');
    } finally {
      setIsSaving(false);
    }
  };

  // Duplicate Item
  const handleDuplicateItem = async (item: SharePointListItem) => {
    if (!activeMenu) return;
    const duplicatedFields = { ...item.fields };
    delete duplicatedFields.id;
    if (duplicatedFields.Title) {
      duplicatedFields.Title = `${duplicatedFields.Title} (Copy)`;
    }

    try {
      const created = await graphService.createListItem(
        activeMenu.sharePointSiteId,
        activeMenu.sharePointListId,
        duplicatedFields
      );
      setItems((prev) => [created, ...prev]);
      addToast('success', 'Record Duplicated', 'Duplicated item created successfully.');
    } catch (err: any) {
      addToast('error', 'Duplicate Error', err.message);
    }
  };

  return (
    <div className="flex-1 overflow-y-auto p-6 md:p-8 space-y-6 bg-slate-50/50 dark:bg-slate-950/40">
      {/* Workspace Header */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 p-6 rounded-2xl bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 shadow-enterprise">
        <div className="flex items-center gap-4">
          <div
            className="w-12 h-12 rounded-xl flex items-center justify-center text-white shadow-md"
            style={{ backgroundColor: activeWs.color }}
          >
            <IconResolver name={activeWs.iconName} className="w-6 h-6" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <span className="text-xs font-semibold text-slate-400 uppercase tracking-wider">Workspace</span>
              <span className="w-1.5 h-1.5 rounded-full bg-slate-300" />
              <span className="text-xs font-mono text-slate-400">{activeWs.id}</span>
            </div>
            <h1 className="text-xl font-extrabold text-slate-900 dark:text-slate-100">{activeWs.name}</h1>
            <p className="text-xs text-slate-500 mt-0.5">{activeWs.description}</p>
          </div>
        </div>

        {currentUser.role === 'Administrator' && (
          <button
            onClick={onOpenCreateMenu}
            className="px-3.5 py-2 rounded-xl bg-brand-600 hover:bg-brand-700 text-white font-semibold text-xs flex items-center gap-1.5 shadow-sm transition-colors self-start md:self-auto"
          >
            <Plus className="w-4 h-4" /> Link SharePoint List
          </button>
        )}
      </div>

      {/* Menu Selector Horizontal Tabs */}
      <div className="flex items-center gap-2 overflow-x-auto pb-1 border-b border-slate-200 dark:border-slate-800">
        {activeWsMenus.map((m) => {
          const isSelected = m.id === activeMenuId;

          return (
            <button
              key={m.id}
              onClick={() => {
                setActiveMenu(m.id);
                navigate(`/lists/${m.id}`);
              }}
              className={`px-4 py-2.5 rounded-xl text-xs font-bold flex items-center gap-2 transition-all shrink-0 ${
                isSelected
                  ? 'bg-brand-600 text-white shadow-md'
                  : 'bg-white dark:bg-slate-900 text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-800 border border-slate-200/80 dark:border-slate-800'
              }`}
            >
              <IconResolver name={m.iconName} className="w-4 h-4" />
              <span>{m.name}</span>
            </button>
          );
        })}
      </div>

      {/* Main Dynamic Table Container */}
      {!activeMenu ? (
        <div className="p-12 text-center bg-white dark:bg-slate-900 rounded-2xl border border-slate-200 dark:border-slate-800">
          <FolderKanban className="w-12 h-12 mx-auto text-slate-300 dark:text-slate-600 mb-2" />
          <h3 className="text-sm font-bold text-slate-700 dark:text-slate-200">No List Selected</h3>
          <p className="text-xs text-slate-400 mt-1">Select a menu tab above or connect a new SharePoint list.</p>
        </div>
      ) : (
        <div className="space-y-4">
          {/* List Meta Bar */}
          <div className="flex items-center justify-between text-xs text-slate-500 px-1">
            <div className="flex items-center gap-2 font-mono">
              <Database className="w-3.5 h-3.5 text-brand-600" />
              <span>
                Site: <strong className="text-slate-800 dark:text-slate-200">{activeMenu.sharePointSiteId}</strong> | List:{' '}
                <strong className="text-slate-800 dark:text-slate-200">{activeMenu.sharePointListName}</strong>
              </span>
            </div>

            <button
              onClick={loadListData}
              disabled={isLoading}
              className="p-1.5 rounded-lg border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 text-slate-600 hover:text-brand-600 flex items-center gap-1 font-semibold text-[11px]"
              title="Refresh Graph Data"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${isLoading ? 'animate-spin' : ''}`} /> Refresh
            </button>
          </div>

          {/* Error Message banner */}
          {errorMsg && (
            <div className="p-4 rounded-xl bg-red-50 dark:bg-red-950/60 border border-red-200 dark:border-red-900 text-red-800 dark:text-red-200 text-xs flex items-center gap-3">
              <AlertCircle className="w-5 h-5 shrink-0 text-red-500" />
              <div className="flex-1">
                <span className="font-bold block">Graph API Error</span>
                <span>{errorMsg}</span>
              </div>
            </div>
          )}

          {/* Dynamic Table Component */}
          <DynamicTable
            menu={activeMenu}
            columns={columns}
            items={items}
            isLoading={isLoading}
            onAddItem={() => {
              setEditingItem(null);
              setIsFormModalOpen(true);
            }}
            onEditItem={(item) => {
              setEditingItem(item);
              setIsFormModalOpen(true);
            }}
            onDeleteItem={handleDeleteItem}
            onBulkDeleteItems={handleBulkDeleteItems}
            onBulkEditItems={handleBulkEditItems}
            onDuplicateItem={handleDuplicateItem}
            onSelectItem={(item) => setSelectedItem(item)}
            canCreate={menuPermissions.create}
            canUpdate={menuPermissions.update}
            canDelete={menuPermissions.delete}
            canExport={menuPermissions.export}
          />
        </div>
      )}

      {/* Form Modal for Creating/Editing Item */}
      {isFormModalOpen && activeMenu && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-xs animate-fade-in">
          <div className="w-full max-w-2xl bg-white dark:bg-slate-900 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-2xl overflow-hidden max-h-[90vh] flex flex-col">
            <div className="p-4 md:p-5 border-b border-slate-200 dark:border-slate-800 flex items-center justify-between bg-slate-50 dark:bg-slate-950/50">
              <h3 className="text-sm font-bold text-slate-900 dark:text-slate-100">
                {editingItem ? `Edit ${activeMenu.name} Record` : `New ${activeMenu.name} Record`}
              </h3>
              <button onClick={() => setIsFormModalOpen(false)} className="text-slate-400 hover:text-slate-600">
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-6 overflow-y-auto flex-1">
              <DynamicForm
                columns={columns}
                initialData={editingItem?.fields}
                onSubmit={handleFormSubmit}
                onCancel={() => setIsFormModalOpen(false)}
                isLoading={isSaving}
              />
            </div>
          </div>
        </div>
      )}

      {/* Slide-out Item Inspector Drawer */}
      <ItemDetailDrawer
        item={selectedItem}
        columns={columns}
        onClose={() => setSelectedItem(null)}
        onEdit={(item) => {
          setSelectedItem(null);
          setEditingItem(item);
          setIsFormModalOpen(true);
        }}
        onDelete={handleDeleteItem}
        onDuplicate={handleDuplicateItem}
      />
    </div>
  );
};
