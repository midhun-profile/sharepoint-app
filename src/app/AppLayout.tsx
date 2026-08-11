import React, { useState } from 'react';
import { Outlet } from 'react-router-dom';
import { AppShell } from '../components/layout/AppShell/AppShell';
import { AdminPanel } from '../components/admin/AdminPanel';
import { WorkspaceModal } from '../components/admin/WorkspaceModal';
import { MenuConfigModal } from '../components/admin/MenuConfigModal';
import { AuditLogDrawer } from '../components/admin/AuditLogDrawer';
import { ToastContainer } from '../components/common/ToastContainer';
import { useAppStore } from '../stores/useAppStore';

/**
 * Skeleton Loading Shell rendered during React.lazy page transitions
 */
const SkeletonShell: React.FC = () => (
  <div className="flex-1 p-6 md:p-8 space-y-6 bg-slate-50/50 dark:bg-slate-950/40 animate-pulse">
    <div className="h-24 bg-slate-200 dark:bg-slate-800/60 rounded-2xl w-full" />
    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
      <div className="h-20 bg-slate-200 dark:bg-slate-800/60 rounded-xl" />
      <div className="h-20 bg-slate-200 dark:bg-slate-800/60 rounded-xl" />
      <div className="h-20 bg-slate-200 dark:bg-slate-800/60 rounded-xl" />
      <div className="h-20 bg-slate-200 dark:bg-slate-800/60 rounded-xl" />
    </div>
    <div className="h-64 bg-slate-200 dark:bg-slate-800/60 rounded-2xl w-full" />
  </div>
);

/**
 * Main Enterprise Application Shell Layout
 */
export const AppLayout: React.FC = () => {
  const { workspaces, addWorkspace, addMenu, activeWorkspaceId } = useAppStore();

  const [globalSearch, setGlobalSearch] = useState('');
  const [isAdminOpen, setIsAdminOpen] = useState(false);
  const [isAuditOpen, setIsAuditOpen] = useState(false);
  const [isWorkspaceModalOpen, setIsWorkspaceModalOpen] = useState(false);
  const [isMenuModalOpen, setIsMenuModalOpen] = useState(false);

  return (
    <AppShell
      onOpenAdmin={() => setIsAdminOpen(true)}
      onOpenAudit={() => setIsAuditOpen(true)}
      onOpenCreateWorkspace={() => setIsWorkspaceModalOpen(true)}
      onOpenCreateMenu={() => setIsMenuModalOpen(true)}
      globalSearch={globalSearch}
      onGlobalSearchChange={setGlobalSearch}
    >
      <React.Suspense fallback={<SkeletonShell />}>
        <Outlet />
      </React.Suspense>

      {/* Admin Panel Modal */}
      <AdminPanel
        isOpen={isAdminOpen}
        onClose={() => setIsAdminOpen(false)}
        onOpenCreateWorkspace={() => setIsWorkspaceModalOpen(true)}
        onOpenCreateMenu={() => setIsMenuModalOpen(true)}
      />

      {/* Workspace Creation Modal */}
      <WorkspaceModal
        isOpen={isWorkspaceModalOpen}
        onClose={() => setIsWorkspaceModalOpen(false)}
        onSubmit={(ws) => addWorkspace(ws)}
      />

      {/* Menu Configuration Modal */}
      <MenuConfigModal
        isOpen={isMenuModalOpen}
        onClose={() => setIsMenuModalOpen(false)}
        onSubmit={(menu) => addMenu(menu)}
        workspaces={workspaces}
        activeWorkspaceId={activeWorkspaceId}
      />

      {/* System Audit Trail Drawer */}
      <AuditLogDrawer
        isOpen={isAuditOpen}
        onClose={() => setIsAuditOpen(false)}
      />

      {/* Accessible Toast Notifications */}
      <ToastContainer />
    </AppShell>
  );
};

