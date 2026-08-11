import React, { useState } from 'react';
import { AdminPanel } from '../../components/admin/AdminPanel';
import { WorkspaceModal } from '../../components/admin/WorkspaceModal';
import { MenuConfigModal } from '../../components/admin/MenuConfigModal';
import { useAppStore } from '../../stores/useAppStore';

export const AdminPage: React.FC = () => {
  const { addWorkspace, addMenu, workspaces, activeWorkspaceId } = useAppStore();

  const [isWorkspaceModalOpen, setIsWorkspaceModalOpen] = useState(false);
  const [isMenuModalOpen, setIsMenuModalOpen] = useState(false);

  return (
    <div className="flex-1 overflow-y-auto bg-slate-50/50 dark:bg-slate-950/40 p-6 md:p-8">
      <AdminPanel
        isOpen={true}
        onClose={() => window.history.back()}
        onOpenCreateWorkspace={() => setIsWorkspaceModalOpen(true)}
        onOpenCreateMenu={() => setIsMenuModalOpen(true)}
      />

      <WorkspaceModal
        isOpen={isWorkspaceModalOpen}
        onClose={() => setIsWorkspaceModalOpen(false)}
        onSubmit={(ws) => addWorkspace(ws)}
      />

      <MenuConfigModal
        isOpen={isMenuModalOpen}
        onClose={() => setIsMenuModalOpen(false)}
        onSubmit={(menu) => addMenu(menu)}
        workspaces={workspaces}
        activeWorkspaceId={activeWorkspaceId}
      />
    </div>
  );
};

export default AdminPage;
