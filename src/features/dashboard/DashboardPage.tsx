import React, { useState } from 'react';
import { Dashboard } from '../../components/dashboard/Dashboard';
import { WorkspaceModal } from '../../components/admin/WorkspaceModal';
import { MenuConfigModal } from '../../components/admin/MenuConfigModal';
import { useAppStore } from '../../stores/useAppStore';

export const DashboardPage: React.FC = () => {
  const { addWorkspace, addMenu, workspaces, activeWorkspaceId } = useAppStore();

  const [isWorkspaceModalOpen, setIsWorkspaceModalOpen] = useState(false);
  const [isMenuModalOpen, setIsMenuModalOpen] = useState(false);

  return (
    <>
      <Dashboard
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
    </>
  );
};

export default DashboardPage;
