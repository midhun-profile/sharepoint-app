import React, { useEffect, useState } from 'react';
import { useParams } from 'react-router-dom';
import { useAppStore } from '../../stores/useAppStore';
import { WorkspaceView } from '../../components/workspace/WorkspaceView';
import { MenuConfigModal } from '../../components/admin/MenuConfigModal';

export const ListPage: React.FC = () => {
  const { listKey } = useParams<{ listKey: string }>();
  const { menus, setActiveMenu, setActiveWorkspace, addMenu, workspaces, activeWorkspaceId } = useAppStore();
  const [isMenuModalOpen, setIsMenuModalOpen] = useState(false);

  useEffect(() => {
    if (listKey) {
      // Find menu by id or sharePointListId
      const targetMenu = menus.find((m) => m.id === listKey || m.sharePointListId === listKey);
      if (targetMenu) {
        setActiveWorkspace(targetMenu.workspaceId);
        setActiveMenu(targetMenu.id);
      }
    }
  }, [listKey, menus, setActiveMenu, setActiveWorkspace]);

  return (
    <>
      <WorkspaceView onOpenCreateMenu={() => setIsMenuModalOpen(true)} />

      <MenuConfigModal
        isOpen={isMenuModalOpen}
        onClose={() => setIsMenuModalOpen(false)}
        onSubmit={(m) => addMenu(m)}
        workspaces={workspaces}
        activeWorkspaceId={activeWorkspaceId}
      />
    </>
  );
};

export default ListPage;
