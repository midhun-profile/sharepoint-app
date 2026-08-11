import React, { useEffect, useState } from 'react';
import { useParams } from 'react-router-dom';
import { WorkspaceView } from '../../components/workspace/WorkspaceView';
import { MenuConfigModal } from '../../components/admin/MenuConfigModal';
import { useAppStore } from '../../stores/useAppStore';

export const WorkspacePage: React.FC = () => {
  const { workspaceId } = useParams<{ workspaceId: string }>();
  const { setActiveWorkspace, addMenu, workspaces, activeWorkspaceId } = useAppStore();
  const [isMenuModalOpen, setIsMenuModalOpen] = useState(false);

  useEffect(() => {
    if (workspaceId) {
      setActiveWorkspace(workspaceId);
    }
  }, [workspaceId, setActiveWorkspace]);

  return (
    <>
      <WorkspaceView onOpenCreateMenu={() => setIsMenuModalOpen(true)} />

      <MenuConfigModal
        isOpen={isMenuModalOpen}
        onClose={() => setIsMenuModalOpen(false)}
        onSubmit={(menu) => addMenu(menu)}
        workspaces={workspaces}
        activeWorkspaceId={activeWorkspaceId || workspaceId || null}
      />
    </>
  );
};

export default WorkspacePage;
