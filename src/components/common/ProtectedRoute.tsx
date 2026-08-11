import React from 'react';
import { Navigate, Outlet } from 'react-router-dom';
import { useIsAuthenticated, useMsal } from '@azure/msal-react';
import { InteractionStatus } from '@azure/msal-browser';
import { useAppStore } from '../../stores/useAppStore';
import { PageSkeleton } from '../ui/SkeletonLoader';

export interface ProtectedRouteProps {
  children?: React.ReactNode;
  redirectTo?: string;
}

/**
 * Enterprise Protected Route Authentication Gate
 * Wraps protected application routes to ensure they are strictly inaccessible to unauthenticated users.
 * Utilizes @azure/msal-react to verify Microsoft Entra ID claims in Live mode,
 * or verifies active user session profile in Demo mode.
 *
 * Renders a full PageSkeleton screen while MSAL is silently determining the login state,
 * preventing flash of unauthenticated content (FOUC).
 */
export const ProtectedRoute: React.FC<ProtectedRouteProps> = ({
  children,
  redirectTo = '/login',
}) => {
  const isAuthenticated = useIsAuthenticated();
  const { inProgress } = useMsal();
  const { isLiveEntraMode, currentUser } = useAppStore();

  // If MSAL is currently handling redirect/login/token interaction
  const isAuthInProgress =
    isLiveEntraMode && inProgress !== InteractionStatus.None;

  if (isAuthInProgress) {
    // Render full structural skeleton screen to maintain layout stability during auth resolution
    return <PageSkeleton />;
  }

  // Check authentication status: either Entra ID MSAL session or active demo mode user profile
  const isAllowed = isLiveEntraMode ? isAuthenticated : Boolean(currentUser);

  if (!isAllowed) {
    return <Navigate to={redirectTo} replace />;
  }

  return children ? <>{children}</> : <Outlet />;
};
