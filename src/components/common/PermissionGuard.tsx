import React from 'react';
import { Outlet } from 'react-router-dom';
import { useAppStore } from '../../stores/useAppStore';
import { UserRole } from '../../types';
import { ShieldAlert, ArrowLeft } from 'lucide-react';

interface PermissionGuardProps {
  requiredRole?: UserRole;
  requiredPermission?: string; // e.g. "admin.manage" or "create"
  fallback?: React.ReactNode;
  children?: React.ReactNode;
}

/**
 * Role & Permission Authorization Gate Wrapper
 * Evaluates active user context against required roles or granular permission matrices.
 */
export const PermissionGuard: React.FC<PermissionGuardProps> = ({
  requiredRole,
  requiredPermission,
  fallback,
  children,
}) => {
  const { currentUser } = useAppStore();

  const userRole = currentUser?.role || 'Employee';

  // Role hierarchy check: Administrator > Manager > Employee
  let isAuthorized = true;

  if (requiredRole) {
    if (requiredRole === 'Administrator') {
      isAuthorized = userRole === 'Administrator';
    } else if (requiredRole === 'Manager') {
      isAuthorized = userRole === 'Administrator' || userRole === 'Manager';
    }
  }

  // Permission key check (if granular permissions are specified)
  if (isAuthorized && requiredPermission) {
    if (requiredPermission === 'admin.manage') {
      isAuthorized = userRole === 'Administrator';
    }
  }

  if (!isAuthorized) {
    if (fallback) {
      return <>{fallback}</>;
    }

    return (
      <div className="flex-1 flex flex-col items-center justify-center p-8 text-center bg-slate-50/50 dark:bg-slate-950/40 min-h-[400px]">
        <div className="max-w-md w-full bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-2xl p-8 shadow-enterprise space-y-4">
          <div className="w-14 h-14 rounded-2xl bg-red-50 dark:bg-red-950/60 border border-red-200 dark:border-red-900 flex items-center justify-center mx-auto text-red-600 dark:text-red-400 shadow-xs">
            <ShieldAlert className="w-7 h-7" />
          </div>

          <div className="space-y-1">
            <h3 className="text-lg font-bold text-slate-900 dark:text-slate-100">
              Access Restricted
            </h3>
            <p className="text-xs text-slate-500 leading-relaxed">
              Your assigned role (<strong className="text-slate-700 dark:text-slate-300">{userRole}</strong>) does not have authorization to view or manage this route.
            </p>
          </div>

          <div className="pt-2">
            <button
              onClick={() => window.history.back()}
              className="inline-flex items-center gap-2 px-4 py-2 rounded-xl bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-200 text-xs font-semibold transition-colors"
            >
              <ArrowLeft className="w-4 h-4" /> Go Back
            </button>
          </div>
        </div>
      </div>
    );
  }

  return children ? <>{children}</> : <Outlet />;
};
