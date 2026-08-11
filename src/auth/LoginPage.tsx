import React from 'react';
import { useNavigate } from 'react-router-dom';
import { useMsal } from '@azure/msal-react';
import { loginRequest } from '../config/msalConfig';
import { useAppStore } from '../stores/useAppStore';
import { Building2, Shield, Lock, ArrowRight, UserCheck } from 'lucide-react';

export const LoginPage: React.FC = () => {
  const navigate = useNavigate();
  const { instance } = useMsal();
  const { isLiveEntraMode, setLiveEntraMode, setUserRole, currentUser } = useAppStore();

  const handleMsalLogin = async () => {
    try {
      await instance.loginPopup(loginRequest);
      navigate('/dashboard');
    } catch (err) {
      console.error('MSAL Login Failed:', err);
    }
  };

  const handleDemoLogin = (role: 'Administrator' | 'Manager' | 'Employee') => {
    setUserRole(role);
    setLiveEntraMode(false);
    navigate('/dashboard');
  };

  return (
    <div className="min-h-screen bg-slate-900 flex items-center justify-center p-4 relative overflow-hidden">
      {/* Background Decorative Gradients */}
      <div className="absolute top-0 left-1/4 w-96 h-96 bg-brand-600/10 rounded-full blur-3xl pointer-events-none" />
      <div className="absolute bottom-0 right-1/4 w-96 h-96 bg-blue-600/10 rounded-full blur-3xl pointer-events-none" />

      <div className="max-w-md w-full bg-slate-950 border border-slate-800 rounded-3xl p-8 shadow-2xl space-y-8 relative z-10">
        {/* Header Branding */}
        <div className="text-center space-y-3">
          <div className="w-14 h-14 rounded-2xl bg-brand-600/20 border border-brand-500/30 flex items-center justify-center mx-auto text-brand-400 shadow-md">
            <Building2 className="w-7 h-7" />
          </div>
          <div>
            <h1 className="text-xl font-extrabold text-white tracking-tight">
              SharePoint Management Portal
            </h1>
            <p className="text-xs text-slate-400 mt-1 leading-relaxed">
              Enterprise Control Plane for Corporate SharePoint Lists & Microsoft Graph Services
            </p>
          </div>
        </div>

        {/* Auth Mode Selection */}
        <div className="space-y-4">
          <div className="p-4 rounded-2xl bg-slate-900/80 border border-slate-800 space-y-3">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2 text-xs font-bold text-white">
                <Shield className="w-4 h-4 text-emerald-400" />
                <span>Microsoft Entra ID (Single Sign-On)</span>
              </div>
              <span className="px-2 py-0.5 rounded-full bg-emerald-500/20 text-emerald-400 text-[10px] font-bold border border-emerald-500/30">
                OAuth 2.0 PKCE
              </span>
            </div>

            <button
              onClick={handleMsalLogin}
              className="w-full py-2.5 px-4 rounded-xl bg-brand-600 hover:bg-brand-500 text-white font-bold text-xs shadow-lg flex items-center justify-center gap-2 transition-all cursor-pointer"
            >
              <Lock className="w-4 h-4" /> Sign In with Microsoft Account
            </button>
          </div>

          <div className="relative flex py-2 items-center">
            <div className="flex-grow border-t border-slate-800" />
            <span className="flex-shrink mx-3 text-[11px] text-slate-500 font-medium uppercase">
              Or Interactive Demo Session
            </span>
            <div className="flex-grow border-t border-slate-800" />
          </div>

          <div className="space-y-2">
            <p className="text-[11px] text-slate-400 font-medium">Select a demo persona to explore RBAC features:</p>
            <div className="grid grid-cols-3 gap-2">
              <button
                onClick={() => handleDemoLogin('Administrator')}
                className="p-2.5 rounded-xl bg-slate-900 hover:bg-slate-800 border border-slate-800 hover:border-brand-500/50 text-white text-[11px] font-bold text-center transition-all cursor-pointer"
              >
                <UserCheck className="w-4 h-4 mx-auto mb-1 text-purple-400" />
                Admin
              </button>
              <button
                onClick={() => handleDemoLogin('Manager')}
                className="p-2.5 rounded-xl bg-slate-900 hover:bg-slate-800 border border-slate-800 hover:border-brand-500/50 text-white text-[11px] font-bold text-center transition-all cursor-pointer"
              >
                <UserCheck className="w-4 h-4 mx-auto mb-1 text-blue-400" />
                Manager
              </button>
              <button
                onClick={() => handleDemoLogin('Employee')}
                className="p-2.5 rounded-xl bg-slate-900 hover:bg-slate-800 border border-slate-800 hover:border-brand-500/50 text-white text-[11px] font-bold text-center transition-all cursor-pointer"
              >
                <UserCheck className="w-4 h-4 mx-auto mb-1 text-emerald-400" />
                Employee
              </button>
            </div>
          </div>
        </div>

        {/* Footer info */}
        <div className="pt-2 text-center border-t border-slate-900">
          <p className="text-[10px] text-slate-500 font-mono">
            Security Mode: {isLiveEntraMode ? 'Live Microsoft Graph API' : 'Zero-Cost Demo Client'}
          </p>
        </div>
      </div>
    </div>
  );
};

export default LoginPage;
