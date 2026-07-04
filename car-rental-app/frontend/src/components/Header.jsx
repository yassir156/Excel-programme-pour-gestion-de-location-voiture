import { useNavigate } from 'react-router-dom';
import { LogOut, Moon, Sun, UserCircle } from 'lucide-react';
import { useAuth } from '../context/AuthContext';
import { useTheme } from '../context/ThemeContext';

const roleLabels = {
  administrateur: 'Administrateur',
  manager: 'Manager',
  agent: 'Agent',
};

export default function Header({ title }) {
  const { user, logout } = useAuth();
  const { theme, toggleTheme } = useTheme();
  const navigate = useNavigate();

  function handleLogout() {
    logout();
    navigate('/login');
  }

  return (
    <header className="flex h-16 shrink-0 items-center justify-between border-b border-slate-200 bg-white px-6 dark:border-navy-700 dark:bg-navy-900">
      <h1 className="text-lg font-semibold text-slate-800 dark:text-white">{title}</h1>

      <div className="flex items-center gap-4">
        <button
          onClick={toggleTheme}
          className="rounded-full p-2 text-slate-500 hover:bg-slate-100 dark:text-slate-300 dark:hover:bg-navy-800"
          title="Basculer le thème"
        >
          {theme === 'dark' ? <Sun size={18} /> : <Moon size={18} />}
        </button>

        <div className="h-8 w-px bg-slate-200 dark:bg-navy-700" />

        <div className="flex items-center gap-2">
          <div className="flex h-9 w-9 items-center justify-center rounded-full bg-brand-100 text-brand-700 dark:bg-navy-700 dark:text-white">
            <UserCircle size={20} />
          </div>
          <div className="hidden text-sm sm:block">
            <p className="font-medium leading-none text-slate-800 dark:text-white">{user?.fullName}</p>
            <p className="text-xs text-slate-400">{roleLabels[user?.role] || user?.role}</p>
          </div>
        </div>

        <button
          onClick={handleLogout}
          className="flex items-center gap-1.5 rounded-lg px-3 py-2 text-sm font-medium text-slate-500 hover:bg-red-50 hover:text-red-600 dark:text-slate-300 dark:hover:bg-red-950"
          title="Déconnexion"
        >
          <LogOut size={16} />
          <span className="hidden sm:inline">Déconnexion</span>
        </button>
      </div>
    </header>
  );
}
