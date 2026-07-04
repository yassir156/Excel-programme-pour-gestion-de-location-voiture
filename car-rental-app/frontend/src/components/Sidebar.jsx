import { NavLink } from 'react-router-dom';
import {
  LayoutDashboard,
  Car,
  Users2,
  CalendarRange,
  FileText,
  Wallet,
  Undo2,
  Wrench,
  BarChart3,
  Settings as SettingsIcon,
  ShieldCheck,
} from 'lucide-react';
import { useAuth } from '../context/AuthContext';

const links = [
  { to: '/', label: 'Tableau de bord', icon: LayoutDashboard, end: true },
  { to: '/vehicles', label: 'Véhicules', icon: Car },
  { to: '/clients', label: 'Clients', icon: Users2 },
  { to: '/reservations', label: 'Réservations', icon: CalendarRange },
  { to: '/contracts', label: 'Contrats', icon: FileText },
  { to: '/payments', label: 'Paiements', icon: Wallet },
  { to: '/returns', label: 'Retours', icon: Undo2 },
  { to: '/maintenance', label: 'Maintenance', icon: Wrench },
  { to: '/statistics', label: 'Statistiques', icon: BarChart3 },
  { to: '/settings', label: 'Paramètres', icon: SettingsIcon },
];

export default function Sidebar() {
  const { user } = useAuth();

  return (
    <aside className="hidden w-64 shrink-0 flex-col border-r border-slate-200 bg-white dark:border-navy-700 dark:bg-navy-900 md:flex">
      <div className="flex items-center gap-2 px-5 py-5">
        <div className="flex h-9 w-9 items-center justify-center rounded-xl bg-brand-600 text-white shadow-soft">
          <Car size={18} />
        </div>
        <div>
          <p className="text-sm font-bold leading-none text-slate-800 dark:text-white">AutoLoc</p>
          <p className="text-xs text-slate-400">Gestion de location</p>
        </div>
      </div>

      <nav className="flex-1 space-y-1 overflow-y-auto px-3 py-2">
        {links.map(({ to, label, icon: Icon, end }) => (
          <NavLink
            key={to}
            to={to}
            end={end}
            className={({ isActive }) =>
              `flex items-center gap-3 rounded-xl px-3 py-2.5 text-sm font-medium transition-colors ${
                isActive
                  ? 'bg-brand-50 text-brand-700 dark:bg-navy-700 dark:text-white'
                  : 'text-slate-600 hover:bg-slate-50 dark:text-slate-300 dark:hover:bg-navy-800'
              }`
            }
          >
            <Icon size={18} />
            {label}
          </NavLink>
        ))}

        {user?.role === 'administrateur' && (
          <NavLink
            to="/users"
            className={({ isActive }) =>
              `flex items-center gap-3 rounded-xl px-3 py-2.5 text-sm font-medium transition-colors ${
                isActive
                  ? 'bg-brand-50 text-brand-700 dark:bg-navy-700 dark:text-white'
                  : 'text-slate-600 hover:bg-slate-50 dark:text-slate-300 dark:hover:bg-navy-800'
              }`
            }
          >
            <ShieldCheck size={18} />
            Utilisateurs
          </NavLink>
        )}
      </nav>

      <div className="border-t border-slate-100 px-5 py-4 text-xs text-slate-400 dark:border-navy-700">
        Car Rental Manager v1.0
      </div>
    </aside>
  );
}
