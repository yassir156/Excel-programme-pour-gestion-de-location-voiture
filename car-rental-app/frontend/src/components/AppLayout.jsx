import { Outlet, useLocation } from 'react-router-dom';
import Sidebar from './Sidebar';
import Header from './Header';

const titleRules = [
  { test: (p) => p === '/', title: 'Tableau de bord' },
  { test: (p) => p.startsWith('/vehicles/new'), title: 'Ajouter un véhicule' },
  { test: (p) => /\/vehicles\/\d+\/edit/.test(p), title: 'Modifier le véhicule' },
  { test: (p) => /\/vehicles\/\d+/.test(p), title: 'Détails du véhicule' },
  { test: (p) => p.startsWith('/vehicles'), title: 'Véhicules' },
  { test: (p) => p.startsWith('/clients/new'), title: 'Ajouter un client' },
  { test: (p) => /\/clients\/\d+\/edit/.test(p), title: 'Modifier le client' },
  { test: (p) => /\/clients\/\d+/.test(p), title: 'Détails du client' },
  { test: (p) => p.startsWith('/clients'), title: 'Clients' },
  { test: (p) => p.startsWith('/reservations/new'), title: 'Nouvelle réservation' },
  { test: (p) => /\/reservations\/\d+\/edit/.test(p), title: 'Modifier la réservation' },
  { test: (p) => p.startsWith('/reservations'), title: 'Réservations' },
  { test: (p) => p.startsWith('/contracts'), title: 'Contrats' },
  { test: (p) => p.startsWith('/payments'), title: 'Paiements' },
  { test: (p) => p.startsWith('/returns'), title: 'Retours' },
  { test: (p) => p.startsWith('/maintenance'), title: 'Maintenance' },
  { test: (p) => p.startsWith('/statistics'), title: 'Statistiques & Rapports' },
  { test: (p) => p.startsWith('/settings'), title: 'Paramètres' },
  { test: (p) => p.startsWith('/users'), title: 'Gestion des utilisateurs' },
];

function getTitle(pathname) {
  const rule = titleRules.find((r) => r.test(pathname));
  return rule ? rule.title : 'Nova Motion Car';
}

export default function AppLayout() {
  const location = useLocation();

  return (
    <div className="flex h-screen overflow-hidden bg-slate-50 dark:bg-navy-950">
      <Sidebar />
      <div className="flex min-w-0 flex-1 flex-col">
        <Header title={getTitle(location.pathname)} />
        <main className="flex-1 overflow-y-auto p-6">
          <Outlet />
        </main>
      </div>
    </div>
  );
}
