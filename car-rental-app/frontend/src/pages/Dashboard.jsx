import { useEffect, useState } from 'react';
import { Link } from 'react-router-dom';
import {
  Car, CheckCircle2, Clock, Wallet, AlertTriangle, CalendarClock, Undo2,
} from 'lucide-react';
import {
  ResponsiveContainer, AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip, PieChart, Pie, Cell,
} from 'recharts';
import api from '../api/client';
import StatCard from '../components/StatCard';
import StatusBadge from '../components/StatusBadge';
import { format } from 'date-fns';

const PIE_COLORS = ['#10b981', '#2f5fed', '#f59e0b', '#a855f7'];

export default function Dashboard() {
  const [data, setData] = useState(null);
  const [settings, setSettings] = useState(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    Promise.all([api.get('/stats/dashboard'), api.get('/settings')])
      .then(([statsRes, settingsRes]) => {
        setData(statsRes.data);
        setSettings(settingsRes.data);
      })
      .finally(() => setLoading(false));
  }, []);

  if (loading || !data) {
    return <div className="text-slate-400">Chargement du tableau de bord...</div>;
  }

  const currency = settings?.currency || 'MAD';
  const revenueData = data.revenueByMonth.map((r) => ({ month: r.month, total: Number(r.total) }));
  const fleetData = [
    { name: 'Disponible', value: data.available },
    { name: 'Louée', value: data.rented },
    { name: 'Maintenance', value: data.maintenance },
    { name: 'Réservée', value: data.reserved },
  ].filter((d) => d.value > 0);

  const alertItems = [
    ...data.alerts.insuranceExpiringSoon.map((v) => ({
      type: 'Assurance', label: `${v.brand} ${v.model} (${v.plate})`, date: v.insuranceExpiry,
    })),
    ...data.alerts.technicalControlExpiringSoon.map((v) => ({
      type: 'Contrôle technique', label: `${v.brand} ${v.model} (${v.plate})`, date: v.technicalControlExpiry,
    })),
  ];

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 xl:grid-cols-4">
        <StatCard label="Véhicules au total" value={data.totalVehicles} icon={Car} tone="brand" />
        <StatCard label="Disponibles" value={data.available} icon={CheckCircle2} tone="accent" />
        <StatCard label="Réservations en cours" value={data.ongoingReservations} icon={Clock} tone="amber" />
        <StatCard
          label="Revenus du mois"
          value={`${data.monthlyRevenue.toLocaleString()} ${currency}`}
          icon={Wallet}
          tone="brand"
        />
      </div>

      <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 xl:grid-cols-4">
        <StatCard label="Véhicules loués" value={data.rented} icon={Car} tone="slate" />
        <StatCard label="En maintenance" value={data.maintenance} icon={AlertTriangle} tone="amber" />
        <StatCard
          label="Paiements en attente"
          value={`${data.pendingPayments.toLocaleString()} ${currency}`}
          icon={Wallet}
          tone="red"
        />
        <StatCard label="Alertes actives" value={alertItems.length} icon={AlertTriangle} tone="red" />
      </div>

      <div className="grid grid-cols-1 gap-4 xl:grid-cols-3">
        <div className="card p-5 xl:col-span-2">
          <h3 className="mb-4 text-sm font-semibold text-slate-700 dark:text-slate-200">Revenus (6 derniers mois)</h3>
          <ResponsiveContainer width="100%" height={260}>
            <AreaChart data={revenueData}>
              <defs>
                <linearGradient id="rev" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor="#2f5fed" stopOpacity={0.35} />
                  <stop offset="95%" stopColor="#2f5fed" stopOpacity={0} />
                </linearGradient>
              </defs>
              <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
              <XAxis dataKey="month" stroke="#94a3b8" fontSize={12} />
              <YAxis stroke="#94a3b8" fontSize={12} />
              <Tooltip />
              <Area type="monotone" dataKey="total" stroke="#2f5fed" fill="url(#rev)" strokeWidth={2} />
            </AreaChart>
          </ResponsiveContainer>
        </div>

        <div className="card p-5">
          <h3 className="mb-4 text-sm font-semibold text-slate-700 dark:text-slate-200">Répartition de la flotte</h3>
          <ResponsiveContainer width="100%" height={260}>
            <PieChart>
              <Pie data={fleetData} dataKey="value" nameKey="name" innerRadius={55} outerRadius={85} paddingAngle={3}>
                {fleetData.map((entry, index) => (
                  <Cell key={entry.name} fill={PIE_COLORS[index % PIE_COLORS.length]} />
                ))}
              </Pie>
              <Tooltip />
            </PieChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div className="grid grid-cols-1 gap-4 xl:grid-cols-3">
        <div className="card p-5 xl:col-span-2">
          <div className="mb-3 flex items-center gap-2">
            <CalendarClock size={18} className="text-brand-600" />
            <h3 className="text-sm font-semibold text-slate-700 dark:text-slate-200">Prochaines réservations</h3>
          </div>
          <div className="divide-y divide-slate-100 dark:divide-navy-700">
            {data.upcomingReservations.length === 0 && (
              <p className="py-4 text-sm text-slate-400">Aucune réservation à venir.</p>
            )}
            {data.upcomingReservations.map((r) => (
              <div key={r.id} className="flex items-center justify-between py-3 text-sm">
                <div>
                  <p className="font-medium text-slate-700 dark:text-slate-200">
                    {r.client?.firstName} {r.client?.lastName} — {r.vehicle?.brand} {r.vehicle?.model}
                  </p>
                  <p className="text-xs text-slate-400">
                    {format(new Date(r.startDate), 'dd/MM/yyyy')} → {format(new Date(r.endDate), 'dd/MM/yyyy')}
                  </p>
                </div>
                <StatusBadge status={r.status} />
              </div>
            ))}
          </div>
          <Link to="/reservations" className="mt-2 inline-block text-sm font-medium text-brand-600 hover:underline">
            Voir toutes les réservations →
          </Link>
        </div>

        <div className="card p-5">
          <div className="mb-3 flex items-center gap-2">
            <Undo2 size={18} className="text-brand-600" />
            <h3 className="text-sm font-semibold text-slate-700 dark:text-slate-200">Retours prévus</h3>
          </div>
          <div className="divide-y divide-slate-100 dark:divide-navy-700">
            {data.upcomingReturns.length === 0 && (
              <p className="py-4 text-sm text-slate-400">Aucun retour prévu.</p>
            )}
            {data.upcomingReturns.map((r) => (
              <div key={r.id} className="py-3 text-sm">
                <p className="font-medium text-slate-700 dark:text-slate-200">{r.vehicle?.brand} {r.vehicle?.model}</p>
                <p className="text-xs text-slate-400">
                  Retour prévu le {format(new Date(r.endDate), 'dd/MM/yyyy')} — {r.client?.firstName} {r.client?.lastName}
                </p>
              </div>
            ))}
          </div>
        </div>
      </div>

      {alertItems.length > 0 && (
        <div className="card border-amber-200 p-5">
          <div className="mb-3 flex items-center gap-2 text-amber-600">
            <AlertTriangle size={18} />
            <h3 className="text-sm font-semibold">Alertes importantes</h3>
          </div>
          <div className="grid grid-cols-1 gap-2 sm:grid-cols-2">
            {alertItems.map((a, i) => (
              <div key={i} className="flex items-center justify-between rounded-lg bg-amber-50 px-3 py-2 text-sm dark:bg-amber-950/40">
                <span className="text-amber-800 dark:text-amber-200">{a.type} — {a.label}</span>
                <span className="text-xs text-amber-600 dark:text-amber-400">{format(new Date(a.date), 'dd/MM/yyyy')}</span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}
