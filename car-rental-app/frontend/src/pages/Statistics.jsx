import { useEffect, useState } from 'react';
import { FileSpreadsheet, FileDown, Wallet, Percent, XCircle, TrendingUp } from 'lucide-react';
import {
  ResponsiveContainer, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip,
} from 'recharts';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import api, { getErrorMessage } from '../api/client';
import PageHeader from '../components/PageHeader';
import StatCard from '../components/StatCard';
import { exportToExcel, exportToCSV } from '../utils/exportData';

export default function Statistics() {
  const [report, setReport] = useState(null);
  const [settings, setSettings] = useState(null);
  const [error, setError] = useState('');

  useEffect(() => {
    Promise.all([api.get('/stats/reports'), api.get('/settings')])
      .then(([r, s]) => {
        setReport(r.data);
        setSettings(s.data);
      })
      .catch((err) => setError(getErrorMessage(err)));
  }, []);

  if (error) return <div className="rounded-lg bg-red-50 px-4 py-2 text-sm text-red-600">{error}</div>;
  if (!report) return <div className="text-slate-400">Chargement des statistiques...</div>;

  const currency = settings?.currency || 'MAD';
  const revenueData = report.revenueByMonth.map((r) => ({ month: r.month, total: Number(r.total) }));
  const vehicleData = report.mostRentedVehicles.map((r) => ({
    name: `${r.vehicle?.brand} ${r.vehicle?.model}`,
    count: Number(r.count),
  }));

  function handleExportExcel() {
    exportToExcel('statistiques-car-rental', {
      Revenus: report.revenueByMonth.map((r) => ({ Mois: r.month, Total: r.total })),
      'Vehicules les plus loues': report.mostRentedVehicles.map((r) => ({
        Vehicule: `${r.vehicle?.brand} ${r.vehicle?.model}`, Immatriculation: r.vehicle?.plate, 'Nombre de locations': r.count,
      })),
      'Clients les plus actifs': report.mostActiveClients.map((r) => ({
        Client: `${r.client?.firstName} ${r.client?.lastName}`, Telephone: r.client?.phone, 'Nombre de locations': r.count,
      })),
    });
  }

  function handleExportCSV() {
    exportToCSV('revenus-car-rental', report.revenueByMonth.map((r) => ({ Mois: r.month, Total: r.total })));
  }

  function handleExportPDF() {
    const doc = new jsPDF();
    doc.setFontSize(14);
    doc.text('Rapport statistique — ' + (settings?.name || 'Agence'), 14, 16);
    autoTable(doc, {
      startY: 24,
      head: [['Mois', `Revenus (${currency})`]],
      body: report.revenueByMonth.map((r) => [r.month, r.total]),
    });
    autoTable(doc, {
      startY: doc.lastAutoTable.finalY + 8,
      head: [['Véhicule', 'Immatriculation', 'Locations']],
      body: report.mostRentedVehicles.map((r) => [`${r.vehicle?.brand} ${r.vehicle?.model}`, r.vehicle?.plate, r.count]),
    });
    autoTable(doc, {
      startY: doc.lastAutoTable.finalY + 8,
      head: [['Client', 'Téléphone', 'Locations']],
      body: report.mostActiveClients.map((r) => [`${r.client?.firstName} ${r.client?.lastName}`, r.client?.phone, r.count]),
    });
    doc.save('rapport-statistiques.pdf');
  }

  return (
    <div className="space-y-6">
      <PageHeader
        title="Statistiques & Rapports"
        actions={
          <>
            <button className="btn-secondary" onClick={handleExportCSV}>
              <FileDown size={16} /> CSV
            </button>
            <button className="btn-secondary" onClick={handleExportExcel}>
              <FileSpreadsheet size={16} /> Excel
            </button>
            <button className="btn-primary" onClick={handleExportPDF}>
              <FileDown size={16} /> PDF
            </button>
          </>
        }
      />

      <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 xl:grid-cols-4">
        <StatCard label="Taux d'occupation" value={`${report.occupancyRate}%`} icon={Percent} tone="brand" />
        <StatCard label="Réservations annulées" value={report.cancelledReservations} icon={XCircle} tone="red" />
        <StatCard label="Paiements en attente" value={`${report.pendingPayments.toLocaleString()} ${currency}`} icon={Wallet} tone="amber" />
        <StatCard label="Revenu total (12 mois)" value={`${revenueData.reduce((a, b) => a + b.total, 0).toLocaleString()} ${currency}`} icon={TrendingUp} tone="accent" />
      </div>

      <div className="card p-5">
        <h3 className="mb-4 text-sm font-semibold text-slate-700 dark:text-slate-200">Revenus mensuels (12 derniers mois)</h3>
        <ResponsiveContainer width="100%" height={280}>
          <BarChart data={revenueData}>
            <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
            <XAxis dataKey="month" stroke="#94a3b8" fontSize={12} />
            <YAxis stroke="#94a3b8" fontSize={12} />
            <Tooltip />
            <Bar dataKey="total" fill="#2f5fed" radius={[6, 6, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
        <div className="card p-5">
          <h3 className="mb-4 text-sm font-semibold text-slate-700 dark:text-slate-200">Véhicules les plus loués</h3>
          <ResponsiveContainer width="100%" height={260}>
            <BarChart data={vehicleData} layout="vertical" margin={{ left: 20 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
              <XAxis type="number" stroke="#94a3b8" fontSize={12} />
              <YAxis type="category" dataKey="name" width={120} stroke="#94a3b8" fontSize={11} />
              <Tooltip />
              <Bar dataKey="count" fill="#10b981" radius={[0, 6, 6, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>

        <div className="card p-5">
          <h3 className="mb-4 text-sm font-semibold text-slate-700 dark:text-slate-200">Clients les plus actifs</h3>
          <div className="divide-y divide-slate-100 dark:divide-navy-700">
            {report.mostActiveClients.map((c, i) => (
              <div key={i} className="flex items-center justify-between py-2.5 text-sm">
                <span className="text-slate-700 dark:text-slate-200">{c.client?.firstName} {c.client?.lastName}</span>
                <span className="badge bg-brand-50 text-brand-700 dark:bg-navy-700 dark:text-brand-200">{c.count} location(s)</span>
              </div>
            ))}
            {report.mostActiveClients.length === 0 && <p className="py-4 text-sm text-slate-400">Aucune donnée.</p>}
          </div>
        </div>
      </div>
    </div>
  );
}
