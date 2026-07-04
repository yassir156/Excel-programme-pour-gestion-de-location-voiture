import { useEffect, useState } from 'react';
import { Link, useNavigate } from 'react-router-dom';
import { format } from 'date-fns';
import { Plus, Pencil, XCircle, FileText } from 'lucide-react';
import api, { getErrorMessage } from '../api/client';
import DataTable from '../components/DataTable';
import PageHeader from '../components/PageHeader';
import StatusBadge from '../components/StatusBadge';
import ConfirmDialog from '../components/ConfirmDialog';

export default function Reservations() {
  const navigate = useNavigate();
  const [reservations, setReservations] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [statusFilter, setStatusFilter] = useState('');
  const [toCancel, setToCancel] = useState(null);

  function load() {
    setLoading(true);
    api
      .get('/reservations', { params: statusFilter ? { status: statusFilter } : {} })
      .then((res) => setReservations(res.data))
      .catch((err) => setError(getErrorMessage(err)))
      .finally(() => setLoading(false));
  }

  useEffect(load, [statusFilter]);

  async function confirmCancel() {
    try {
      await api.post(`/reservations/${toCancel.id}/cancel`);
      setToCancel(null);
      load();
    } catch (err) {
      setError(getErrorMessage(err));
      setToCancel(null);
    }
  }

  async function generateContract(reservation) {
    try {
      await api.post(`/contracts/from-reservation/${reservation.id}`);
      navigate('/contracts');
    } catch (err) {
      setError(getErrorMessage(err));
    }
  }

  const columns = [
    {
      key: 'client',
      header: 'Client',
      searchValue: (r) => `${r.client?.firstName} ${r.client?.lastName}`,
      render: (r) => `${r.client?.firstName} ${r.client?.lastName}`,
    },
    {
      key: 'vehicle',
      header: 'Véhicule',
      searchValue: (r) => `${r.vehicle?.brand} ${r.vehicle?.model} ${r.vehicle?.plate}`,
      render: (r) => `${r.vehicle?.brand} ${r.vehicle?.model} (${r.vehicle?.plate})`,
    },
    {
      key: 'dates',
      header: 'Période',
      render: (r) => `${format(new Date(r.startDate), 'dd/MM/yy')} → ${format(new Date(r.endDate), 'dd/MM/yy')} (${r.days}j)`,
    },
    { key: 'totalPrice', header: 'Total', render: (r) => `${r.totalPrice} MAD` },
    { key: 'status', header: 'Statut', render: (r) => <StatusBadge status={r.status} /> },
    {
      key: 'actions',
      header: 'Actions',
      render: (r) => (
        <div className="flex gap-2">
          {['en_attente', 'confirmee'].includes(r.status) && (
            <button className="btn-secondary px-2 py-1" onClick={() => navigate(`/reservations/${r.id}/edit`)} title="Modifier">
              <Pencil size={15} />
            </button>
          )}
          {r.status !== 'annulee' && r.status !== 'terminee' && (
            <button className="btn-secondary px-2 py-1" onClick={() => generateContract(r)} title="Générer le contrat">
              <FileText size={15} />
            </button>
          )}
          {['en_attente', 'confirmee'].includes(r.status) && (
            <button className="btn-danger px-2 py-1" onClick={() => setToCancel(r)} title="Annuler">
              <XCircle size={15} />
            </button>
          )}
        </div>
      ),
    },
  ];

  return (
    <div>
      <PageHeader
        title="Réservations"
        subtitle={`${reservations.length} réservation(s)`}
        actions={
          <Link to="/reservations/new" className="btn-primary">
            <Plus size={16} /> Nouvelle réservation
          </Link>
        }
      />

      {error && <div className="mb-4 rounded-lg bg-red-50 px-4 py-2 text-sm text-red-600">{error}</div>}

      <DataTable
        columns={columns}
        data={reservations}
        searchPlaceholder="Rechercher par client ou véhicule..."
        emptyMessage={loading ? 'Chargement...' : 'Aucune réservation trouvée.'}
        extraFilters={
          <select className="input w-44" value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}>
            <option value="">Tous les statuts</option>
            <option value="en_attente">En attente</option>
            <option value="confirmee">Confirmée</option>
            <option value="annulee">Annulée</option>
            <option value="terminee">Terminée</option>
          </select>
        }
      />

      <ConfirmDialog
        open={!!toCancel}
        title="Annuler la réservation"
        message="Voulez-vous vraiment annuler cette réservation ?"
        onConfirm={confirmCancel}
        onCancel={() => setToCancel(null)}
      />
    </div>
  );
}
