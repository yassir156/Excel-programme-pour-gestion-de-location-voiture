import { useEffect, useState } from 'react';
import { Link, useNavigate } from 'react-router-dom';
import { Plus, Pencil, Trash2, Eye, Car as CarIcon } from 'lucide-react';
import api, { getErrorMessage, uploadsBaseUrl } from '../api/client';
import { useAuth } from '../context/AuthContext';
import DataTable from '../components/DataTable';
import PageHeader from '../components/PageHeader';
import StatusBadge from '../components/StatusBadge';
import ConfirmDialog from '../components/ConfirmDialog';

export default function Vehicles() {
  const { user } = useAuth();
  const navigate = useNavigate();
  const [vehicles, setVehicles] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [statusFilter, setStatusFilter] = useState('');
  const [toDelete, setToDelete] = useState(null);

  const canDelete = ['administrateur', 'manager'].includes(user?.role);

  function load() {
    setLoading(true);
    api
      .get('/vehicles', { params: statusFilter ? { status: statusFilter } : {} })
      .then((res) => setVehicles(res.data))
      .catch((err) => setError(getErrorMessage(err)))
      .finally(() => setLoading(false));
  }

  useEffect(load, [statusFilter]);

  async function confirmDelete() {
    try {
      await api.delete(`/vehicles/${toDelete.id}`);
      setToDelete(null);
      load();
    } catch (err) {
      setError(getErrorMessage(err));
      setToDelete(null);
    }
  }

  const columns = [
    {
      key: 'vehicle',
      header: 'Véhicule',
      searchValue: (r) => `${r.brand} ${r.model} ${r.plate}`,
      render: (r) => (
        <div className="flex items-center gap-3">
          <div className="flex h-10 w-10 items-center justify-center overflow-hidden rounded-lg bg-slate-100 dark:bg-navy-700">
            {r.photos?.[0] ? (
              <img src={`${uploadsBaseUrl}${r.photos[0]}`} alt="" className="h-full w-full object-cover" />
            ) : (
              <CarIcon size={18} className="text-slate-400" />
            )}
          </div>
          <div>
            <p className="font-medium text-slate-800 dark:text-white">{r.brand} {r.model}</p>
            <p className="text-xs text-slate-400">{r.plate} · {r.year}</p>
          </div>
        </div>
      ),
    },
    { key: 'category', header: 'Catégorie', render: (r) => r.category || '—' },
    { key: 'transmission', header: 'Boîte', render: (r) => (r.transmission === 'automatique' ? 'Automatique' : 'Manuelle') },
    { key: 'dailyPrice', header: 'Prix / jour', render: (r) => `${r.dailyPrice} MAD` },
    { key: 'status', header: 'Statut', render: (r) => <StatusBadge status={r.status} /> },
    {
      key: 'actions',
      header: 'Actions',
      render: (r) => (
        <div className="flex gap-2">
          <button className="btn-secondary px-2 py-1" onClick={() => navigate(`/vehicles/${r.id}`)} title="Détails">
            <Eye size={15} />
          </button>
          <button className="btn-secondary px-2 py-1" onClick={() => navigate(`/vehicles/${r.id}/edit`)} title="Modifier">
            <Pencil size={15} />
          </button>
          {canDelete && (
            <button className="btn-danger px-2 py-1" onClick={() => setToDelete(r)} title="Supprimer">
              <Trash2 size={15} />
            </button>
          )}
        </div>
      ),
    },
  ];

  return (
    <div>
      <PageHeader
        title="Véhicules"
        subtitle={`${vehicles.length} véhicule(s) enregistré(s)`}
        actions={
          <Link to="/vehicles/new" className="btn-primary">
            <Plus size={16} /> Ajouter un véhicule
          </Link>
        }
      />

      {error && <div className="mb-4 rounded-lg bg-red-50 px-4 py-2 text-sm text-red-600">{error}</div>}

      <DataTable
        columns={columns}
        data={vehicles}
        searchPlaceholder="Rechercher par marque, modèle, immatriculation..."
        emptyMessage={loading ? 'Chargement...' : 'Aucun véhicule trouvé.'}
        extraFilters={
          <select className="input w-40" value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}>
            <option value="">Tous les statuts</option>
            <option value="disponible">Disponible</option>
            <option value="louee">Louée</option>
            <option value="maintenance">Maintenance</option>
            <option value="reservee">Réservée</option>
          </select>
        }
      />

      <ConfirmDialog
        open={!!toDelete}
        message={`Supprimer le véhicule ${toDelete?.brand} ${toDelete?.model} (${toDelete?.plate}) ?`}
        onConfirm={confirmDelete}
        onCancel={() => setToDelete(null)}
      />
    </div>
  );
}
