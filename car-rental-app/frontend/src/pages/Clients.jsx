import { useEffect, useState } from 'react';
import { Link, useNavigate } from 'react-router-dom';
import { Plus, Pencil, Trash2, Eye, AlertTriangle } from 'lucide-react';
import { differenceInDays } from 'date-fns';
import api, { getErrorMessage } from '../api/client';
import { useAuth } from '../context/AuthContext';
import DataTable from '../components/DataTable';
import PageHeader from '../components/PageHeader';
import ConfirmDialog from '../components/ConfirmDialog';

export default function Clients() {
  const { user } = useAuth();
  const navigate = useNavigate();
  const [clients, setClients] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [toDelete, setToDelete] = useState(null);

  const canDelete = ['administrateur', 'manager'].includes(user?.role);

  function load() {
    setLoading(true);
    api.get('/clients').then((res) => setClients(res.data)).catch((err) => setError(getErrorMessage(err))).finally(() => setLoading(false));
  }

  useEffect(load, []);

  async function confirmDelete() {
    try {
      await api.delete(`/clients/${toDelete.id}`);
      setToDelete(null);
      load();
    } catch (err) {
      setError(getErrorMessage(err));
      setToDelete(null);
    }
  }

  const columns = [
    {
      key: 'name',
      header: 'Client',
      searchValue: (r) => `${r.firstName} ${r.lastName} ${r.phone} ${r.cin}`,
      render: (r) => (
        <div>
          <p className="font-medium text-slate-800 dark:text-white">{r.firstName} {r.lastName}</p>
          <p className="text-xs text-slate-400">CIN: {r.cin}</p>
        </div>
      ),
    },
    { key: 'phone', header: 'Téléphone' },
    { key: 'email', header: 'Email', render: (r) => r.email || '—' },
    {
      key: 'license',
      header: 'Permis',
      render: (r) => {
        if (!r.licenseExpiry) return '—';
        const days = differenceInDays(new Date(r.licenseExpiry), new Date());
        const soon = days <= 30;
        return (
          <span className={`flex items-center gap-1 ${soon ? 'text-amber-600' : ''}`}>
            {soon && <AlertTriangle size={13} />}
            {new Date(r.licenseExpiry).toLocaleDateString('fr-FR')}
          </span>
        );
      },
    },
    {
      key: 'actions',
      header: 'Actions',
      render: (r) => (
        <div className="flex gap-2">
          <button className="btn-secondary px-2 py-1" onClick={() => navigate(`/clients/${r.id}`)} title="Détails">
            <Eye size={15} />
          </button>
          <button className="btn-secondary px-2 py-1" onClick={() => navigate(`/clients/${r.id}/edit`)} title="Modifier">
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
        title="Clients"
        subtitle={`${clients.length} client(s) enregistré(s)`}
        actions={
          <Link to="/clients/new" className="btn-primary">
            <Plus size={16} /> Ajouter un client
          </Link>
        }
      />

      {error && <div className="mb-4 rounded-lg bg-red-50 px-4 py-2 text-sm text-red-600">{error}</div>}

      <DataTable
        columns={columns}
        data={clients}
        searchPlaceholder="Rechercher par nom, téléphone, CIN..."
        emptyMessage={loading ? 'Chargement...' : 'Aucun client trouvé.'}
      />

      <ConfirmDialog
        open={!!toDelete}
        message={`Supprimer le client ${toDelete?.firstName} ${toDelete?.lastName} ?`}
        onConfirm={confirmDelete}
        onCancel={() => setToDelete(null)}
      />
    </div>
  );
}
