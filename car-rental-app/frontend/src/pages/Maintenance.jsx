import { useEffect, useState } from 'react';
import { format } from 'date-fns';
import { Plus, Trash2 } from 'lucide-react';
import api, { getErrorMessage } from '../api/client';
import { useAuth } from '../context/AuthContext';
import DataTable from '../components/DataTable';
import PageHeader from '../components/PageHeader';
import Modal from '../components/Modal';
import ConfirmDialog from '../components/ConfirmDialog';

const TYPE_LABELS = {
  vidange: 'Vidange', reparation: 'Réparation', pneus: 'Pneus',
  assurance: 'Assurance', controle_technique: 'Contrôle technique', autre: 'Autre',
};

const emptyForm = { vehicleId: '', type: 'vidange', date: '', cost: '', description: '', nextDueDate: '' };

export default function Maintenance() {
  const { user } = useAuth();
  const [items, setItems] = useState([]);
  const [vehicles, setVehicles] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [modalOpen, setModalOpen] = useState(false);
  const [form, setForm] = useState(emptyForm);
  const [formError, setFormError] = useState('');
  const [saving, setSaving] = useState(false);
  const [toDelete, setToDelete] = useState(null);

  const canDelete = ['administrateur', 'manager'].includes(user?.role);

  function load() {
    setLoading(true);
    Promise.all([api.get('/maintenance'), api.get('/vehicles')])
      .then(([m, v]) => {
        setItems(m.data);
        setVehicles(v.data);
      })
      .catch((err) => setError(getErrorMessage(err)))
      .finally(() => setLoading(false));
  }

  useEffect(load, []);

  function openModal() {
    setForm(emptyForm);
    setFormError('');
    setModalOpen(true);
  }

  async function handleSubmit(e) {
    e.preventDefault();
    setFormError('');
    if (!form.vehicleId || !form.type || !form.date) {
      setFormError('Véhicule, type et date sont requis.');
      return;
    }
    setSaving(true);
    try {
      await api.post('/maintenance', { ...form, cost: Number(form.cost) || 0, nextDueDate: form.nextDueDate || null });
      setModalOpen(false);
      load();
    } catch (err) {
      setFormError(getErrorMessage(err));
    } finally {
      setSaving(false);
    }
  }

  async function confirmDelete() {
    try {
      await api.delete(`/maintenance/${toDelete.id}`);
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
      searchValue: (r) => `${r.vehicle?.brand} ${r.vehicle?.model} ${r.vehicle?.plate}`,
      render: (r) => `${r.vehicle?.brand} ${r.vehicle?.model} (${r.vehicle?.plate})`,
    },
    { key: 'type', header: 'Type', render: (r) => TYPE_LABELS[r.type] || r.type },
    { key: 'date', header: 'Date', render: (r) => format(new Date(r.date), 'dd/MM/yyyy') },
    { key: 'cost', header: 'Coût', render: (r) => `${r.cost} MAD` },
    { key: 'description', header: 'Description', render: (r) => r.description || '—' },
    {
      key: 'nextDueDate',
      header: 'Prochaine échéance',
      render: (r) => (r.nextDueDate ? format(new Date(r.nextDueDate), 'dd/MM/yyyy') : '—'),
    },
    ...(canDelete
      ? [{
          key: 'actions',
          header: 'Actions',
          render: (r) => (
            <button className="btn-danger px-2 py-1" onClick={() => setToDelete(r)}>
              <Trash2 size={15} />
            </button>
          ),
        }]
      : []),
  ];

  return (
    <div>
      <PageHeader
        title="Maintenance"
        subtitle={`${items.length} opération(s) enregistrée(s)`}
        actions={
          <button className="btn-primary" onClick={openModal}>
            <Plus size={16} /> Ajouter une opération
          </button>
        }
      />

      {error && <div className="mb-4 rounded-lg bg-red-50 px-4 py-2 text-sm text-red-600">{error}</div>}

      <DataTable
        columns={columns}
        data={items}
        searchPlaceholder="Rechercher par véhicule..."
        emptyMessage={loading ? 'Chargement...' : 'Aucune opération de maintenance enregistrée.'}
      />

      <Modal open={modalOpen} title="Ajouter une opération de maintenance" onClose={() => setModalOpen(false)}>
        <form onSubmit={handleSubmit} className="space-y-4">
          {formError && <div className="rounded-lg bg-red-50 px-4 py-2 text-sm text-red-600">{formError}</div>}

          <div>
            <label className="label">Véhicule</label>
            <select className="input" value={form.vehicleId} onChange={(e) => setForm((f) => ({ ...f, vehicleId: e.target.value }))}>
              <option value="">Sélectionner un véhicule</option>
              {vehicles.map((v) => (
                <option key={v.id} value={v.id}>{v.brand} {v.model} ({v.plate})</option>
              ))}
            </select>
          </div>

          <div className="grid grid-cols-2 gap-4">
            <div>
              <label className="label">Type</label>
              <select className="input" value={form.type} onChange={(e) => setForm((f) => ({ ...f, type: e.target.value }))}>
                {Object.entries(TYPE_LABELS).map(([value, label]) => (
                  <option key={value} value={value}>{label}</option>
                ))}
              </select>
            </div>
            <div>
              <label className="label">Date</label>
              <input type="date" className="input" value={form.date} onChange={(e) => setForm((f) => ({ ...f, date: e.target.value }))} />
            </div>
          </div>

          <div className="grid grid-cols-2 gap-4">
            <div>
              <label className="label">Coût (MAD)</label>
              <input type="number" className="input" value={form.cost} onChange={(e) => setForm((f) => ({ ...f, cost: e.target.value }))} />
            </div>
            <div>
              <label className="label">Prochaine échéance</label>
              <input type="date" className="input" value={form.nextDueDate} onChange={(e) => setForm((f) => ({ ...f, nextDueDate: e.target.value }))} />
            </div>
          </div>

          <div>
            <label className="label">Description</label>
            <textarea className="input" rows={2} value={form.description} onChange={(e) => setForm((f) => ({ ...f, description: e.target.value }))} />
          </div>

          <div className="flex justify-end gap-3 pt-2">
            <button type="button" className="btn-secondary" onClick={() => setModalOpen(false)}>Annuler</button>
            <button type="submit" className="btn-primary" disabled={saving}>{saving ? 'Enregistrement...' : 'Enregistrer'}</button>
          </div>
        </form>
      </Modal>

      <ConfirmDialog
        open={!!toDelete}
        message="Supprimer cette opération de maintenance ?"
        onConfirm={confirmDelete}
        onCancel={() => setToDelete(null)}
      />
    </div>
  );
}
