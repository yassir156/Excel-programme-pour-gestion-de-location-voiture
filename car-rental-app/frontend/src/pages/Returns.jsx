import { useEffect, useMemo, useState } from 'react';
import { format } from 'date-fns';
import { Plus } from 'lucide-react';
import api, { getErrorMessage } from '../api/client';
import DataTable from '../components/DataTable';
import PageHeader from '../components/PageHeader';
import Modal from '../components/Modal';

const emptyForm = {
  reservationId: '', mileageReturn: '', fuelReturn: '', condition: '',
  lateFee: 0, fuelFee: 0, damageFee: 0, extraMileageFee: 0, notes: '',
};

export default function Returns() {
  const [returns, setReturns] = useState([]);
  const [reservations, setReservations] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [modalOpen, setModalOpen] = useState(false);
  const [form, setForm] = useState(emptyForm);
  const [formError, setFormError] = useState('');
  const [saving, setSaving] = useState(false);

  function load() {
    setLoading(true);
    Promise.all([api.get('/returns'), api.get('/reservations', { params: { status: 'confirmee' } })])
      .then(([r, res]) => {
        setReturns(r.data);
        const alreadyReturned = new Set(r.data.map((x) => x.reservationId));
        setReservations(res.data.filter((x) => !alreadyReturned.has(x.id)));
      })
      .catch((err) => setError(getErrorMessage(err)))
      .finally(() => setLoading(false));
  }

  useEffect(load, []);

  const totalExtra = useMemo(
    () => Number(form.lateFee || 0) + Number(form.fuelFee || 0) + Number(form.damageFee || 0) + Number(form.extraMileageFee || 0),
    [form.lateFee, form.fuelFee, form.damageFee, form.extraMileageFee]
  );

  function openModal() {
    setForm(emptyForm);
    setFormError('');
    setModalOpen(true);
  }

  async function handleSubmit(e) {
    e.preventDefault();
    setFormError('');
    if (!form.reservationId || form.mileageReturn === '') {
      setFormError('Réservation et kilométrage de retour sont requis.');
      return;
    }
    setSaving(true);
    try {
      await api.post('/returns', form);
      setModalOpen(false);
      load();
    } catch (err) {
      setFormError(getErrorMessage(err));
    } finally {
      setSaving(false);
    }
  }

  const columns = [
    {
      key: 'vehicle',
      header: 'Véhicule',
      render: (r) => `${r.vehicle?.brand} ${r.vehicle?.model} (${r.vehicle?.plate})`,
    },
    {
      key: 'client',
      header: 'Client',
      render: (r) => `${r.reservation?.client?.firstName || ''} ${r.reservation?.client?.lastName || ''}`,
    },
    { key: 'returnDate', header: 'Date de retour', render: (r) => format(new Date(r.returnDate), 'dd/MM/yyyy HH:mm') },
    { key: 'mileageReturn', header: 'Kilométrage', render: (r) => `${r.mileageReturn.toLocaleString()} km` },
    { key: 'condition', header: 'État', render: (r) => r.condition || '—' },
    { key: 'totalExtra', header: 'Frais supplémentaires', render: (r) => `${r.totalExtra} MAD` },
  ];

  return (
    <div>
      <PageHeader
        title="Retours de véhicules"
        subtitle={`${returns.length} retour(s) enregistré(s)`}
        actions={
          <button className="btn-primary" onClick={openModal}>
            <Plus size={16} /> Enregistrer un retour
          </button>
        }
      />

      {error && <div className="mb-4 rounded-lg bg-red-50 px-4 py-2 text-sm text-red-600">{error}</div>}

      <DataTable
        columns={columns}
        data={returns}
        searchPlaceholder="Rechercher..."
        emptyMessage={loading ? 'Chargement...' : 'Aucun retour enregistré.'}
      />

      <Modal open={modalOpen} title="Enregistrer un retour" onClose={() => setModalOpen(false)} size="lg">
        <form onSubmit={handleSubmit} className="space-y-4">
          {formError && <div className="rounded-lg bg-red-50 px-4 py-2 text-sm text-red-600">{formError}</div>}

          <div>
            <label className="label">Réservation (véhicule loué)</label>
            <select className="input" value={form.reservationId} onChange={(e) => setForm((f) => ({ ...f, reservationId: e.target.value }))}>
              <option value="">Sélectionner une réservation confirmée</option>
              {reservations.map((r) => (
                <option key={r.id} value={r.id}>
                  {r.client?.firstName} {r.client?.lastName} — {r.vehicle?.brand} {r.vehicle?.model} ({r.vehicle?.plate})
                </option>
              ))}
            </select>
          </div>

          <div className="grid grid-cols-1 gap-4 sm:grid-cols-3">
            <div>
              <label className="label">Kilométrage retour</label>
              <input type="number" className="input" value={form.mileageReturn} onChange={(e) => setForm((f) => ({ ...f, mileageReturn: e.target.value }))} />
            </div>
            <div>
              <label className="label">Carburant retour</label>
              <input className="input" placeholder="ex: 3/4" value={form.fuelReturn} onChange={(e) => setForm((f) => ({ ...f, fuelReturn: e.target.value }))} />
            </div>
            <div>
              <label className="label">État du véhicule</label>
              <input className="input" placeholder="Bon état..." value={form.condition} onChange={(e) => setForm((f) => ({ ...f, condition: e.target.value }))} />
            </div>
          </div>

          <div className="grid grid-cols-1 gap-4 sm:grid-cols-4">
            <div>
              <label className="label">Frais de retard</label>
              <input type="number" className="input" value={form.lateFee} onChange={(e) => setForm((f) => ({ ...f, lateFee: e.target.value }))} />
            </div>
            <div>
              <label className="label">Carburant manquant</label>
              <input type="number" className="input" value={form.fuelFee} onChange={(e) => setForm((f) => ({ ...f, fuelFee: e.target.value }))} />
            </div>
            <div>
              <label className="label">Dommages</label>
              <input type="number" className="input" value={form.damageFee} onChange={(e) => setForm((f) => ({ ...f, damageFee: e.target.value }))} />
            </div>
            <div>
              <label className="label">Km supplémentaires</label>
              <input type="number" className="input" value={form.extraMileageFee} onChange={(e) => setForm((f) => ({ ...f, extraMileageFee: e.target.value }))} />
            </div>
          </div>

          <div>
            <label className="label">Notes</label>
            <textarea className="input" rows={2} value={form.notes} onChange={(e) => setForm((f) => ({ ...f, notes: e.target.value }))} />
          </div>

          <div className="rounded-lg bg-slate-50 px-4 py-2 text-sm font-medium text-slate-700 dark:bg-navy-900 dark:text-slate-200">
            Total des frais supplémentaires : {totalExtra} MAD
          </div>

          <div className="flex justify-end gap-3 pt-2">
            <button type="button" className="btn-secondary" onClick={() => setModalOpen(false)}>Annuler</button>
            <button type="submit" className="btn-primary" disabled={saving}>{saving ? 'Enregistrement...' : 'Enregistrer'}</button>
          </div>
        </form>
      </Modal>
    </div>
  );
}
