import { useEffect, useState } from 'react';
import { format } from 'date-fns';
import { Plus, FileDown } from 'lucide-react';
import api, { getErrorMessage } from '../api/client';
import DataTable from '../components/DataTable';
import PageHeader from '../components/PageHeader';
import StatusBadge from '../components/StatusBadge';
import Modal from '../components/Modal';
import { generateReceiptPDF } from '../utils/pdf';

const METHOD_LABELS = { especes: 'Espèces', carte: 'Carte bancaire', virement: 'Virement', cheque: 'Chèque' };

export default function Payments() {
  const [payments, setPayments] = useState([]);
  const [reservations, setReservations] = useState([]);
  const [settings, setSettings] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [modalOpen, setModalOpen] = useState(false);
  const [form, setForm] = useState({ reservationId: '', amount: '', method: 'especes', depositPaid: false });
  const [formError, setFormError] = useState('');
  const [saving, setSaving] = useState(false);

  function load() {
    setLoading(true);
    Promise.all([api.get('/payments'), api.get('/reservations'), api.get('/settings')])
      .then(([p, r, s]) => {
        setPayments(p.data);
        setReservations(r.data.filter((x) => x.status !== 'annulee'));
        setSettings(s.data);
      })
      .catch((err) => setError(getErrorMessage(err)))
      .finally(() => setLoading(false));
  }

  useEffect(load, []);

  function openModal() {
    setForm({ reservationId: '', amount: '', method: 'especes', depositPaid: false });
    setFormError('');
    setModalOpen(true);
  }

  async function handleSubmit(e) {
    e.preventDefault();
    setFormError('');
    if (!form.reservationId || !form.amount || Number(form.amount) <= 0) {
      setFormError('Réservation et montant valide sont requis.');
      return;
    }
    setSaving(true);
    try {
      await api.post('/payments', { ...form, amount: Number(form.amount) });
      setModalOpen(false);
      load();
    } catch (err) {
      setFormError(getErrorMessage(err));
    } finally {
      setSaving(false);
    }
  }

  const columns = [
    { key: 'receiptNumber', header: 'N° Reçu' },
    {
      key: 'client',
      header: 'Client',
      searchValue: (r) => `${r.client?.firstName} ${r.client?.lastName}`,
      render: (r) => `${r.client?.firstName} ${r.client?.lastName}`,
    },
    { key: 'amount', header: 'Montant payé', render: (r) => `${r.amount} MAD` },
    { key: 'remaining', header: 'Reste à payer', render: (r) => `${r.remaining} MAD` },
    { key: 'method', header: 'Mode', render: (r) => METHOD_LABELS[r.method] || r.method },
    { key: 'status', header: 'Statut', render: (r) => <StatusBadge status={r.status} /> },
    { key: 'paidAt', header: 'Date', render: (r) => format(new Date(r.paidAt), 'dd/MM/yyyy') },
    {
      key: 'actions',
      header: 'Actions',
      render: (r) => (
        <button className="btn-secondary px-2 py-1" onClick={() => generateReceiptPDF(r, settings)} title="Télécharger le reçu">
          <FileDown size={15} />
        </button>
      ),
    },
  ];

  return (
    <div>
      <PageHeader
        title="Paiements"
        subtitle={`${payments.length} paiement(s) enregistré(s)`}
        actions={
          <button className="btn-primary" onClick={openModal}>
            <Plus size={16} /> Enregistrer un paiement
          </button>
        }
      />

      {error && <div className="mb-4 rounded-lg bg-red-50 px-4 py-2 text-sm text-red-600">{error}</div>}

      <DataTable
        columns={columns}
        data={payments}
        searchPlaceholder="Rechercher par client..."
        emptyMessage={loading ? 'Chargement...' : 'Aucun paiement enregistré.'}
      />

      <Modal open={modalOpen} title="Enregistrer un paiement" onClose={() => setModalOpen(false)}>
        <form onSubmit={handleSubmit} className="space-y-4">
          {formError && <div className="rounded-lg bg-red-50 px-4 py-2 text-sm text-red-600">{formError}</div>}

          <div>
            <label className="label">Réservation</label>
            <select
              className="input"
              value={form.reservationId}
              onChange={(e) => setForm((f) => ({ ...f, reservationId: e.target.value }))}
            >
              <option value="">Sélectionner une réservation</option>
              {reservations.map((r) => (
                <option key={r.id} value={r.id}>
                  {r.client?.firstName} {r.client?.lastName} — {r.vehicle?.brand} {r.vehicle?.model} — {r.totalPrice} MAD
                </option>
              ))}
            </select>
          </div>

          <div>
            <label className="label">Montant (MAD)</label>
            <input
              type="number"
              className="input"
              value={form.amount}
              onChange={(e) => setForm((f) => ({ ...f, amount: e.target.value }))}
            />
          </div>

          <div>
            <label className="label">Mode de paiement</label>
            <select className="input" value={form.method} onChange={(e) => setForm((f) => ({ ...f, method: e.target.value }))}>
              <option value="especes">Espèces</option>
              <option value="carte">Carte bancaire</option>
              <option value="virement">Virement</option>
              <option value="cheque">Chèque</option>
            </select>
          </div>

          <label className="flex items-center gap-2 text-sm text-slate-600 dark:text-slate-300">
            <input
              type="checkbox"
              checked={form.depositPaid}
              onChange={(e) => setForm((f) => ({ ...f, depositPaid: e.target.checked }))}
            />
            Caution payée
          </label>

          <div className="flex justify-end gap-3 pt-2">
            <button type="button" className="btn-secondary" onClick={() => setModalOpen(false)}>Annuler</button>
            <button type="submit" className="btn-primary" disabled={saving}>{saving ? 'Enregistrement...' : 'Enregistrer'}</button>
          </div>
        </form>
      </Modal>
    </div>
  );
}
