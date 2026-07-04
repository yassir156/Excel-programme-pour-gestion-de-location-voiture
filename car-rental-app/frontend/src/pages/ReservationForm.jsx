import { useEffect, useState } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import { Save, X, CheckCircle2, AlertCircle } from 'lucide-react';
import api, { getErrorMessage } from '../api/client';
import PageHeader from '../components/PageHeader';

export default function ReservationForm() {
  const { id } = useParams();
  const isEdit = !!id;
  const navigate = useNavigate();

  const [clients, setClients] = useState([]);
  const [vehicles, setVehicles] = useState([]);
  const [form, setForm] = useState({ clientId: '', vehicleId: '', startDate: '', endDate: '', notes: '' });
  const [availability, setAvailability] = useState(null);
  const [checking, setChecking] = useState(false);
  const [errors, setErrors] = useState({});
  const [submitError, setSubmitError] = useState('');
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    api.get('/clients').then((res) => setClients(res.data));
    api.get('/vehicles').then((res) => setVehicles(res.data));
  }, []);

  useEffect(() => {
    if (!isEdit) return;
    api.get(`/reservations/${id}`).then((res) => {
      const r = res.data;
      setForm({
        clientId: r.clientId, vehicleId: r.vehicleId, startDate: r.startDate, endDate: r.endDate, notes: r.notes || '',
      });
    });
  }, [id, isEdit]);

  useEffect(() => {
    if (!form.vehicleId || !form.startDate || !form.endDate) {
      setAvailability(null);
      return;
    }
    setChecking(true);
    const params = { vehicleId: form.vehicleId, startDate: form.startDate, endDate: form.endDate };
    if (isEdit) params.excludeReservationId = id;
    const timeout = setTimeout(() => {
      api
        .get('/reservations/check-availability', { params })
        .then((res) => setAvailability(res.data))
        .catch(() => setAvailability(null))
        .finally(() => setChecking(false));
    }, 300);
    return () => clearTimeout(timeout);
  }, [form.vehicleId, form.startDate, form.endDate, isEdit, id]);

  function update(field, value) {
    setForm((f) => ({ ...f, [field]: value }));
  }

  function validate() {
    const errs = {};
    if (!form.clientId) errs.clientId = 'Client requis';
    if (!form.vehicleId) errs.vehicleId = 'Véhicule requis';
    if (!form.startDate) errs.startDate = 'Date de début requise';
    if (!form.endDate) errs.endDate = 'Date de fin requise';
    if (form.startDate && form.endDate && new Date(form.startDate) >= new Date(form.endDate)) {
      errs.endDate = 'La date de fin doit être après la date de début';
    }
    if (availability && !availability.available) {
      errs.vehicleId = 'Ce véhicule est déjà réservé sur cette période';
    }
    setErrors(errs);
    return Object.keys(errs).length === 0;
  }

  async function handleSubmit(e) {
    e.preventDefault();
    setSubmitError('');
    if (!validate()) return;
    setSaving(true);
    try {
      if (isEdit) {
        await api.put(`/reservations/${id}`, form);
      } else {
        await api.post('/reservations', form);
      }
      navigate('/reservations');
    } catch (err) {
      setSubmitError(getErrorMessage(err, "Impossible d'enregistrer la réservation."));
    } finally {
      setSaving(false);
    }
  }

  return (
    <div>
      <PageHeader title={isEdit ? 'Modifier la réservation' : 'Nouvelle réservation'} />

      <form onSubmit={handleSubmit} className="card max-w-3xl space-y-6 p-6">
        {submitError && <div className="rounded-lg bg-red-50 px-4 py-2 text-sm text-red-600">{submitError}</div>}

        <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
          <Field label="Client" error={errors.clientId}>
            <select className="input" value={form.clientId} onChange={(e) => update('clientId', e.target.value)}>
              <option value="">Sélectionner un client</option>
              {clients.map((c) => (
                <option key={c.id} value={c.id}>{c.firstName} {c.lastName} — {c.phone}</option>
              ))}
            </select>
          </Field>
          <Field label="Véhicule" error={errors.vehicleId}>
            <select className="input" value={form.vehicleId} onChange={(e) => update('vehicleId', e.target.value)}>
              <option value="">Sélectionner un véhicule</option>
              {vehicles.map((v) => (
                <option key={v.id} value={v.id}>{v.brand} {v.model} ({v.plate}) — {v.dailyPrice} MAD/j</option>
              ))}
            </select>
          </Field>
          <Field label="Date de début" error={errors.startDate}>
            <input type="date" className="input" value={form.startDate} onChange={(e) => update('startDate', e.target.value)} />
          </Field>
          <Field label="Date de fin" error={errors.endDate}>
            <input type="date" className="input" value={form.endDate} onChange={(e) => update('endDate', e.target.value)} />
          </Field>
        </div>

        <Field label="Notes">
          <textarea className="input" rows={3} value={form.notes} onChange={(e) => update('notes', e.target.value)} />
        </Field>

        {checking && <p className="text-sm text-slate-400">Vérification de la disponibilité...</p>}
        {!checking && availability && (
          <div
            className={`flex items-center gap-2 rounded-lg px-4 py-3 text-sm ${
              availability.available
                ? 'bg-emerald-50 text-emerald-700 dark:bg-emerald-950 dark:text-emerald-300'
                : 'bg-red-50 text-red-700 dark:bg-red-950 dark:text-red-300'
            }`}
          >
            {availability.available ? <CheckCircle2 size={16} /> : <AlertCircle size={16} />}
            {availability.available
              ? `Véhicule disponible — ${availability.days} jour(s), total estimé ${availability.totalPrice} MAD`
              : 'Ce véhicule est déjà réservé sur cette période.'}
          </div>
        )}

        <div className="flex justify-end gap-3 border-t border-slate-100 pt-4 dark:border-navy-700">
          <button type="button" className="btn-secondary" onClick={() => navigate('/reservations')}>
            <X size={16} /> Annuler
          </button>
          <button type="submit" className="btn-primary" disabled={saving}>
            <Save size={16} /> {saving ? 'Enregistrement...' : 'Enregistrer'}
          </button>
        </div>
      </form>
    </div>
  );
}

function Field({ label, error, children }) {
  return (
    <div>
      <label className="label">{label}</label>
      {children}
      {error && <p className="mt-1 text-xs text-red-500">{error}</p>}
    </div>
  );
}
