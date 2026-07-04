import { useEffect, useState } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import { Save, Upload, X } from 'lucide-react';
import api, { getErrorMessage, uploadsBaseUrl } from '../api/client';
import PageHeader from '../components/PageHeader';

const emptyForm = {
  firstName: '', lastName: '', phone: '', email: '', address: '', cin: '', licenseNumber: '', licenseExpiry: '',
};

export default function ClientForm() {
  const { id } = useParams();
  const isEdit = !!id;
  const navigate = useNavigate();
  const [form, setForm] = useState(emptyForm);
  const [documents, setDocuments] = useState([]);
  const [files, setFiles] = useState([]);
  const [errors, setErrors] = useState({});
  const [submitError, setSubmitError] = useState('');
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    if (!isEdit) return;
    api.get(`/clients/${id}`).then((res) => {
      const c = res.data;
      setForm({
        firstName: c.firstName, lastName: c.lastName, phone: c.phone, email: c.email || '',
        address: c.address || '', cin: c.cin, licenseNumber: c.licenseNumber || '', licenseExpiry: c.licenseExpiry || '',
      });
      setDocuments(c.documents || []);
    });
  }, [id, isEdit]);

  function update(field, value) {
    setForm((f) => ({ ...f, [field]: value }));
  }

  function validate() {
    const errs = {};
    if (!form.firstName) errs.firstName = 'Prénom requis';
    if (!form.lastName) errs.lastName = 'Nom requis';
    if (!form.phone) errs.phone = 'Téléphone requis';
    if (!form.cin) errs.cin = 'CIN / Passeport requis';
    if (form.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(form.email)) errs.email = 'Email invalide';
    setErrors(errs);
    return Object.keys(errs).length === 0;
  }

  async function handleSubmit(e) {
    e.preventDefault();
    setSubmitError('');
    if (!validate()) return;
    setSaving(true);
    try {
      const payload = { ...form, licenseExpiry: form.licenseExpiry || null, email: form.email || null };
      let clientId = id;
      if (isEdit) {
        await api.put(`/clients/${id}`, payload);
      } else {
        const res = await api.post('/clients', payload);
        clientId = res.data.id;
      }
      if (files.length > 0) {
        const data = new FormData();
        files.forEach((f) => data.append('documents', f));
        await api.post(`/clients/${clientId}/documents`, data, { headers: { 'Content-Type': 'multipart/form-data' } });
      }
      navigate('/clients');
    } catch (err) {
      setSubmitError(getErrorMessage(err, "Impossible d'enregistrer le client."));
    } finally {
      setSaving(false);
    }
  }

  return (
    <div>
      <PageHeader title={isEdit ? 'Modifier le client' : 'Ajouter un client'} />

      <form onSubmit={handleSubmit} className="card space-y-6 p-6">
        {submitError && <div className="rounded-lg bg-red-50 px-4 py-2 text-sm text-red-600">{submitError}</div>}

        <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-3">
          <Field label="Prénom" error={errors.firstName}>
            <input className="input" value={form.firstName} onChange={(e) => update('firstName', e.target.value)} />
          </Field>
          <Field label="Nom" error={errors.lastName}>
            <input className="input" value={form.lastName} onChange={(e) => update('lastName', e.target.value)} />
          </Field>
          <Field label="Téléphone" error={errors.phone}>
            <input className="input" value={form.phone} onChange={(e) => update('phone', e.target.value)} />
          </Field>
          <Field label="Email" error={errors.email}>
            <input className="input" value={form.email} onChange={(e) => update('email', e.target.value)} />
          </Field>
          <Field label="Adresse">
            <input className="input" value={form.address} onChange={(e) => update('address', e.target.value)} />
          </Field>
          <Field label="CIN / Passeport" error={errors.cin}>
            <input className="input" value={form.cin} onChange={(e) => update('cin', e.target.value)} />
          </Field>
          <Field label="Numéro de permis">
            <input className="input" value={form.licenseNumber} onChange={(e) => update('licenseNumber', e.target.value)} />
          </Field>
          <Field label="Expiration du permis">
            <input type="date" className="input" value={form.licenseExpiry || ''} onChange={(e) => update('licenseExpiry', e.target.value)} />
          </Field>
        </div>

        <div>
          <label className="label">Copies des documents (CIN, permis...)</label>
          {documents.length > 0 && (
            <ul className="mb-3 space-y-1 text-sm">
              {documents.map((d, i) => (
                <li key={i}>
                  <a className="text-brand-600 hover:underline" href={`${uploadsBaseUrl}${d}`} target="_blank" rel="noreferrer">
                    Document {i + 1}
                  </a>
                </li>
              ))}
            </ul>
          )}
          <label className="flex w-fit cursor-pointer items-center gap-2 rounded-lg border border-dashed border-slate-300 px-4 py-2 text-sm text-slate-500 hover:bg-slate-50 dark:border-navy-600 dark:hover:bg-navy-900">
            <Upload size={16} />
            {files.length > 0 ? `${files.length} fichier(s) sélectionné(s)` : 'Ajouter des documents'}
            <input type="file" multiple className="hidden" onChange={(e) => setFiles(Array.from(e.target.files))} />
          </label>
        </div>

        <div className="flex justify-end gap-3 border-t border-slate-100 pt-4 dark:border-navy-700">
          <button type="button" className="btn-secondary" onClick={() => navigate('/clients')}>
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
