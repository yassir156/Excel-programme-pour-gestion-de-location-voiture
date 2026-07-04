import { useEffect, useRef, useState } from 'react';
import { Save, Upload, Database, UploadCloud } from 'lucide-react';
import api, { getErrorMessage, uploadsBaseUrl } from '../api/client';
import { useAuth } from '../context/AuthContext';
import PageHeader from '../components/PageHeader';

export default function Settings() {
  const { user } = useAuth();
  const [form, setForm] = useState(null);
  const [saving, setSaving] = useState(false);
  const [message, setMessage] = useState('');
  const [error, setError] = useState('');
  const logoInput = useRef(null);
  const restoreInput = useRef(null);

  const canEdit = ['administrateur', 'manager'].includes(user?.role);
  const isAdmin = user?.role === 'administrateur';

  useEffect(() => {
    api.get('/settings').then((res) => setForm(res.data));
  }, []);

  function update(field, value) {
    setForm((f) => ({ ...f, [field]: value }));
  }

  async function handleSubmit(e) {
    e.preventDefault();
    setMessage('');
    setError('');
    setSaving(true);
    try {
      const res = await api.put('/settings', {
        name: form.name, address: form.address, phone: form.phone, email: form.email,
        currency: form.currency, taxRate: Number(form.taxRate) || 0, contractTerms: form.contractTerms,
      });
      setForm(res.data);
      setMessage('Paramètres enregistrés avec succès.');
    } catch (err) {
      setError(getErrorMessage(err));
    } finally {
      setSaving(false);
    }
  }

  async function handleLogoChange(e) {
    const file = e.target.files[0];
    if (!file) return;
    const data = new FormData();
    data.append('logo', file);
    try {
      const res = await api.post('/settings/logo', data, { headers: { 'Content-Type': 'multipart/form-data' } });
      setForm(res.data);
    } catch (err) {
      setError(getErrorMessage(err));
    }
  }

  async function handleBackup() {
    try {
      const res = await api.get('/settings/backup', { responseType: 'blob' });
      const url = window.URL.createObjectURL(new Blob([res.data]));
      const link = document.createElement('a');
      link.href = url;
      link.download = `sauvegarde-car-rental-${Date.now()}.sqlite`;
      link.click();
    } catch (err) {
      setError(getErrorMessage(err, 'Impossible de générer la sauvegarde.'));
    }
  }

  async function handleRestore(e) {
    const file = e.target.files[0];
    if (!file) return;
    const data = new FormData();
    data.append('database', file);
    try {
      const res = await api.post('/settings/restore', data, { headers: { 'Content-Type': 'multipart/form-data' } });
      setMessage(res.data.message);
    } catch (err) {
      setError(getErrorMessage(err));
    }
  }

  if (!form) return <div className="text-slate-400">Chargement...</div>;

  return (
    <div className="space-y-6">
      <PageHeader title="Paramètres de l'agence" />

      {message && <div className="rounded-lg bg-emerald-50 px-4 py-2 text-sm text-emerald-700">{message}</div>}
      {error && <div className="rounded-lg bg-red-50 px-4 py-2 text-sm text-red-600">{error}</div>}

      <form onSubmit={handleSubmit} className="card space-y-6 p-6">
        <div className="flex items-center gap-4">
          <div className="flex h-16 w-16 items-center justify-center overflow-hidden rounded-xl bg-slate-100 dark:bg-navy-700">
            {form.logo ? <img src={`${uploadsBaseUrl}${form.logo}`} alt="Logo" className="h-full w-full object-cover" /> : <span className="text-xs text-slate-400">Logo</span>}
          </div>
          <button type="button" className="btn-secondary" onClick={() => logoInput.current?.click()} disabled={!canEdit}>
            <Upload size={16} /> Changer le logo
          </button>
          <input ref={logoInput} type="file" accept="image/*" className="hidden" onChange={handleLogoChange} />
        </div>

        <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
          <Field label="Nom de l'agence">
            <input className="input" disabled={!canEdit} value={form.name || ''} onChange={(e) => update('name', e.target.value)} />
          </Field>
          <Field label="Téléphone">
            <input className="input" disabled={!canEdit} value={form.phone || ''} onChange={(e) => update('phone', e.target.value)} />
          </Field>
          <Field label="Email">
            <input className="input" disabled={!canEdit} value={form.email || ''} onChange={(e) => update('email', e.target.value)} />
          </Field>
          <Field label="Adresse">
            <input className="input" disabled={!canEdit} value={form.address || ''} onChange={(e) => update('address', e.target.value)} />
          </Field>
          <Field label="Devise">
            <input className="input" disabled={!canEdit} value={form.currency || ''} onChange={(e) => update('currency', e.target.value)} />
          </Field>
          <Field label="Taux de taxe (%)">
            <input type="number" className="input" disabled={!canEdit} value={form.taxRate || 0} onChange={(e) => update('taxRate', e.target.value)} />
          </Field>
        </div>

        <Field label="Conditions générales du contrat">
          <textarea className="input" rows={6} disabled={!canEdit} value={form.contractTerms || ''} onChange={(e) => update('contractTerms', e.target.value)} />
        </Field>

        {canEdit && (
          <div className="flex justify-end border-t border-slate-100 pt-4 dark:border-navy-700">
            <button type="submit" className="btn-primary" disabled={saving}>
              <Save size={16} /> {saving ? 'Enregistrement...' : 'Enregistrer'}
            </button>
          </div>
        )}
      </form>

      {isAdmin && (
        <div className="card space-y-4 p-6">
          <h3 className="text-sm font-semibold text-slate-700 dark:text-slate-200">Sauvegarde et restauration</h3>
          <p className="text-sm text-slate-400">
            Téléchargez une copie complète de la base de données ou restaurez-la à partir d'un fichier de sauvegarde.
          </p>
          <div className="flex flex-wrap gap-3">
            <button className="btn-secondary" onClick={handleBackup}>
              <Database size={16} /> Télécharger la sauvegarde
            </button>
            <button className="btn-secondary" onClick={() => restoreInput.current?.click()}>
              <UploadCloud size={16} /> Restaurer une sauvegarde
            </button>
            <input ref={restoreInput} type="file" accept=".sqlite" className="hidden" onChange={handleRestore} />
          </div>
        </div>
      )}
    </div>
  );
}

function Field({ label, children }) {
  return (
    <div>
      <label className="label">{label}</label>
      {children}
    </div>
  );
}
