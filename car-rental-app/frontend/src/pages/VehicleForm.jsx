import { useEffect, useMemo, useState } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import { Save, Upload, X, Trash2 } from 'lucide-react';
import api, { getErrorMessage, uploadsBaseUrl } from '../api/client';
import PageHeader from '../components/PageHeader';

const emptyForm = {
  brand: '', model: '', year: new Date().getFullYear(), plate: '', category: '', color: '',
  mileage: 0, fuelType: 'essence', transmission: 'manuelle', dailyPrice: '', deposit: '',
  status: 'disponible', insuranceExpiry: '', technicalControlExpiry: '',
};

export default function VehicleForm() {
  const { id } = useParams();
  const isEdit = !!id;
  const navigate = useNavigate();
  const [form, setForm] = useState(emptyForm);
  const [photos, setPhotos] = useState([]);
  const [files, setFiles] = useState([]);
  const [errors, setErrors] = useState({});
  const [submitError, setSubmitError] = useState('');
  const [saving, setSaving] = useState(false);
  const [deletingPhoto, setDeletingPhoto] = useState(null);

  const filePreviews = useMemo(() => files.map((f) => URL.createObjectURL(f)), [files]);

  useEffect(() => {
    if (!isEdit) return;
    api.get(`/vehicles/${id}`).then((res) => {
      const v = res.data;
      setForm({
        brand: v.brand, model: v.model, year: v.year, plate: v.plate, category: v.category || '',
        color: v.color || '', mileage: v.mileage, fuelType: v.fuelType, transmission: v.transmission,
        dailyPrice: v.dailyPrice, deposit: v.deposit, status: v.status,
        insuranceExpiry: v.insuranceExpiry || '', technicalControlExpiry: v.technicalControlExpiry || '',
      });
      setPhotos(v.photos || []);
    });
  }, [id, isEdit]);

  function update(field, value) {
    setForm((f) => ({ ...f, [field]: value }));
  }

  function validate() {
    const errs = {};
    if (!form.brand) errs.brand = 'Marque requise';
    if (!form.model) errs.model = 'Modèle requis';
    if (!form.plate) errs.plate = 'Immatriculation requise';
    if (!form.year || form.year < 1980) errs.year = 'Année invalide';
    if (form.dailyPrice === '' || Number(form.dailyPrice) <= 0) errs.dailyPrice = 'Prix par jour invalide';
    if (form.deposit === '' || Number(form.deposit) < 0) errs.deposit = 'Caution invalide';
    setErrors(errs);
    return Object.keys(errs).length === 0;
  }

  async function handleSubmit(e) {
    e.preventDefault();
    setSubmitError('');
    if (!validate()) return;
    setSaving(true);
    try {
      const payload = {
        ...form,
        year: Number(form.year),
        mileage: Number(form.mileage) || 0,
        dailyPrice: Number(form.dailyPrice),
        deposit: Number(form.deposit),
        insuranceExpiry: form.insuranceExpiry || null,
        technicalControlExpiry: form.technicalControlExpiry || null,
      };
      let vehicleId = id;
      if (isEdit) {
        await api.put(`/vehicles/${id}`, payload);
      } else {
        const res = await api.post('/vehicles', payload);
        vehicleId = res.data.id;
      }
      if (files.length > 0) {
        const data = new FormData();
        files.forEach((f) => data.append('photos', f));
        await api.post(`/vehicles/${vehicleId}/photos`, data, { headers: { 'Content-Type': 'multipart/form-data' } });
      }
      navigate('/vehicles');
    } catch (err) {
      setSubmitError(getErrorMessage(err, "Impossible d'enregistrer le véhicule."));
    } finally {
      setSaving(false);
    }
  }

  function removePendingFile(index) {
    setFiles((prev) => prev.filter((_, i) => i !== index));
  }

  async function deleteUploadedPhoto(photo) {
    setDeletingPhoto(photo);
    try {
      const res = await api.delete(`/vehicles/${id}/photos`, { data: { photo } });
      setPhotos(res.data.photos || []);
    } catch (err) {
      setSubmitError(getErrorMessage(err, 'Impossible de supprimer cette photo.'));
    } finally {
      setDeletingPhoto(null);
    }
  }

  return (
    <div>
      <PageHeader title={isEdit ? 'Modifier le véhicule' : 'Ajouter un véhicule'} />

      <form onSubmit={handleSubmit} className="card space-y-6 p-6">
        {submitError && <div className="rounded-lg bg-red-50 px-4 py-2 text-sm text-red-600">{submitError}</div>}

        <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-3">
          <Field label="Marque" error={errors.brand}>
            <input className="input" value={form.brand} onChange={(e) => update('brand', e.target.value)} />
          </Field>
          <Field label="Modèle" error={errors.model}>
            <input className="input" value={form.model} onChange={(e) => update('model', e.target.value)} />
          </Field>
          <Field label="Année" error={errors.year}>
            <input type="number" className="input" value={form.year} onChange={(e) => update('year', e.target.value)} />
          </Field>
          <Field label="Immatriculation" error={errors.plate}>
            <input className="input" value={form.plate} onChange={(e) => update('plate', e.target.value)} />
          </Field>
          <Field label="Catégorie">
            <input className="input" placeholder="SUV, Berline..." value={form.category} onChange={(e) => update('category', e.target.value)} />
          </Field>
          <Field label="Couleur">
            <input className="input" value={form.color} onChange={(e) => update('color', e.target.value)} />
          </Field>
          <Field label="Kilométrage">
            <input type="number" className="input" value={form.mileage} onChange={(e) => update('mileage', e.target.value)} />
          </Field>
          <Field label="Carburant">
            <select className="input" value={form.fuelType} onChange={(e) => update('fuelType', e.target.value)}>
              <option value="essence">Essence</option>
              <option value="diesel">Diesel</option>
              <option value="hybride">Hybride</option>
              <option value="electrique">Électrique</option>
            </select>
          </Field>
          <Field label="Boîte de vitesses">
            <select className="input" value={form.transmission} onChange={(e) => update('transmission', e.target.value)}>
              <option value="manuelle">Manuelle</option>
              <option value="automatique">Automatique</option>
            </select>
          </Field>
          <Field label="Prix par jour (MAD)" error={errors.dailyPrice}>
            <input type="number" className="input" value={form.dailyPrice} onChange={(e) => update('dailyPrice', e.target.value)} />
          </Field>
          <Field label="Caution (MAD)" error={errors.deposit}>
            <input type="number" className="input" value={form.deposit} onChange={(e) => update('deposit', e.target.value)} />
          </Field>
          <Field label="Statut">
            <select className="input" value={form.status} onChange={(e) => update('status', e.target.value)}>
              <option value="disponible">Disponible</option>
              <option value="louee">Louée</option>
              <option value="maintenance">Maintenance</option>
              <option value="reservee">Réservée</option>
            </select>
          </Field>
          <Field label="Expiration assurance">
            <input type="date" className="input" value={form.insuranceExpiry || ''} onChange={(e) => update('insuranceExpiry', e.target.value)} />
          </Field>
          <Field label="Expiration contrôle technique">
            <input type="date" className="input" value={form.technicalControlExpiry || ''} onChange={(e) => update('technicalControlExpiry', e.target.value)} />
          </Field>
        </div>

        <div>
          <label className="label">Photos du véhicule</label>

          {photos.length > 0 && (
            <div className="mb-3 flex flex-wrap gap-3">
              {photos.map((p, i) => (
                <div key={i} className="group relative h-20 w-20">
                  <img src={`${uploadsBaseUrl}${p}`} alt="" className="h-20 w-20 rounded-lg object-cover" />
                  <button
                    type="button"
                    onClick={() => deleteUploadedPhoto(p)}
                    disabled={deletingPhoto === p}
                    className="absolute -right-2 -top-2 flex h-6 w-6 items-center justify-center rounded-full bg-red-600 text-white opacity-0 shadow transition-opacity group-hover:opacity-100 disabled:opacity-70"
                    title="Supprimer cette photo"
                  >
                    <Trash2 size={12} />
                  </button>
                </div>
              ))}
            </div>
          )}

          {files.length > 0 && (
            <div className="mb-3 flex flex-wrap gap-3">
              {files.map((f, i) => (
                <div key={i} className="group relative h-20 w-20">
                  <img src={filePreviews[i]} alt="" className="h-20 w-20 rounded-lg border-2 border-dashed border-brand-300 object-cover" />
                  <button
                    type="button"
                    onClick={() => removePendingFile(i)}
                    className="absolute -right-2 -top-2 flex h-6 w-6 items-center justify-center rounded-full bg-red-600 text-white opacity-0 shadow transition-opacity group-hover:opacity-100"
                    title="Retirer cette photo"
                  >
                    <X size={12} />
                  </button>
                </div>
              ))}
            </div>
          )}

          <label className="flex w-fit cursor-pointer items-center gap-2 rounded-lg border border-dashed border-slate-300 px-4 py-2 text-sm text-slate-500 hover:bg-slate-50 dark:border-navy-600 dark:hover:bg-navy-900">
            <Upload size={16} />
            {isEdit ? 'Ajouter des photos' : 'Ajouter des photos (envoyées après création)'}
            <input
              type="file"
              multiple
              accept="image/*"
              className="hidden"
              onChange={(e) => setFiles((prev) => [...prev, ...Array.from(e.target.files)])}
            />
          </label>
          <p className="mt-1 text-xs text-slate-400">Formats acceptés : JPG, PNG. 8 photos maximum par envoi.</p>
        </div>

        <div className="flex justify-end gap-3 border-t border-slate-100 pt-4 dark:border-navy-700">
          <button type="button" className="btn-secondary" onClick={() => navigate('/vehicles')}>
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
