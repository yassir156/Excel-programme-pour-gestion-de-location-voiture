import { useEffect, useState } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import { ArrowLeft, Pencil } from 'lucide-react';
import { format } from 'date-fns';
import api, { uploadsBaseUrl } from '../api/client';
import PageHeader from '../components/PageHeader';
import StatusBadge from '../components/StatusBadge';

export default function VehicleDetails() {
  const { id } = useParams();
  const navigate = useNavigate();
  const [vehicle, setVehicle] = useState(null);
  const [history, setHistory] = useState({ reservations: [], maintenances: [] });
  const [activePhoto, setActivePhoto] = useState(0);

  useEffect(() => {
    api.get(`/vehicles/${id}`).then((res) => {
      setVehicle(res.data);
      setActivePhoto(0);
    });
    api.get(`/vehicles/${id}/history`).then((res) => setHistory(res.data));
  }, [id]);

  if (!vehicle) return <div className="text-slate-400">Chargement...</div>;

  return (
    <div>
      <PageHeader
        title={`${vehicle.brand} ${vehicle.model}`}
        subtitle={vehicle.plate}
        actions={
          <>
            <button className="btn-secondary" onClick={() => navigate('/vehicles')}>
              <ArrowLeft size={16} /> Retour
            </button>
            <button className="btn-primary" onClick={() => navigate(`/vehicles/${id}/edit`)}>
              <Pencil size={16} /> Modifier
            </button>
          </>
        }
      />

      <div className="grid grid-cols-1 gap-6 lg:grid-cols-3">
        <div className="card space-y-4 p-6 lg:col-span-1">
          {vehicle.photos?.length > 0 ? (
            <>
              <img
                src={`${uploadsBaseUrl}${vehicle.photos[activePhoto] || vehicle.photos[0]}`}
                alt=""
                className="h-48 w-full rounded-xl object-cover"
              />
              {vehicle.photos.length > 1 && (
                <div className="flex flex-wrap gap-2">
                  {vehicle.photos.map((p, i) => (
                    <button
                      key={p}
                      type="button"
                      onClick={() => setActivePhoto(i)}
                      className={`h-14 w-14 overflow-hidden rounded-lg border-2 ${
                        i === activePhoto ? 'border-brand-500' : 'border-transparent'
                      }`}
                    >
                      <img src={`${uploadsBaseUrl}${p}`} alt="" className="h-full w-full object-cover" />
                    </button>
                  ))}
                </div>
              )}
            </>
          ) : (
            <div className="flex h-48 items-center justify-center rounded-xl bg-slate-100 text-slate-400 dark:bg-navy-700">
              Aucune photo
            </div>
          )}
          <div className="flex justify-between text-sm">
            <span className="text-slate-400">Statut</span>
            <StatusBadge status={vehicle.status} />
          </div>
          <Info label="Catégorie" value={vehicle.category || '—'} />
          <Info label="Année" value={vehicle.year} />
          <Info label="Couleur" value={vehicle.color || '—'} />
          <Info label="Kilométrage" value={`${vehicle.mileage.toLocaleString()} km`} />
          <Info label="Carburant" value={vehicle.fuelType} />
          <Info label="Boîte" value={vehicle.transmission} />
          <Info label="Prix / jour" value={`${vehicle.dailyPrice} MAD`} />
          <Info label="Caution" value={`${vehicle.deposit} MAD`} />
          <Info label="Assurance" value={vehicle.insuranceExpiry ? format(new Date(vehicle.insuranceExpiry), 'dd/MM/yyyy') : '—'} />
          <Info label="Contrôle technique" value={vehicle.technicalControlExpiry ? format(new Date(vehicle.technicalControlExpiry), 'dd/MM/yyyy') : '—'} />
        </div>

        <div className="space-y-6 lg:col-span-2">
          <div className="card p-6">
            <h3 className="mb-4 text-sm font-semibold text-slate-700 dark:text-slate-200">Historique des locations</h3>
            {history.reservations.length === 0 && <p className="text-sm text-slate-400">Aucune location enregistrée.</p>}
            <div className="divide-y divide-slate-100 dark:divide-navy-700">
              {history.reservations.map((r) => (
                <div key={r.id} className="flex items-center justify-between py-3 text-sm">
                  <div>
                    <p className="font-medium text-slate-700 dark:text-slate-200">{r.client?.firstName} {r.client?.lastName}</p>
                    <p className="text-xs text-slate-400">
                      {format(new Date(r.startDate), 'dd/MM/yyyy')} → {format(new Date(r.endDate), 'dd/MM/yyyy')} · {r.totalPrice} MAD
                    </p>
                  </div>
                  <StatusBadge status={r.status} />
                </div>
              ))}
            </div>
          </div>

          <div className="card p-6">
            <h3 className="mb-4 text-sm font-semibold text-slate-700 dark:text-slate-200">Historique de maintenance</h3>
            {history.maintenances.length === 0 && <p className="text-sm text-slate-400">Aucune opération enregistrée.</p>}
            <div className="divide-y divide-slate-100 dark:divide-navy-700">
              {history.maintenances.map((m) => (
                <div key={m.id} className="flex items-center justify-between py-3 text-sm">
                  <div>
                    <p className="font-medium capitalize text-slate-700 dark:text-slate-200">{m.type.replace('_', ' ')}</p>
                    <p className="text-xs text-slate-400">{m.description || '—'}</p>
                  </div>
                  <div className="text-right">
                    <p className="text-slate-700 dark:text-slate-200">{m.cost} MAD</p>
                    <p className="text-xs text-slate-400">{format(new Date(m.date), 'dd/MM/yyyy')}</p>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function Info({ label, value }) {
  return (
    <div className="flex justify-between border-t border-slate-100 pt-2 text-sm dark:border-navy-700">
      <span className="text-slate-400">{label}</span>
      <span className="font-medium capitalize text-slate-700 dark:text-slate-200">{value}</span>
    </div>
  );
}
