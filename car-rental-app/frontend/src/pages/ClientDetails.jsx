import { useEffect, useState } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import { ArrowLeft, Pencil } from 'lucide-react';
import { format } from 'date-fns';
import api, { uploadsBaseUrl } from '../api/client';
import PageHeader from '../components/PageHeader';
import StatusBadge from '../components/StatusBadge';

export default function ClientDetails() {
  const { id } = useParams();
  const navigate = useNavigate();
  const [client, setClient] = useState(null);
  const [reservations, setReservations] = useState([]);

  useEffect(() => {
    api.get(`/clients/${id}`).then((res) => setClient(res.data));
    api.get(`/clients/${id}/history`).then((res) => setReservations(res.data));
  }, [id]);

  if (!client) return <div className="text-slate-400">Chargement...</div>;

  return (
    <div>
      <PageHeader
        title={`${client.firstName} ${client.lastName}`}
        subtitle={client.phone}
        actions={
          <>
            <button className="btn-secondary" onClick={() => navigate('/clients')}>
              <ArrowLeft size={16} /> Retour
            </button>
            <button className="btn-primary" onClick={() => navigate(`/clients/${id}/edit`)}>
              <Pencil size={16} /> Modifier
            </button>
          </>
        }
      />

      <div className="grid grid-cols-1 gap-6 lg:grid-cols-3">
        <div className="card space-y-3 p-6 lg:col-span-1">
          <Info label="Email" value={client.email || '—'} />
          <Info label="Adresse" value={client.address || '—'} />
          <Info label="CIN / Passeport" value={client.cin} />
          <Info label="Numéro de permis" value={client.licenseNumber || '—'} />
          <Info label="Expiration permis" value={client.licenseExpiry ? format(new Date(client.licenseExpiry), 'dd/MM/yyyy') : '—'} />
          {client.documents?.length > 0 && (
            <div>
              <p className="label mt-3">Documents</p>
              <ul className="space-y-1 text-sm">
                {client.documents.map((d, i) => (
                  <li key={i}>
                    <a className="text-brand-600 hover:underline" href={`${uploadsBaseUrl}${d}`} target="_blank" rel="noreferrer">
                      Document {i + 1}
                    </a>
                  </li>
                ))}
              </ul>
            </div>
          )}
        </div>

        <div className="card p-6 lg:col-span-2">
          <h3 className="mb-4 text-sm font-semibold text-slate-700 dark:text-slate-200">Historique des locations</h3>
          {reservations.length === 0 && <p className="text-sm text-slate-400">Aucune location enregistrée.</p>}
          <div className="divide-y divide-slate-100 dark:divide-navy-700">
            {reservations.map((r) => (
              <div key={r.id} className="flex items-center justify-between py-3 text-sm">
                <div>
                  <p className="font-medium text-slate-700 dark:text-slate-200">{r.vehicle?.brand} {r.vehicle?.model}</p>
                  <p className="text-xs text-slate-400">
                    {format(new Date(r.startDate), 'dd/MM/yyyy')} → {format(new Date(r.endDate), 'dd/MM/yyyy')} · {r.totalPrice} MAD
                  </p>
                </div>
                <StatusBadge status={r.status} />
              </div>
            ))}
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
      <span className="font-medium text-slate-700 dark:text-slate-200">{value}</span>
    </div>
  );
}
