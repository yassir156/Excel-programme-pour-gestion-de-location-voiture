import { useEffect, useState } from 'react';
import { format } from 'date-fns';
import { FileDown, CheckSquare, Square } from 'lucide-react';
import api, { getErrorMessage } from '../api/client';
import DataTable from '../components/DataTable';
import PageHeader from '../components/PageHeader';
import { generateContractPDF } from '../utils/pdf';

export default function Contracts() {
  const [contracts, setContracts] = useState([]);
  const [settings, setSettings] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');

  function load() {
    setLoading(true);
    Promise.all([api.get('/contracts'), api.get('/settings')])
      .then(([c, s]) => {
        setContracts(c.data);
        setSettings(s.data);
      })
      .catch((err) => setError(getErrorMessage(err)))
      .finally(() => setLoading(false));
  }

  useEffect(load, []);

  async function toggleSign(contract, field) {
    try {
      const res = await api.put(`/contracts/${contract.id}/sign`, { [field]: !contract[field] });
      setContracts((prev) => prev.map((c) => (c.id === contract.id ? { ...c, ...res.data } : c)));
    } catch (err) {
      setError(getErrorMessage(err));
    }
  }

  const columns = [
    { key: 'contractNumber', header: 'N° Contrat' },
    {
      key: 'client',
      header: 'Client',
      searchValue: (r) => `${r.client?.firstName} ${r.client?.lastName}`,
      render: (r) => `${r.client?.firstName} ${r.client?.lastName}`,
    },
    {
      key: 'vehicle',
      header: 'Véhicule',
      searchValue: (r) => `${r.vehicle?.brand} ${r.vehicle?.model}`,
      render: (r) => `${r.vehicle?.brand} ${r.vehicle?.model} (${r.vehicle?.plate})`,
    },
    {
      key: 'dates',
      header: 'Période',
      render: (r) => `${format(new Date(r.startDate), 'dd/MM/yy')} → ${format(new Date(r.endDate), 'dd/MM/yy')}`,
    },
    { key: 'totalPrice', header: 'Total', render: (r) => `${r.totalPrice} MAD` },
    {
      key: 'signatures',
      header: 'Signatures',
      render: (r) => (
        <div className="flex flex-col gap-1 text-xs">
          <button className="flex items-center gap-1" onClick={() => toggleSign(r, 'signedClient')}>
            {r.signedClient ? <CheckSquare size={14} className="text-emerald-600" /> : <Square size={14} className="text-slate-400" />}
            Client
          </button>
          <button className="flex items-center gap-1" onClick={() => toggleSign(r, 'signedAgency')}>
            {r.signedAgency ? <CheckSquare size={14} className="text-emerald-600" /> : <Square size={14} className="text-slate-400" />}
            Agence
          </button>
        </div>
      ),
    },
    {
      key: 'actions',
      header: 'Actions',
      render: (r) => (
        <button className="btn-secondary px-2 py-1" onClick={() => generateContractPDF(r, settings)} title="Télécharger le PDF">
          <FileDown size={15} />
        </button>
      ),
    },
  ];

  return (
    <div>
      <PageHeader title="Contrats de location" subtitle={`${contracts.length} contrat(s)`} />

      {error && <div className="mb-4 rounded-lg bg-red-50 px-4 py-2 text-sm text-red-600">{error}</div>}

      <DataTable
        columns={columns}
        data={contracts}
        searchPlaceholder="Rechercher par client, véhicule..."
        emptyMessage={loading ? 'Chargement...' : 'Aucun contrat généré. Générez-en un depuis une réservation.'}
      />
    </div>
  );
}
