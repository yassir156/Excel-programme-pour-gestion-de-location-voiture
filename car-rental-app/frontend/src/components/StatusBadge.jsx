const STYLES = {
  // vehicle statuses
  disponible: 'bg-emerald-50 text-emerald-700 dark:bg-emerald-950 dark:text-emerald-300',
  louee: 'bg-brand-50 text-brand-700 dark:bg-navy-700 dark:text-brand-200',
  maintenance: 'bg-amber-50 text-amber-700 dark:bg-amber-950 dark:text-amber-300',
  reservee: 'bg-purple-50 text-purple-700 dark:bg-purple-950 dark:text-purple-300',
  // reservation statuses
  en_attente: 'bg-amber-50 text-amber-700 dark:bg-amber-950 dark:text-amber-300',
  confirmee: 'bg-brand-50 text-brand-700 dark:bg-navy-700 dark:text-brand-200',
  annulee: 'bg-red-50 text-red-700 dark:bg-red-950 dark:text-red-300',
  terminee: 'bg-emerald-50 text-emerald-700 dark:bg-emerald-950 dark:text-emerald-300',
  // payment statuses
  paye: 'bg-emerald-50 text-emerald-700 dark:bg-emerald-950 dark:text-emerald-300',
  partiel: 'bg-amber-50 text-amber-700 dark:bg-amber-950 dark:text-amber-300',
  impaye: 'bg-red-50 text-red-700 dark:bg-red-950 dark:text-red-300',
};

const LABELS = {
  disponible: 'Disponible',
  louee: 'Louée',
  maintenance: 'Maintenance',
  reservee: 'Réservée',
  en_attente: 'En attente',
  confirmee: 'Confirmée',
  annulee: 'Annulée',
  terminee: 'Terminée',
  paye: 'Payé',
  partiel: 'Partiel',
  impaye: 'Impayé',
};

export default function StatusBadge({ status }) {
  return (
    <span className={`badge ${STYLES[status] || 'bg-slate-100 text-slate-600'}`}>
      {LABELS[status] || status}
    </span>
  );
}
