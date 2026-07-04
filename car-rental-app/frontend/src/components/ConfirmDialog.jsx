import Modal from './Modal';
import { AlertTriangle } from 'lucide-react';

export default function ConfirmDialog({ open, title = 'Confirmation', message, onConfirm, onCancel, danger = true }) {
  return (
    <Modal open={open} title={title} onClose={onCancel} size="sm">
      <div className="flex items-start gap-3">
        <div className={`flex h-10 w-10 shrink-0 items-center justify-center rounded-full ${danger ? 'bg-red-50 text-red-600 dark:bg-red-950' : 'bg-amber-50 text-amber-600 dark:bg-amber-950'}`}>
          <AlertTriangle size={20} />
        </div>
        <p className="text-sm text-slate-600 dark:text-slate-300">{message}</p>
      </div>
      <div className="mt-6 flex justify-end gap-3">
        <button className="btn-secondary" onClick={onCancel}>Annuler</button>
        <button className={danger ? 'btn-danger' : 'btn-primary'} onClick={onConfirm}>Confirmer</button>
      </div>
    </Modal>
  );
}
