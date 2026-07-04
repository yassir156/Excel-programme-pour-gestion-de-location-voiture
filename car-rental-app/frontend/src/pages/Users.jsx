import { useEffect, useState } from 'react';
import { Plus, Pencil, Trash2, ShieldOff } from 'lucide-react';
import api, { getErrorMessage } from '../api/client';
import { useAuth } from '../context/AuthContext';
import DataTable from '../components/DataTable';
import PageHeader from '../components/PageHeader';
import Modal from '../components/Modal';
import ConfirmDialog from '../components/ConfirmDialog';

const ROLE_LABELS = { administrateur: 'Administrateur', manager: 'Manager', agent: 'Agent' };
const emptyForm = { username: '', password: '', fullName: '', role: 'agent', active: true };

export default function Users() {
  const { user: currentUser } = useAuth();
  const [users, setUsers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [modalOpen, setModalOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [form, setForm] = useState(emptyForm);
  const [formError, setFormError] = useState('');
  const [saving, setSaving] = useState(false);
  const [toDelete, setToDelete] = useState(null);

  function load() {
    setLoading(true);
    api.get('/users').then((res) => setUsers(res.data)).catch((err) => setError(getErrorMessage(err))).finally(() => setLoading(false));
  }

  useEffect(load, []);

  if (currentUser?.role !== 'administrateur') {
    return (
      <div className="card flex flex-col items-center gap-3 p-10 text-center text-slate-500">
        <ShieldOff size={32} />
        Seul un administrateur peut gérer les comptes utilisateurs.
      </div>
    );
  }

  function openCreate() {
    setEditingId(null);
    setForm(emptyForm);
    setFormError('');
    setModalOpen(true);
  }

  function openEdit(u) {
    setEditingId(u.id);
    setForm({ username: u.username, password: '', fullName: u.fullName, role: u.role, active: u.active });
    setFormError('');
    setModalOpen(true);
  }

  async function handleSubmit(e) {
    e.preventDefault();
    setFormError('');
    if (!form.fullName || (!editingId && (!form.username || !form.password))) {
      setFormError('Nom complet requis (identifiant et mot de passe requis à la création).');
      return;
    }
    setSaving(true);
    try {
      if (editingId) {
        const payload = { fullName: form.fullName, role: form.role, active: form.active };
        if (form.password) payload.password = form.password;
        await api.put(`/users/${editingId}`, payload);
      } else {
        await api.post('/users', form);
      }
      setModalOpen(false);
      load();
    } catch (err) {
      setFormError(getErrorMessage(err));
    } finally {
      setSaving(false);
    }
  }

  async function confirmDelete() {
    try {
      await api.delete(`/users/${toDelete.id}`);
      setToDelete(null);
      load();
    } catch (err) {
      setError(getErrorMessage(err));
      setToDelete(null);
    }
  }

  const columns = [
    { key: 'fullName', header: 'Nom complet' },
    { key: 'username', header: 'Identifiant' },
    { key: 'role', header: 'Rôle', render: (r) => ROLE_LABELS[r.role] || r.role },
    {
      key: 'active',
      header: 'Statut',
      render: (r) => (
        <span className={`badge ${r.active ? 'bg-emerald-50 text-emerald-700' : 'bg-slate-100 text-slate-500'}`}>
          {r.active ? 'Actif' : 'Désactivé'}
        </span>
      ),
    },
    {
      key: 'actions',
      header: 'Actions',
      render: (r) => (
        <div className="flex gap-2">
          <button className="btn-secondary px-2 py-1" onClick={() => openEdit(r)}>
            <Pencil size={15} />
          </button>
          {r.id !== currentUser.id && (
            <button className="btn-danger px-2 py-1" onClick={() => setToDelete(r)}>
              <Trash2 size={15} />
            </button>
          )}
        </div>
      ),
    },
  ];

  return (
    <div>
      <PageHeader
        title="Gestion des utilisateurs"
        subtitle={`${users.length} compte(s)`}
        actions={
          <button className="btn-primary" onClick={openCreate}>
            <Plus size={16} /> Ajouter un utilisateur
          </button>
        }
      />

      {error && <div className="mb-4 rounded-lg bg-red-50 px-4 py-2 text-sm text-red-600">{error}</div>}

      <DataTable columns={columns} data={users} searchPlaceholder="Rechercher..." emptyMessage={loading ? 'Chargement...' : 'Aucun utilisateur.'} />

      <Modal open={modalOpen} title={editingId ? "Modifier l'utilisateur" : 'Ajouter un utilisateur'} onClose={() => setModalOpen(false)}>
        <form onSubmit={handleSubmit} className="space-y-4">
          {formError && <div className="rounded-lg bg-red-50 px-4 py-2 text-sm text-red-600">{formError}</div>}

          <div>
            <label className="label">Nom complet</label>
            <input className="input" value={form.fullName} onChange={(e) => setForm((f) => ({ ...f, fullName: e.target.value }))} />
          </div>

          {!editingId && (
            <div>
              <label className="label">Identifiant</label>
              <input className="input" value={form.username} onChange={(e) => setForm((f) => ({ ...f, username: e.target.value }))} />
            </div>
          )}

          <div>
            <label className="label">{editingId ? 'Nouveau mot de passe (optionnel)' : 'Mot de passe'}</label>
            <input type="password" className="input" value={form.password} onChange={(e) => setForm((f) => ({ ...f, password: e.target.value }))} />
          </div>

          <div>
            <label className="label">Rôle</label>
            <select className="input" value={form.role} onChange={(e) => setForm((f) => ({ ...f, role: e.target.value }))}>
              <option value="administrateur">Administrateur</option>
              <option value="manager">Manager</option>
              <option value="agent">Agent</option>
            </select>
          </div>

          {editingId && (
            <label className="flex items-center gap-2 text-sm text-slate-600 dark:text-slate-300">
              <input type="checkbox" checked={form.active} onChange={(e) => setForm((f) => ({ ...f, active: e.target.checked }))} />
              Compte actif
            </label>
          )}

          <div className="flex justify-end gap-3 pt-2">
            <button type="button" className="btn-secondary" onClick={() => setModalOpen(false)}>Annuler</button>
            <button type="submit" className="btn-primary" disabled={saving}>{saving ? 'Enregistrement...' : 'Enregistrer'}</button>
          </div>
        </form>
      </Modal>

      <ConfirmDialog
        open={!!toDelete}
        message={`Supprimer l'utilisateur ${toDelete?.fullName} ?`}
        onConfirm={confirmDelete}
        onCancel={() => setToDelete(null)}
      />
    </div>
  );
}
