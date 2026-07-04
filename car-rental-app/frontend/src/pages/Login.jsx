import { useState } from 'react';
import { useNavigate, Navigate } from 'react-router-dom';
import { Lock, User, Eye, EyeOff } from 'lucide-react';
import { useAuth } from '../context/AuthContext';
import { getErrorMessage } from '../api/client';
import logo from '../assets/logo.png';

export default function Login() {
  const { user, login } = useAuth();
  const navigate = useNavigate();
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);

  if (user) return <Navigate to="/" replace />;

  async function handleSubmit(e) {
    e.preventDefault();
    setError('');
    if (!username || !password) {
      setError("Veuillez saisir l'identifiant et le mot de passe.");
      return;
    }
    setLoading(true);
    try {
      await login(username, password);
      navigate('/');
    } catch (err) {
      setError(getErrorMessage(err, 'Identifiants incorrects.'));
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="flex min-h-screen items-center justify-center bg-gradient-to-br from-navy-950 via-navy-900 to-brand-900 p-4">
      <div className="grid w-full max-w-4xl overflow-hidden rounded-2xl bg-white shadow-2xl dark:bg-navy-800 md:grid-cols-2">
        <div className="hidden flex-col justify-between bg-gradient-to-br from-brand-700 to-navy-900 p-10 text-white md:flex">
          <div className="flex items-center gap-3">
            <img src={logo} alt="Nova Motion Car" className="h-12 w-12 rounded-full object-cover shadow-lg" />
            <span className="text-lg font-bold tracking-wide">NOVA MOTION CAR</span>
          </div>
          <div>
            <h2 className="text-2xl font-bold leading-snug">
              Gérez votre flotte, vos clients et vos réservations en toute simplicité.
            </h2>
            <p className="mt-3 text-sm text-white/70">
              Drive with confidence — la plateforme complète pour piloter votre agence de location de voitures.
            </p>
          </div>
          <p className="text-xs text-white/50">© {new Date().getFullYear()} Nova Motion Car</p>
        </div>

        <div className="p-8 sm:p-10">
          <h1 className="text-2xl font-bold text-slate-800 dark:text-white">Bienvenue</h1>
          <p className="mt-1 text-sm text-slate-400">Connectez-vous pour accéder à votre espace.</p>

          <form onSubmit={handleSubmit} className="mt-8 space-y-4">
            {error && (
              <div className="rounded-lg bg-red-50 px-4 py-2.5 text-sm text-red-600 dark:bg-red-950 dark:text-red-300">
                {error}
              </div>
            )}

            <div>
              <label className="label">Identifiant</label>
              <div className="relative">
                <User size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                <input
                  className="input pl-9"
                  value={username}
                  onChange={(e) => setUsername(e.target.value)}
                  placeholder="admin"
                  autoFocus
                />
              </div>
            </div>

            <div>
              <label className="label">Mot de passe</label>
              <div className="relative">
                <Lock size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                <input
                  type={showPassword ? 'text' : 'password'}
                  className="input pl-9 pr-9"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  placeholder="••••••••"
                />
                <button
                  type="button"
                  className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400"
                  onClick={() => setShowPassword((s) => !s)}
                >
                  {showPassword ? <EyeOff size={16} /> : <Eye size={16} />}
                </button>
              </div>
            </div>

            <button type="submit" className="btn-primary w-full py-2.5" disabled={loading}>
              {loading ? 'Connexion...' : 'Se connecter'}
            </button>
          </form>

          <div className="mt-6 rounded-lg bg-slate-50 p-3 text-xs text-slate-500 dark:bg-navy-900 dark:text-slate-400">
            Comptes de démonstration : <b>admin</b> / admin123, <b>manager</b> / manager123, <b>agent</b> / agent123
          </div>
        </div>
      </div>
    </div>
  );
}
