import { Navigate, Route, Routes } from 'react-router-dom';
import { useAuth } from './context/AuthContext';
import AppLayout from './components/AppLayout';
import ProtectedRoute from './components/ProtectedRoute';

import Login from './pages/Login';
import Dashboard from './pages/Dashboard';
import Vehicles from './pages/Vehicles';
import VehicleForm from './pages/VehicleForm';
import VehicleDetails from './pages/VehicleDetails';
import Clients from './pages/Clients';
import ClientForm from './pages/ClientForm';
import ClientDetails from './pages/ClientDetails';
import Reservations from './pages/Reservations';
import ReservationForm from './pages/ReservationForm';
import Contracts from './pages/Contracts';
import Payments from './pages/Payments';
import Returns from './pages/Returns';
import Maintenance from './pages/Maintenance';
import Statistics from './pages/Statistics';
import Settings from './pages/Settings';
import Users from './pages/Users';

export default function App() {
  const { loading } = useAuth();

  if (loading) {
    return (
      <div className="flex h-screen items-center justify-center bg-slate-50 dark:bg-navy-950">
        <div className="h-10 w-10 animate-spin rounded-full border-4 border-brand-200 border-t-brand-600" />
      </div>
    );
  }

  return (
    <Routes>
      <Route path="/login" element={<Login />} />
      <Route
        element={
          <ProtectedRoute>
            <AppLayout />
          </ProtectedRoute>
        }
      >
        <Route path="/" element={<Dashboard />} />
        <Route path="/vehicles" element={<Vehicles />} />
        <Route path="/vehicles/new" element={<VehicleForm />} />
        <Route path="/vehicles/:id/edit" element={<VehicleForm />} />
        <Route path="/vehicles/:id" element={<VehicleDetails />} />
        <Route path="/clients" element={<Clients />} />
        <Route path="/clients/new" element={<ClientForm />} />
        <Route path="/clients/:id/edit" element={<ClientForm />} />
        <Route path="/clients/:id" element={<ClientDetails />} />
        <Route path="/reservations" element={<Reservations />} />
        <Route path="/reservations/new" element={<ReservationForm />} />
        <Route path="/reservations/:id/edit" element={<ReservationForm />} />
        <Route path="/contracts" element={<Contracts />} />
        <Route path="/payments" element={<Payments />} />
        <Route path="/returns" element={<Returns />} />
        <Route path="/maintenance" element={<Maintenance />} />
        <Route path="/statistics" element={<Statistics />} />
        <Route path="/settings" element={<Settings />} />
        <Route path="/users" element={<Users />} />
      </Route>
      <Route path="*" element={<Navigate to="/" replace />} />
    </Routes>
  );
}
