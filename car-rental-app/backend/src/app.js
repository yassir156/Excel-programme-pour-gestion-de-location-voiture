const express = require('express');
const cors = require('cors');
const { uploadsDir } = require('./uploads');
const { notFound, errorHandler } = require('./middleware/errorHandler');

const authRoutes = require('./routes/auth.routes');
const usersRoutes = require('./routes/users.routes');
const vehiclesRoutes = require('./routes/vehicles.routes');
const clientsRoutes = require('./routes/clients.routes');
const reservationsRoutes = require('./routes/reservations.routes');
const contractsRoutes = require('./routes/contracts.routes');
const paymentsRoutes = require('./routes/payments.routes');
const returnsRoutes = require('./routes/returns.routes');
const maintenanceRoutes = require('./routes/maintenance.routes');
const statsRoutes = require('./routes/stats.routes');
const settingsRoutes = require('./routes/settings.routes');

function createApp() {
  const app = express();
  app.use(cors());
  app.use(express.json({ limit: '10mb' }));
  app.use('/uploads', express.static(uploadsDir));

  app.get('/api/health', (req, res) => res.json({ status: 'ok' }));

  app.use('/api/auth', authRoutes);
  app.use('/api/users', usersRoutes);
  app.use('/api/vehicles', vehiclesRoutes);
  app.use('/api/clients', clientsRoutes);
  app.use('/api/reservations', reservationsRoutes);
  app.use('/api/contracts', contractsRoutes);
  app.use('/api/payments', paymentsRoutes);
  app.use('/api/returns', returnsRoutes);
  app.use('/api/maintenance', maintenanceRoutes);
  app.use('/api/stats', statsRoutes);
  app.use('/api/settings', settingsRoutes);

  app.use('/api', notFound);
  app.use(errorHandler);

  return app;
}

module.exports = { createApp };
