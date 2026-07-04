const { sequelize } = require('../db');
const User = require('./User');
const Vehicle = require('./Vehicle');
const Client = require('./Client');
const Reservation = require('./Reservation');
const Contract = require('./Contract');
const Payment = require('./Payment');
const Return = require('./Return');
const Maintenance = require('./Maintenance');
const AgencySettings = require('./AgencySettings');

// Client <-> Reservation
Client.hasMany(Reservation, { foreignKey: 'clientId', as: 'reservations' });
Reservation.belongsTo(Client, { foreignKey: 'clientId', as: 'client' });

// Vehicle <-> Reservation
Vehicle.hasMany(Reservation, { foreignKey: 'vehicleId', as: 'reservations' });
Reservation.belongsTo(Vehicle, { foreignKey: 'vehicleId', as: 'vehicle' });

// Reservation <-> Contract
Reservation.hasOne(Contract, { foreignKey: 'reservationId', as: 'contract' });
Contract.belongsTo(Reservation, { foreignKey: 'reservationId', as: 'reservation' });
Contract.belongsTo(Client, { foreignKey: 'clientId', as: 'client' });
Contract.belongsTo(Vehicle, { foreignKey: 'vehicleId', as: 'vehicle' });

// Reservation <-> Payment
Reservation.hasMany(Payment, { foreignKey: 'reservationId', as: 'payments' });
Payment.belongsTo(Reservation, { foreignKey: 'reservationId', as: 'reservation' });
Payment.belongsTo(Client, { foreignKey: 'clientId', as: 'client' });

// Reservation <-> Return
Reservation.hasOne(Return, { foreignKey: 'reservationId', as: 'return' });
Return.belongsTo(Reservation, { foreignKey: 'reservationId', as: 'reservation' });
Return.belongsTo(Vehicle, { foreignKey: 'vehicleId', as: 'vehicle' });

// Vehicle <-> Maintenance
Vehicle.hasMany(Maintenance, { foreignKey: 'vehicleId', as: 'maintenances' });
Maintenance.belongsTo(Vehicle, { foreignKey: 'vehicleId', as: 'vehicle' });

async function initDatabase() {
  await sequelize.sync();
  const settingsCount = await AgencySettings.count();
  if (settingsCount === 0) {
    await AgencySettings.create({});
  }
}

module.exports = {
  sequelize,
  initDatabase,
  User,
  Vehicle,
  Client,
  Reservation,
  Contract,
  Payment,
  Return,
  Maintenance,
  AgencySettings,
};
