const { DataTypes, Model } = require('sequelize');
const { sequelize } = require('../db');

class Reservation extends Model {}

Reservation.init(
  {
    id: { type: DataTypes.INTEGER, autoIncrement: true, primaryKey: true },
    clientId: { type: DataTypes.INTEGER, allowNull: false },
    vehicleId: { type: DataTypes.INTEGER, allowNull: false },
    startDate: { type: DataTypes.DATEONLY, allowNull: false },
    endDate: { type: DataTypes.DATEONLY, allowNull: false },
    days: { type: DataTypes.INTEGER, allowNull: false },
    totalPrice: { type: DataTypes.FLOAT, allowNull: false, defaultValue: 0 },
    status: {
      type: DataTypes.ENUM('en_attente', 'confirmee', 'annulee', 'terminee'),
      defaultValue: 'en_attente',
    },
    notes: { type: DataTypes.STRING, allowNull: true },
  },
  { sequelize, modelName: 'Reservation', tableName: 'reservations' }
);

module.exports = Reservation;
