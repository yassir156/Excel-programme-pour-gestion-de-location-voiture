const { DataTypes, Model } = require('sequelize');
const { sequelize } = require('../db');

class Maintenance extends Model {}

Maintenance.init(
  {
    id: { type: DataTypes.INTEGER, autoIncrement: true, primaryKey: true },
    vehicleId: { type: DataTypes.INTEGER, allowNull: false },
    type: {
      type: DataTypes.ENUM('vidange', 'reparation', 'pneus', 'assurance', 'controle_technique', 'autre'),
      allowNull: false,
    },
    date: { type: DataTypes.DATEONLY, allowNull: false },
    cost: { type: DataTypes.FLOAT, defaultValue: 0 },
    description: { type: DataTypes.STRING, allowNull: true },
    nextDueDate: { type: DataTypes.DATEONLY, allowNull: true },
  },
  { sequelize, modelName: 'Maintenance', tableName: 'maintenance' }
);

module.exports = Maintenance;
