const { DataTypes, Model } = require('sequelize');
const { sequelize } = require('../db');

class Contract extends Model {}

Contract.init(
  {
    id: { type: DataTypes.INTEGER, autoIncrement: true, primaryKey: true },
    reservationId: { type: DataTypes.INTEGER, allowNull: false, unique: true },
    clientId: { type: DataTypes.INTEGER, allowNull: false },
    vehicleId: { type: DataTypes.INTEGER, allowNull: false },
    contractNumber: { type: DataTypes.STRING, allowNull: false, unique: true },
    startDate: { type: DataTypes.DATEONLY, allowNull: false },
    endDate: { type: DataTypes.DATEONLY, allowNull: false },
    totalPrice: { type: DataTypes.FLOAT, allowNull: false },
    deposit: { type: DataTypes.FLOAT, allowNull: false, defaultValue: 0 },
    terms: { type: DataTypes.TEXT, allowNull: true },
    signedClient: { type: DataTypes.BOOLEAN, defaultValue: false },
    signedAgency: { type: DataTypes.BOOLEAN, defaultValue: false },
  },
  { sequelize, modelName: 'Contract', tableName: 'contracts' }
);

module.exports = Contract;
