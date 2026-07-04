const { DataTypes, Model } = require('sequelize');
const { sequelize } = require('../db');

class Return extends Model {}

Return.init(
  {
    id: { type: DataTypes.INTEGER, autoIncrement: true, primaryKey: true },
    reservationId: { type: DataTypes.INTEGER, allowNull: false, unique: true },
    vehicleId: { type: DataTypes.INTEGER, allowNull: false },
    returnDate: { type: DataTypes.DATE, allowNull: false, defaultValue: DataTypes.NOW },
    mileageReturn: { type: DataTypes.INTEGER, allowNull: false },
    fuelReturn: { type: DataTypes.STRING, allowNull: true },
    condition: { type: DataTypes.STRING, allowNull: true },
    lateFee: { type: DataTypes.FLOAT, defaultValue: 0 },
    fuelFee: { type: DataTypes.FLOAT, defaultValue: 0 },
    damageFee: { type: DataTypes.FLOAT, defaultValue: 0 },
    extraMileageFee: { type: DataTypes.FLOAT, defaultValue: 0 },
    totalExtra: { type: DataTypes.FLOAT, defaultValue: 0 },
    notes: { type: DataTypes.STRING, allowNull: true },
  },
  { sequelize, modelName: 'Return', tableName: 'returns' }
);

module.exports = Return;
