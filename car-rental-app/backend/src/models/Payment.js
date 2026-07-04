const { DataTypes, Model } = require('sequelize');
const { sequelize } = require('../db');

class Payment extends Model {}

Payment.init(
  {
    id: { type: DataTypes.INTEGER, autoIncrement: true, primaryKey: true },
    reservationId: { type: DataTypes.INTEGER, allowNull: false },
    clientId: { type: DataTypes.INTEGER, allowNull: false },
    receiptNumber: { type: DataTypes.STRING, allowNull: false, unique: true },
    amount: { type: DataTypes.FLOAT, allowNull: false },
    method: {
      type: DataTypes.ENUM('especes', 'carte', 'virement', 'cheque'),
      allowNull: false,
      defaultValue: 'especes',
    },
    depositPaid: { type: DataTypes.BOOLEAN, defaultValue: false },
    status: {
      type: DataTypes.ENUM('paye', 'partiel', 'impaye'),
      defaultValue: 'impaye',
    },
    remaining: { type: DataTypes.FLOAT, defaultValue: 0 },
    paidAt: { type: DataTypes.DATE, defaultValue: DataTypes.NOW },
  },
  { sequelize, modelName: 'Payment', tableName: 'payments' }
);

module.exports = Payment;
