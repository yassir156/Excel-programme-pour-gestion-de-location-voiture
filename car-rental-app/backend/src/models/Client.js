const { DataTypes, Model } = require('sequelize');
const { sequelize } = require('../db');

class Client extends Model {}

Client.init(
  {
    id: { type: DataTypes.INTEGER, autoIncrement: true, primaryKey: true },
    firstName: { type: DataTypes.STRING, allowNull: false },
    lastName: { type: DataTypes.STRING, allowNull: false },
    phone: { type: DataTypes.STRING, allowNull: false },
    email: { type: DataTypes.STRING, allowNull: true },
    address: { type: DataTypes.STRING, allowNull: true },
    cin: { type: DataTypes.STRING, allowNull: false, unique: true },
    licenseNumber: { type: DataTypes.STRING, allowNull: true },
    licenseExpiry: { type: DataTypes.DATEONLY, allowNull: true },
    documents: {
      type: DataTypes.TEXT,
      defaultValue: '[]',
      get() {
        const raw = this.getDataValue('documents');
        try { return JSON.parse(raw || '[]'); } catch { return []; }
      },
      set(value) {
        this.setDataValue('documents', JSON.stringify(value || []));
      },
    },
  },
  { sequelize, modelName: 'Client', tableName: 'clients' }
);

module.exports = Client;
