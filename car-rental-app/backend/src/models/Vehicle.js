const { DataTypes, Model } = require('sequelize');
const { sequelize } = require('../db');

class Vehicle extends Model {}

Vehicle.init(
  {
    id: { type: DataTypes.INTEGER, autoIncrement: true, primaryKey: true },
    brand: { type: DataTypes.STRING, allowNull: false },
    model: { type: DataTypes.STRING, allowNull: false },
    year: { type: DataTypes.INTEGER, allowNull: false },
    plate: { type: DataTypes.STRING, allowNull: false, unique: true },
    category: { type: DataTypes.STRING, allowNull: true },
    color: { type: DataTypes.STRING, allowNull: true },
    mileage: { type: DataTypes.INTEGER, defaultValue: 0 },
    fuelType: {
      type: DataTypes.ENUM('essence', 'diesel', 'hybride', 'electrique'),
      defaultValue: 'essence',
    },
    transmission: {
      type: DataTypes.ENUM('manuelle', 'automatique'),
      defaultValue: 'manuelle',
    },
    dailyPrice: { type: DataTypes.FLOAT, allowNull: false, defaultValue: 0 },
    deposit: { type: DataTypes.FLOAT, allowNull: false, defaultValue: 0 },
    status: {
      type: DataTypes.ENUM('disponible', 'louee', 'maintenance', 'reservee'),
      defaultValue: 'disponible',
    },
    insuranceExpiry: { type: DataTypes.DATEONLY, allowNull: true },
    technicalControlExpiry: { type: DataTypes.DATEONLY, allowNull: true },
    photos: {
      type: DataTypes.TEXT,
      defaultValue: '[]',
      get() {
        const raw = this.getDataValue('photos');
        try { return JSON.parse(raw || '[]'); } catch { return []; }
      },
      set(value) {
        this.setDataValue('photos', JSON.stringify(value || []));
      },
    },
  },
  { sequelize, modelName: 'Vehicle', tableName: 'vehicles' }
);

module.exports = Vehicle;
