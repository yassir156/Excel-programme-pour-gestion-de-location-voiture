const { DataTypes, Model } = require('sequelize');
const { sequelize } = require('../db');

class User extends Model {}

User.init(
  {
    id: { type: DataTypes.INTEGER, autoIncrement: true, primaryKey: true },
    username: { type: DataTypes.STRING, allowNull: false, unique: true },
    passwordHash: { type: DataTypes.STRING, allowNull: false },
    fullName: { type: DataTypes.STRING, allowNull: false },
    role: {
      type: DataTypes.ENUM('administrateur', 'manager', 'agent'),
      allowNull: false,
      defaultValue: 'agent',
    },
    active: { type: DataTypes.BOOLEAN, defaultValue: true },
  },
  { sequelize, modelName: 'User', tableName: 'users' }
);

module.exports = User;
