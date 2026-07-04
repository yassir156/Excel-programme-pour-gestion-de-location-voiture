const { DataTypes, Model } = require('sequelize');
const { sequelize } = require('../db');

class AgencySettings extends Model {}

AgencySettings.init(
  {
    id: { type: DataTypes.INTEGER, autoIncrement: true, primaryKey: true },
    name: { type: DataTypes.STRING, defaultValue: 'Mon Agence de Location' },
    address: { type: DataTypes.STRING, defaultValue: '' },
    phone: { type: DataTypes.STRING, defaultValue: '' },
    email: { type: DataTypes.STRING, defaultValue: '' },
    logo: { type: DataTypes.STRING, allowNull: true },
    currency: { type: DataTypes.STRING, defaultValue: 'MAD' },
    taxRate: { type: DataTypes.FLOAT, defaultValue: 0 },
    contractTerms: {
      type: DataTypes.TEXT,
      defaultValue:
        "1. Le locataire s'engage à restituer le véhicule dans l'état où il l'a reçu.\n" +
        "2. Toute infraction au code de la route est à la charge exclusive du locataire.\n" +
        "3. La caution sera restituée après vérification de l'état du véhicule.\n" +
        "4. Tout retard de restitution entraîne des frais supplémentaires.\n" +
        "5. Le carburant manquant sera facturé au tarif en vigueur.",
    },
  },
  { sequelize, modelName: 'AgencySettings', tableName: 'agency_settings' }
);

module.exports = AgencySettings;
