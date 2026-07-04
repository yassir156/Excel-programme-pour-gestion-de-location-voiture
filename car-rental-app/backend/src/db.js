const path = require('path');
const fs = require('fs');
const { Sequelize } = require('sequelize');

function resolveDbPath() {
  const dbDir = process.env.CAR_RENTAL_DB_DIR || path.join(__dirname, '..', 'database');
  if (!fs.existsSync(dbDir)) {
    fs.mkdirSync(dbDir, { recursive: true });
  }
  return path.join(dbDir, 'car-rental.sqlite');
}

const storage = resolveDbPath();

const sequelize = new Sequelize({
  dialect: 'sqlite',
  storage,
  logging: false,
});

module.exports = { sequelize, dbPath: storage };
