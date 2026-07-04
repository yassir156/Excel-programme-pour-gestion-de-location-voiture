const { sequelize } = require('../src/models');
const { seedDemoData } = require('./seedLogic');

seedDemoData()
  .then(() => sequelize.close())
  .catch(async (err) => {
    console.error('Erreur lors du seed:', err);
    await sequelize.close();
    process.exit(1);
  });
