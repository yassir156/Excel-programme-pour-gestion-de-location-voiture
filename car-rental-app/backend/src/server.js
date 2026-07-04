const { createApp } = require('./app');
const { initDatabase } = require('./models');

const DEFAULT_PORT = process.env.CAR_RENTAL_PORT || 4310;

async function startServer(port = DEFAULT_PORT) {
  await initDatabase();
  const app = createApp();
  return new Promise((resolve, reject) => {
    const server = app.listen(port, '127.0.0.1', () => {
      resolve(server);
    });
    server.on('error', reject);
  });
}

module.exports = { startServer };

if (require.main === module) {
  startServer().then(() => {
    console.log(`Serveur backend démarré sur http://127.0.0.1:${DEFAULT_PORT}`);
  }).catch((err) => {
    console.error('Échec du démarrage du serveur:', err);
    process.exit(1);
  });
}
