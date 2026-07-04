function notFound(req, res) {
  res.status(404).json({ message: 'Ressource introuvable.' });
}

function errorHandler(err, req, res, next) { // eslint-disable-line no-unused-vars
  console.error(err);
  if (err.name === 'SequelizeUniqueConstraintError') {
    return res.status(409).json({ message: 'Une entrée avec ces informations existe déjà.', details: err.errors?.map((e) => e.message) });
  }
  if (err.name === 'SequelizeValidationError') {
    return res.status(400).json({ message: 'Données invalides.', details: err.errors?.map((e) => e.message) });
  }
  const status = err.status || 500;
  res.status(status).json({ message: err.message || 'Erreur serveur interne.' });
}

module.exports = { notFound, errorHandler };
