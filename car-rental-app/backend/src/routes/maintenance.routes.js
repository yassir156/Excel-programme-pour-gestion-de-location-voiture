const express = require('express');
const { Maintenance, Vehicle } = require('../models');
const { authenticate, authorize } = require('../middleware/auth');

const router = express.Router();
router.use(authenticate);

router.get('/', async (req, res, next) => {
  try {
    const { vehicleId } = req.query;
    const where = {};
    if (vehicleId) where.vehicleId = vehicleId;
    const items = await Maintenance.findAll({ where, include: [{ model: Vehicle, as: 'vehicle' }], order: [['date', 'DESC']] });
    res.json(items);
  } catch (err) {
    next(err);
  }
});

router.post('/', authorize('administrateur', 'manager', 'agent'), async (req, res, next) => {
  try {
    const { vehicleId, type, date, cost, description, nextDueDate } = req.body;
    if (!vehicleId || !type || !date) {
      return res.status(400).json({ message: 'Véhicule, type et date sont requis.' });
    }
    const vehicle = await Vehicle.findByPk(vehicleId);
    if (!vehicle) return res.status(404).json({ message: 'Véhicule introuvable.' });

    const item = await Maintenance.create({ vehicleId, type, date, cost, description, nextDueDate });

    if (type === 'assurance' && nextDueDate) vehicle.insuranceExpiry = nextDueDate;
    if (type === 'controle_technique' && nextDueDate) vehicle.technicalControlExpiry = nextDueDate;
    await vehicle.save();

    const full = await Maintenance.findByPk(item.id, { include: [{ model: Vehicle, as: 'vehicle' }] });
    res.status(201).json(full);
  } catch (err) {
    next(err);
  }
});

router.delete('/:id', authorize('administrateur', 'manager'), async (req, res, next) => {
  try {
    const item = await Maintenance.findByPk(req.params.id);
    if (!item) return res.status(404).json({ message: 'Opération introuvable.' });
    await item.destroy();
    res.json({ message: 'Opération de maintenance supprimée.' });
  } catch (err) {
    next(err);
  }
});

module.exports = router;
