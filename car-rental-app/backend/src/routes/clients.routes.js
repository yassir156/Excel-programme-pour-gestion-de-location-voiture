const express = require('express');
const { Op } = require('sequelize');
const { Client, Reservation, Vehicle } = require('../models');
const { authenticate, authorize } = require('../middleware/auth');
const { upload } = require('../uploads');
const path = require('path');

const router = express.Router();
router.use(authenticate);

router.get('/', async (req, res, next) => {
  try {
    const { search } = req.query;
    const where = {};
    if (search) {
      where[Op.or] = [
        { firstName: { [Op.like]: `%${search}%` } },
        { lastName: { [Op.like]: `%${search}%` } },
        { phone: { [Op.like]: `%${search}%` } },
        { cin: { [Op.like]: `%${search}%` } },
        { licenseNumber: { [Op.like]: `%${search}%` } },
      ];
    }
    const clients = await Client.findAll({ where, order: [['createdAt', 'DESC']] });
    res.json(clients);
  } catch (err) {
    next(err);
  }
});

router.get('/:id', async (req, res, next) => {
  try {
    const client = await Client.findByPk(req.params.id);
    if (!client) return res.status(404).json({ message: 'Client introuvable.' });
    res.json(client);
  } catch (err) {
    next(err);
  }
});

router.get('/:id/history', async (req, res, next) => {
  try {
    const reservations = await Reservation.findAll({
      where: { clientId: req.params.id },
      include: [{ model: Vehicle, as: 'vehicle' }],
      order: [['startDate', 'DESC']],
    });
    res.json(reservations);
  } catch (err) {
    next(err);
  }
});

router.post('/', authorize('administrateur', 'manager', 'agent'), async (req, res, next) => {
  try {
    const client = await Client.create(req.body);
    res.status(201).json(client);
  } catch (err) {
    next(err);
  }
});

router.put('/:id', authorize('administrateur', 'manager', 'agent'), async (req, res, next) => {
  try {
    const client = await Client.findByPk(req.params.id);
    if (!client) return res.status(404).json({ message: 'Client introuvable.' });
    await client.update(req.body);
    res.json(client);
  } catch (err) {
    next(err);
  }
});

router.delete('/:id', authorize('administrateur', 'manager'), async (req, res, next) => {
  try {
    const client = await Client.findByPk(req.params.id);
    if (!client) return res.status(404).json({ message: 'Client introuvable.' });
    const activeReservations = await Reservation.count({
      where: { clientId: client.id, status: { [Op.in]: ['en_attente', 'confirmee'] } },
    });
    if (activeReservations > 0) {
      return res.status(400).json({ message: 'Impossible de supprimer : réservations actives liées à ce client.' });
    }
    await client.destroy();
    res.json({ message: 'Client supprimé.' });
  } catch (err) {
    next(err);
  }
});

router.post('/:id/documents', authorize('administrateur', 'manager', 'agent'), upload.array('documents', 5), async (req, res, next) => {
  try {
    const client = await Client.findByPk(req.params.id);
    if (!client) return res.status(404).json({ message: 'Client introuvable.' });
    const newDocs = (req.files || []).map((f) => `/uploads/${path.basename(f.path)}`);
    client.documents = [...client.documents, ...newDocs];
    await client.save();
    res.json(client);
  } catch (err) {
    next(err);
  }
});

module.exports = router;
