const express = require('express');
const { Contract, Reservation, Client, Vehicle, AgencySettings } = require('../models');
const { authenticate, authorize } = require('../middleware/auth');
const { generateNumber } = require('../utils/numbering');
const { syncVehicleStatus } = require('../services/reservationService');

const router = express.Router();
router.use(authenticate);

const includeAll = [
  { model: Client, as: 'client' },
  { model: Vehicle, as: 'vehicle' },
  { model: Reservation, as: 'reservation' },
];

router.get('/', async (req, res, next) => {
  try {
    const contracts = await Contract.findAll({ include: includeAll, order: [['createdAt', 'DESC']] });
    res.json(contracts);
  } catch (err) {
    next(err);
  }
});

router.get('/:id', async (req, res, next) => {
  try {
    const contract = await Contract.findByPk(req.params.id, { include: includeAll });
    if (!contract) return res.status(404).json({ message: 'Contrat introuvable.' });
    const settings = await AgencySettings.findOne();
    res.json({ contract, settings });
  } catch (err) {
    next(err);
  }
});

router.post('/from-reservation/:reservationId', authorize('administrateur', 'manager', 'agent'), async (req, res, next) => {
  try {
    const reservation = await Reservation.findByPk(req.params.reservationId, { include: includeAll.filter(i => i.as !== 'reservation') });
    if (!reservation) return res.status(404).json({ message: 'Réservation introuvable.' });

    const existing = await Contract.findOne({ where: { reservationId: reservation.id } });
    if (existing) {
      const full = await Contract.findByPk(existing.id, { include: includeAll });
      return res.json(full);
    }

    const settings = await AgencySettings.findOne();
    const contract = await Contract.create({
      reservationId: reservation.id,
      clientId: reservation.clientId,
      vehicleId: reservation.vehicleId,
      contractNumber: generateNumber('CTR'),
      startDate: reservation.startDate,
      endDate: reservation.endDate,
      totalPrice: reservation.totalPrice,
      deposit: reservation.vehicle?.deposit || 0,
      terms: settings?.contractTerms || '',
    });

    if (reservation.status === 'en_attente') {
      reservation.status = 'confirmee';
      await reservation.save();
    }
    await syncVehicleStatus(reservation.vehicleId);

    const full = await Contract.findByPk(contract.id, { include: includeAll });
    res.status(201).json(full);
  } catch (err) {
    next(err);
  }
});

router.put('/:id/sign', authorize('administrateur', 'manager', 'agent'), async (req, res, next) => {
  try {
    const contract = await Contract.findByPk(req.params.id);
    if (!contract) return res.status(404).json({ message: 'Contrat introuvable.' });
    const { signedClient, signedAgency } = req.body;
    if (signedClient !== undefined) contract.signedClient = signedClient;
    if (signedAgency !== undefined) contract.signedAgency = signedAgency;
    await contract.save();
    res.json(contract);
  } catch (err) {
    next(err);
  }
});

module.exports = router;
