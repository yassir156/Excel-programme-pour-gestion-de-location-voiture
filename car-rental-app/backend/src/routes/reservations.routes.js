const express = require('express');
const { Op } = require('sequelize');
const { Reservation, Client, Vehicle } = require('../models');
const { authenticate, authorize } = require('../middleware/auth');
const { isVehicleAvailable, computePrice, syncVehicleStatus } = require('../services/reservationService');

const router = express.Router();
router.use(authenticate);

const includeAll = [
  { model: Client, as: 'client' },
  { model: Vehicle, as: 'vehicle' },
];

router.get('/', async (req, res, next) => {
  try {
    const { status, clientId, vehicleId, from, to } = req.query;
    const where = {};
    if (status) where.status = status;
    if (clientId) where.clientId = clientId;
    if (vehicleId) where.vehicleId = vehicleId;
    if (from) where.endDate = { [Op.gte]: from };
    if (to) where.startDate = { ...(where.startDate || {}), [Op.lte]: to };
    const reservations = await Reservation.findAll({ where, include: includeAll, order: [['startDate', 'DESC']] });
    res.json(reservations);
  } catch (err) {
    next(err);
  }
});

router.get('/check-availability', async (req, res, next) => {
  try {
    const { vehicleId, startDate, endDate, excludeReservationId } = req.query;
    if (!vehicleId || !startDate || !endDate) {
      return res.status(400).json({ message: 'vehicleId, startDate et endDate sont requis.' });
    }
    const available = await isVehicleAvailable(vehicleId, startDate, endDate, excludeReservationId || null);
    const { days, totalPrice } = await computePrice(vehicleId, startDate, endDate);
    res.json({ available, days, totalPrice });
  } catch (err) {
    next(err);
  }
});

router.get('/:id', async (req, res, next) => {
  try {
    const reservation = await Reservation.findByPk(req.params.id, { include: includeAll });
    if (!reservation) return res.status(404).json({ message: 'Réservation introuvable.' });
    res.json(reservation);
  } catch (err) {
    next(err);
  }
});

router.post('/', authorize('administrateur', 'manager', 'agent'), async (req, res, next) => {
  try {
    const { clientId, vehicleId, startDate, endDate, notes } = req.body;
    if (!clientId || !vehicleId || !startDate || !endDate) {
      return res.status(400).json({ message: 'Client, véhicule, date de début et date de fin sont requis.' });
    }
    if (new Date(startDate) >= new Date(endDate)) {
      return res.status(400).json({ message: 'La date de fin doit être après la date de début.' });
    }
    const available = await isVehicleAvailable(vehicleId, startDate, endDate);
    if (!available) {
      return res.status(409).json({ message: 'Ce véhicule est déjà réservé sur cette période.' });
    }
    const { days, totalPrice } = await computePrice(vehicleId, startDate, endDate);
    const reservation = await Reservation.create({
      clientId,
      vehicleId,
      startDate,
      endDate,
      days,
      totalPrice,
      notes,
      status: 'en_attente',
    });
    await syncVehicleStatus(vehicleId);
    const full = await Reservation.findByPk(reservation.id, { include: includeAll });
    res.status(201).json(full);
  } catch (err) {
    next(err);
  }
});

router.put('/:id', authorize('administrateur', 'manager', 'agent'), async (req, res, next) => {
  try {
    const reservation = await Reservation.findByPk(req.params.id);
    if (!reservation) return res.status(404).json({ message: 'Réservation introuvable.' });
    const previousVehicleId = reservation.vehicleId;
    const { vehicleId, startDate, endDate, notes, status } = req.body;
    const nextVehicleId = vehicleId || reservation.vehicleId;
    const nextStart = startDate || reservation.startDate;
    const nextEnd = endDate || reservation.endDate;

    if (startDate || endDate || vehicleId) {
      if (new Date(nextStart) >= new Date(nextEnd)) {
        return res.status(400).json({ message: 'La date de fin doit être après la date de début.' });
      }
      const available = await isVehicleAvailable(nextVehicleId, nextStart, nextEnd, reservation.id);
      if (!available) {
        return res.status(409).json({ message: 'Ce véhicule est déjà réservé sur cette période.' });
      }
      const { days, totalPrice } = await computePrice(nextVehicleId, nextStart, nextEnd);
      reservation.vehicleId = nextVehicleId;
      reservation.startDate = nextStart;
      reservation.endDate = nextEnd;
      reservation.days = days;
      reservation.totalPrice = totalPrice;
    }
    if (notes !== undefined) reservation.notes = notes;
    if (status !== undefined) reservation.status = status;
    await reservation.save();
    await syncVehicleStatus(reservation.vehicleId);
    if (previousVehicleId !== reservation.vehicleId) {
      await syncVehicleStatus(previousVehicleId);
    }
    const full = await Reservation.findByPk(reservation.id, { include: includeAll });
    res.json(full);
  } catch (err) {
    next(err);
  }
});

router.post('/:id/cancel', authorize('administrateur', 'manager', 'agent'), async (req, res, next) => {
  try {
    const reservation = await Reservation.findByPk(req.params.id);
    if (!reservation) return res.status(404).json({ message: 'Réservation introuvable.' });
    reservation.status = 'annulee';
    await reservation.save();
    await syncVehicleStatus(reservation.vehicleId);
    res.json(reservation);
  } catch (err) {
    next(err);
  }
});

router.delete('/:id', authorize('administrateur', 'manager'), async (req, res, next) => {
  try {
    const reservation = await Reservation.findByPk(req.params.id);
    if (!reservation) return res.status(404).json({ message: 'Réservation introuvable.' });
    const { vehicleId } = reservation;
    await reservation.destroy();
    await syncVehicleStatus(vehicleId);
    res.json({ message: 'Réservation supprimée.' });
  } catch (err) {
    next(err);
  }
});

module.exports = router;
