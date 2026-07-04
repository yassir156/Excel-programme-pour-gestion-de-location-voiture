const express = require('express');
const { Return, Reservation, Vehicle, Client } = require('../models');
const { authenticate, authorize } = require('../middleware/auth');
const { syncVehicleStatus } = require('../services/reservationService');

const router = express.Router();
router.use(authenticate);

router.get('/', async (req, res, next) => {
  try {
    const returns = await Return.findAll({
      include: [
        { model: Vehicle, as: 'vehicle' },
        { model: Reservation, as: 'reservation', include: [{ model: Client, as: 'client' }] },
      ],
      order: [['returnDate', 'DESC']],
    });
    res.json(returns);
  } catch (err) {
    next(err);
  }
});

router.post('/', authorize('administrateur', 'manager', 'agent'), async (req, res, next) => {
  try {
    const {
      reservationId,
      mileageReturn,
      fuelReturn,
      condition,
      lateFee = 0,
      fuelFee = 0,
      damageFee = 0,
      extraMileageFee = 0,
      notes,
    } = req.body;

    if (!reservationId || mileageReturn === undefined) {
      return res.status(400).json({ message: 'Réservation et kilométrage de retour sont requis.' });
    }

    const reservation = await Reservation.findByPk(reservationId, { include: [{ model: Vehicle, as: 'vehicle' }] });
    if (!reservation) return res.status(404).json({ message: 'Réservation introuvable.' });

    const existing = await Return.findOne({ where: { reservationId } });
    if (existing) {
      return res.status(409).json({ message: 'Un retour a déjà été enregistré pour cette réservation.' });
    }

    const totalExtra = Number(lateFee) + Number(fuelFee) + Number(damageFee) + Number(extraMileageFee);

    const ret = await Return.create({
      reservationId,
      vehicleId: reservation.vehicleId,
      mileageReturn,
      fuelReturn,
      condition,
      lateFee,
      fuelFee,
      damageFee,
      extraMileageFee,
      totalExtra,
      notes,
    });

    reservation.status = 'terminee';
    await reservation.save();

    if (reservation.vehicle) {
      reservation.vehicle.mileage = Math.max(reservation.vehicle.mileage, Number(mileageReturn));
      await reservation.vehicle.save();
    }
    await syncVehicleStatus(reservation.vehicleId);

    const full = await Return.findByPk(ret.id, {
      include: [{ model: Vehicle, as: 'vehicle' }, { model: Reservation, as: 'reservation' }],
    });
    res.status(201).json(full);
  } catch (err) {
    next(err);
  }
});

module.exports = router;
