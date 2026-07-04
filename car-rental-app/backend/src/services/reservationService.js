const { Op } = require('sequelize');
const { Reservation, Vehicle } = require('../models');

const ACTIVE_STATUSES = ['en_attente', 'confirmee'];

function diffInDays(start, end) {
  const ms = new Date(end).getTime() - new Date(start).getTime();
  return Math.max(1, Math.round(ms / (1000 * 60 * 60 * 24)));
}

async function isVehicleAvailable(vehicleId, startDate, endDate, excludeReservationId = null) {
  const where = {
    vehicleId,
    status: { [Op.in]: ACTIVE_STATUSES },
    startDate: { [Op.lt]: endDate },
    endDate: { [Op.gt]: startDate },
  };
  if (excludeReservationId) {
    where.id = { [Op.ne]: excludeReservationId };
  }
  const overlapping = await Reservation.findOne({ where });
  return !overlapping;
}

async function computePrice(vehicleId, startDate, endDate) {
  const vehicle = await Vehicle.findByPk(vehicleId);
  if (!vehicle) {
    const err = new Error('Véhicule introuvable.');
    err.status = 404;
    throw err;
  }
  const days = diffInDays(startDate, endDate);
  const totalPrice = Number((days * vehicle.dailyPrice).toFixed(2));
  return { days, totalPrice, vehicle };
}

/**
 * Recomputes a vehicle's status from its reservations so it never gets stuck
 * on "louee"/"reservee" after a cancellation, deletion or return. A manual
 * "maintenance" status always takes priority and is left untouched.
 */
async function syncVehicleStatus(vehicleId) {
  const vehicle = await Vehicle.findByPk(vehicleId);
  if (!vehicle || vehicle.status === 'maintenance') return vehicle;

  const todayStr = new Date().toISOString().slice(0, 10);

  const activeNow = await Reservation.findOne({
    where: {
      vehicleId,
      status: { [Op.in]: ACTIVE_STATUSES },
      startDate: { [Op.lte]: todayStr },
      endDate: { [Op.gte]: todayStr },
    },
  });

  let nextStatus;
  if (activeNow) {
    nextStatus = 'louee';
  } else {
    const upcoming = await Reservation.findOne({
      where: {
        vehicleId,
        status: { [Op.in]: ACTIVE_STATUSES },
        startDate: { [Op.gt]: todayStr },
      },
    });
    nextStatus = upcoming ? 'reservee' : 'disponible';
  }

  if (vehicle.status !== nextStatus) {
    vehicle.status = nextStatus;
    await vehicle.save();
  }
  return vehicle;
}

module.exports = { isVehicleAvailable, computePrice, diffInDays, syncVehicleStatus, ACTIVE_STATUSES };
