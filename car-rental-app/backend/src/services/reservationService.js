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

module.exports = { isVehicleAvailable, computePrice, diffInDays, ACTIVE_STATUSES };
