const { Op } = require('sequelize');
const { Vehicle, Maintenance, Reservation } = require('../models');

const SOON_DAYS = 30;

function addDays(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

async function getAlerts() {
  const today = new Date();
  const soon = addDays(today, SOON_DAYS);
  const todayStr = today.toISOString().slice(0, 10);
  const soonStr = soon.toISOString().slice(0, 10);

  const insuranceSoon = await Vehicle.findAll({
    where: { insuranceExpiry: { [Op.ne]: null, [Op.lte]: soonStr } },
    order: [['insuranceExpiry', 'ASC']],
  });

  const technicalSoon = await Vehicle.findAll({
    where: { technicalControlExpiry: { [Op.ne]: null, [Op.lte]: soonStr } },
    order: [['technicalControlExpiry', 'ASC']],
  });

  const maintenanceSoon = await Maintenance.findAll({
    where: { nextDueDate: { [Op.ne]: null, [Op.lte]: soonStr } },
    include: [{ model: Vehicle, as: 'vehicle' }],
    order: [['nextDueDate', 'ASC']],
  });

  const upcomingReturns = await Reservation.findAll({
    where: { status: 'confirmee', endDate: { [Op.lte]: soonStr, [Op.gte]: todayStr } },
    order: [['endDate', 'ASC']],
    limit: 10,
  });

  return {
    insuranceExpiringSoon: insuranceSoon,
    technicalControlExpiringSoon: technicalSoon,
    maintenanceDueSoon: maintenanceSoon,
    upcomingReturns,
  };
}

module.exports = { getAlerts, SOON_DAYS };
