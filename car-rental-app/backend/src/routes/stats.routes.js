const express = require('express');
const { Op, fn, col, literal } = require('sequelize');
const { Vehicle, Reservation, Payment, Client } = require('../models');
const { authenticate } = require('../middleware/auth');
const { getAlerts } = require('../services/alertService');

const router = express.Router();
router.use(authenticate);

function monthRange(monthsBack) {
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth() - monthsBack + 1, 1);
  return start;
}

router.get('/dashboard', async (req, res, next) => {
  try {
    const now = new Date();
    const monthStart = new Date(now.getFullYear(), now.getMonth(), 1).toISOString().slice(0, 10);
    const monthEnd = new Date(now.getFullYear(), now.getMonth() + 1, 0).toISOString().slice(0, 10);

    const [totalVehicles, available, rented, maintenance, reserved] = await Promise.all([
      Vehicle.count(),
      Vehicle.count({ where: { status: 'disponible' } }),
      Vehicle.count({ where: { status: 'louee' } }),
      Vehicle.count({ where: { status: 'maintenance' } }),
      Vehicle.count({ where: { status: 'reservee' } }),
    ]);

    const ongoingReservations = await Reservation.count({ where: { status: { [Op.in]: ['en_attente', 'confirmee'] } } });

    const monthlyRevenue = await Payment.sum('amount', {
      where: { paidAt: { [Op.gte]: `${monthStart} 00:00:00`, [Op.lte]: `${monthEnd} 23:59:59` } },
    });

    const pendingPayments = await Payment.sum('remaining', { where: { status: { [Op.in]: ['partiel', 'impaye'] } } });

    const upcomingReservations = await Reservation.findAll({
      where: { status: { [Op.in]: ['en_attente', 'confirmee'] } },
      include: [{ association: 'client' }, { association: 'vehicle' }],
      order: [['startDate', 'ASC']],
      limit: 6,
    });

    const upcomingReturns = await Reservation.findAll({
      where: { status: 'confirmee' },
      include: [{ association: 'client' }, { association: 'vehicle' }],
      order: [['endDate', 'ASC']],
      limit: 6,
    });

    const start = monthRange(6);
    const revenueRows = await Payment.findAll({
      attributes: [
        [fn('strftime', '%Y-%m', col('paidAt')), 'month'],
        [fn('SUM', col('amount')), 'total'],
      ],
      where: { paidAt: { [Op.gte]: start } },
      group: [literal("strftime('%Y-%m', paidAt)")],
      order: [[literal("strftime('%Y-%m', paidAt)"), 'ASC']],
      raw: true,
    });

    const alerts = await getAlerts();

    res.json({
      totalVehicles,
      available,
      rented,
      maintenance,
      reserved,
      ongoingReservations,
      monthlyRevenue: monthlyRevenue || 0,
      pendingPayments: pendingPayments || 0,
      upcomingReservations,
      upcomingReturns,
      revenueByMonth: revenueRows,
      alerts,
    });
  } catch (err) {
    next(err);
  }
});

router.get('/reports', async (req, res, next) => {
  try {
    const start = monthRange(11);

    const revenueByMonth = await Payment.findAll({
      attributes: [
        [fn('strftime', '%Y-%m', col('paidAt')), 'month'],
        [fn('SUM', col('amount')), 'total'],
      ],
      where: { paidAt: { [Op.gte]: start } },
      group: [literal("strftime('%Y-%m', paidAt)")],
      order: [[literal("strftime('%Y-%m', paidAt)"), 'ASC']],
      raw: true,
    });

    const mostRentedVehicles = await Reservation.findAll({
      attributes: ['vehicleId', [fn('COUNT', col('Reservation.id')), 'count']],
      include: [{ association: 'vehicle', attributes: ['brand', 'model', 'plate'] }],
      group: ['vehicleId', 'vehicle.id'],
      order: [[literal('count'), 'DESC']],
      limit: 10,
    });

    const mostActiveClients = await Reservation.findAll({
      attributes: ['clientId', [fn('COUNT', col('Reservation.id')), 'count']],
      include: [{ association: 'client', attributes: ['firstName', 'lastName', 'phone'] }],
      group: ['clientId', 'client.id'],
      order: [[literal('count'), 'DESC']],
      limit: 10,
    });

    const totalVehicles = await Vehicle.count();
    const activeReservations = await Reservation.count({ where: { status: { [Op.in]: ['confirmee', 'en_attente'] } } });
    const occupancyRate = totalVehicles > 0 ? Number(((activeReservations / totalVehicles) * 100).toFixed(1)) : 0;

    const cancelledReservations = await Reservation.count({ where: { status: 'annulee' } });
    const pendingPayments = await Payment.sum('remaining', { where: { status: { [Op.in]: ['partiel', 'impaye'] } } });

    res.json({
      revenueByMonth,
      mostRentedVehicles,
      mostActiveClients,
      occupancyRate,
      cancelledReservations,
      pendingPayments: pendingPayments || 0,
    });
  } catch (err) {
    next(err);
  }
});

module.exports = router;
