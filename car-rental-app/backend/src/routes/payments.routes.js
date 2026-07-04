const express = require('express');
const { Payment, Reservation, Client, Vehicle } = require('../models');
const { authenticate, authorize } = require('../middleware/auth');
const { generateNumber } = require('../utils/numbering');

const router = express.Router();
router.use(authenticate);

const includeAll = [
  { model: Client, as: 'client' },
  {
    model: Reservation,
    as: 'reservation',
    include: [{ model: Vehicle, as: 'vehicle' }],
  },
];

router.get('/', async (req, res, next) => {
  try {
    const { reservationId, status } = req.query;
    const where = {};
    if (reservationId) where.reservationId = reservationId;
    if (status) where.status = status;
    const payments = await Payment.findAll({ where, include: includeAll, order: [['paidAt', 'DESC']] });
    res.json(payments);
  } catch (err) {
    next(err);
  }
});

router.get('/:id', async (req, res, next) => {
  try {
    const payment = await Payment.findByPk(req.params.id, { include: includeAll });
    if (!payment) return res.status(404).json({ message: 'Paiement introuvable.' });
    res.json(payment);
  } catch (err) {
    next(err);
  }
});

router.post('/', authorize('administrateur', 'manager', 'agent'), async (req, res, next) => {
  try {
    const { reservationId, amount, method, depositPaid } = req.body;
    if (!reservationId || amount === undefined) {
      return res.status(400).json({ message: 'Réservation et montant sont requis.' });
    }
    const reservation = await Reservation.findByPk(reservationId);
    if (!reservation) return res.status(404).json({ message: 'Réservation introuvable.' });

    const previousPayments = await Payment.sum('amount', { where: { reservationId } });
    const totalPaid = (previousPayments || 0) + Number(amount);
    const remaining = Math.max(0, Number((reservation.totalPrice - totalPaid).toFixed(2)));
    let status = 'partiel';
    if (remaining <= 0) status = 'paye';
    if (totalPaid <= 0) status = 'impaye';

    const payment = await Payment.create({
      reservationId,
      clientId: reservation.clientId,
      receiptNumber: generateNumber('REC'),
      amount,
      method: method || 'especes',
      depositPaid: !!depositPaid,
      status,
      remaining,
    });

    const full = await Payment.findByPk(payment.id, { include: includeAll });
    res.status(201).json(full);
  } catch (err) {
    next(err);
  }
});

router.delete('/:id', authorize('administrateur', 'manager'), async (req, res, next) => {
  try {
    const payment = await Payment.findByPk(req.params.id);
    if (!payment) return res.status(404).json({ message: 'Paiement introuvable.' });
    await payment.destroy();
    res.json({ message: 'Paiement supprimé.' });
  } catch (err) {
    next(err);
  }
});

module.exports = router;
