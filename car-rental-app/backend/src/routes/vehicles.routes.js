const express = require('express');
const fs = require('fs');
const { Op } = require('sequelize');
const { Vehicle, Reservation, Client, Maintenance } = require('../models');
const { authenticate, authorize } = require('../middleware/auth');
const { upload, uploadsDir } = require('../uploads');
const path = require('path');

const router = express.Router();
router.use(authenticate);

router.get('/', async (req, res, next) => {
  try {
    const { search, status, category, fuelType, transmission } = req.query;
    const where = {};
    if (status) where.status = status;
    if (category) where.category = category;
    if (fuelType) where.fuelType = fuelType;
    if (transmission) where.transmission = transmission;
    if (search) {
      where[Op.or] = [
        { brand: { [Op.like]: `%${search}%` } },
        { model: { [Op.like]: `%${search}%` } },
        { plate: { [Op.like]: `%${search}%` } },
      ];
    }
    const vehicles = await Vehicle.findAll({ where, order: [['createdAt', 'DESC']] });
    res.json(vehicles);
  } catch (err) {
    next(err);
  }
});

router.get('/:id', async (req, res, next) => {
  try {
    const vehicle = await Vehicle.findByPk(req.params.id);
    if (!vehicle) return res.status(404).json({ message: 'Véhicule introuvable.' });
    res.json(vehicle);
  } catch (err) {
    next(err);
  }
});

router.get('/:id/history', async (req, res, next) => {
  try {
    const reservations = await Reservation.findAll({
      where: { vehicleId: req.params.id },
      include: [{ model: Client, as: 'client' }],
      order: [['startDate', 'DESC']],
    });
    const maintenances = await Maintenance.findAll({
      where: { vehicleId: req.params.id },
      order: [['date', 'DESC']],
    });
    res.json({ reservations, maintenances });
  } catch (err) {
    next(err);
  }
});

router.post('/', authorize('administrateur', 'manager', 'agent'), async (req, res, next) => {
  try {
    const vehicle = await Vehicle.create(req.body);
    res.status(201).json(vehicle);
  } catch (err) {
    next(err);
  }
});

router.put('/:id', authorize('administrateur', 'manager', 'agent'), async (req, res, next) => {
  try {
    const vehicle = await Vehicle.findByPk(req.params.id);
    if (!vehicle) return res.status(404).json({ message: 'Véhicule introuvable.' });
    await vehicle.update(req.body);
    res.json(vehicle);
  } catch (err) {
    next(err);
  }
});

router.delete('/:id', authorize('administrateur', 'manager'), async (req, res, next) => {
  try {
    const vehicle = await Vehicle.findByPk(req.params.id);
    if (!vehicle) return res.status(404).json({ message: 'Véhicule introuvable.' });
    const activeReservations = await Reservation.count({
      where: { vehicleId: vehicle.id, status: { [Op.in]: ['en_attente', 'confirmee'] } },
    });
    if (activeReservations > 0) {
      return res.status(400).json({ message: 'Impossible de supprimer : réservations actives liées à ce véhicule.' });
    }
    await vehicle.destroy();
    res.json({ message: 'Véhicule supprimé.' });
  } catch (err) {
    next(err);
  }
});

router.post('/:id/photos', authorize('administrateur', 'manager', 'agent'), upload.array('photos', 8), async (req, res, next) => {
  try {
    const vehicle = await Vehicle.findByPk(req.params.id);
    if (!vehicle) return res.status(404).json({ message: 'Véhicule introuvable.' });
    const newPhotos = (req.files || []).map((f) => `/uploads/${path.basename(f.path)}`);
    vehicle.photos = [...vehicle.photos, ...newPhotos];
    await vehicle.save();
    res.json(vehicle);
  } catch (err) {
    next(err);
  }
});

router.delete('/:id/photos', authorize('administrateur', 'manager', 'agent'), async (req, res, next) => {
  try {
    const { photo } = req.body;
    if (!photo) return res.status(400).json({ message: 'Photo à supprimer non spécifiée.' });
    const vehicle = await Vehicle.findByPk(req.params.id);
    if (!vehicle) return res.status(404).json({ message: 'Véhicule introuvable.' });
    vehicle.photos = vehicle.photos.filter((p) => p !== photo);
    await vehicle.save();
    const filePath = path.join(uploadsDir, path.basename(photo));
    fs.unlink(filePath, () => {});
    res.json(vehicle);
  } catch (err) {
    next(err);
  }
});

module.exports = router;
