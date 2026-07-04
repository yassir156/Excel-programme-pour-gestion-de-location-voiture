const express = require('express');
const fs = require('fs');
const path = require('path');
const { AgencySettings } = require('../models');
const { authenticate, authorize } = require('../middleware/auth');
const { upload } = require('../uploads');
const { sequelize, dbPath } = require('../db');

const router = express.Router();
router.use(authenticate);

router.get('/', async (req, res, next) => {
  try {
    let settings = await AgencySettings.findOne();
    if (!settings) settings = await AgencySettings.create({});
    res.json(settings);
  } catch (err) {
    next(err);
  }
});

router.put('/', authorize('administrateur', 'manager'), async (req, res, next) => {
  try {
    let settings = await AgencySettings.findOne();
    if (!settings) settings = await AgencySettings.create({});
    await settings.update(req.body);
    res.json(settings);
  } catch (err) {
    next(err);
  }
});

router.post('/logo', authorize('administrateur', 'manager'), upload.single('logo'), async (req, res, next) => {
  try {
    let settings = await AgencySettings.findOne();
    if (!settings) settings = await AgencySettings.create({});
    if (!req.file) return res.status(400).json({ message: 'Aucun fichier fourni.' });
    settings.logo = `/uploads/${path.basename(req.file.path)}`;
    await settings.save();
    res.json(settings);
  } catch (err) {
    next(err);
  }
});

router.get('/backup', authorize('administrateur'), async (req, res, next) => {
  try {
    await sequelize.query('PRAGMA wal_checkpoint(FULL);').catch(() => {});
    res.download(dbPath, `sauvegarde-car-rental-${Date.now()}.sqlite`);
  } catch (err) {
    next(err);
  }
});

router.post('/restore', authorize('administrateur'), upload.single('database'), async (req, res, next) => {
  try {
    if (!req.file) return res.status(400).json({ message: 'Aucun fichier de sauvegarde fourni.' });
    await sequelize.close();
    fs.copyFileSync(req.file.path, dbPath);
    res.json({ message: 'Base de données restaurée. Veuillez redémarrer l\'application.' });
  } catch (err) {
    next(err);
  }
});

module.exports = router;
