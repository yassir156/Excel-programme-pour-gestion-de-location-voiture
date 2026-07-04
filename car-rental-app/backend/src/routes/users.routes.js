const express = require('express');
const bcrypt = require('bcryptjs');
const { User } = require('../models');
const { authenticate, authorize } = require('../middleware/auth');

const router = express.Router();
router.use(authenticate, authorize('administrateur'));

router.get('/', async (req, res, next) => {
  try {
    const users = await User.findAll({ attributes: { exclude: ['passwordHash'] }, order: [['id', 'ASC']] });
    res.json(users);
  } catch (err) {
    next(err);
  }
});

router.post('/', async (req, res, next) => {
  try {
    const { username, password, fullName, role } = req.body;
    if (!username || !password || !fullName) {
      return res.status(400).json({ message: 'Nom, identifiant et mot de passe sont requis.' });
    }
    const passwordHash = await bcrypt.hash(password, 10);
    const user = await User.create({ username, passwordHash, fullName, role: role || 'agent' });
    const { passwordHash: _omit, ...safe } = user.toJSON();
    res.status(201).json(safe);
  } catch (err) {
    next(err);
  }
});

router.put('/:id', async (req, res, next) => {
  try {
    const user = await User.findByPk(req.params.id);
    if (!user) return res.status(404).json({ message: 'Utilisateur introuvable.' });
    const { fullName, role, active, password } = req.body;
    if (fullName !== undefined) user.fullName = fullName;
    if (role !== undefined) user.role = role;
    if (active !== undefined) user.active = active;
    if (password) user.passwordHash = await bcrypt.hash(password, 10);
    await user.save();
    const { passwordHash: _omit, ...safe } = user.toJSON();
    res.json(safe);
  } catch (err) {
    next(err);
  }
});

router.delete('/:id', async (req, res, next) => {
  try {
    if (Number(req.params.id) === req.user.id) {
      return res.status(400).json({ message: 'Vous ne pouvez pas supprimer votre propre compte.' });
    }
    const user = await User.findByPk(req.params.id);
    if (!user) return res.status(404).json({ message: 'Utilisateur introuvable.' });
    await user.destroy();
    res.json({ message: 'Utilisateur supprimé.' });
  } catch (err) {
    next(err);
  }
});

module.exports = router;
