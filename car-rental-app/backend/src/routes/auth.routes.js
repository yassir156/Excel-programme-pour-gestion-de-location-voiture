const express = require('express');
const bcrypt = require('bcryptjs');
const { User } = require('../models');
const { signToken } = require('../utils/jwt');
const { authenticate } = require('../middleware/auth');

const router = express.Router();

router.post('/login', async (req, res, next) => {
  try {
    const { username, password } = req.body;
    if (!username || !password) {
      return res.status(400).json({ message: "Nom d'utilisateur et mot de passe requis." });
    }
    const user = await User.findOne({ where: { username } });
    if (!user || !user.active) {
      return res.status(401).json({ message: 'Identifiants incorrects.' });
    }
    const valid = await bcrypt.compare(password, user.passwordHash);
    if (!valid) {
      return res.status(401).json({ message: 'Identifiants incorrects.' });
    }
    const token = signToken({ id: user.id, username: user.username, role: user.role, fullName: user.fullName });
    res.json({
      token,
      user: { id: user.id, username: user.username, role: user.role, fullName: user.fullName },
    });
  } catch (err) {
    next(err);
  }
});

router.get('/me', authenticate, async (req, res, next) => {
  try {
    const user = await User.findByPk(req.user.id, { attributes: { exclude: ['passwordHash'] } });
    if (!user) return res.status(404).json({ message: 'Utilisateur introuvable.' });
    res.json(user);
  } catch (err) {
    next(err);
  }
});

router.post('/change-password', authenticate, async (req, res, next) => {
  try {
    const { currentPassword, newPassword } = req.body;
    if (!newPassword || newPassword.length < 4) {
      return res.status(400).json({ message: 'Le nouveau mot de passe doit contenir au moins 4 caractères.' });
    }
    const user = await User.findByPk(req.user.id);
    const valid = await bcrypt.compare(currentPassword || '', user.passwordHash);
    if (!valid) {
      return res.status(401).json({ message: 'Mot de passe actuel incorrect.' });
    }
    user.passwordHash = await bcrypt.hash(newPassword, 10);
    await user.save();
    res.json({ message: 'Mot de passe modifié avec succès.' });
  } catch (err) {
    next(err);
  }
});

module.exports = router;
