const jwt = require('jsonwebtoken');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

function resolveSecretPath() {
  const dir = process.env.CAR_RENTAL_DB_DIR || path.join(__dirname, '..', '..', 'database');
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  return path.join(dir, '.jwt-secret');
}

function getSecret() {
  const secretPath = resolveSecretPath();
  if (fs.existsSync(secretPath)) {
    return fs.readFileSync(secretPath, 'utf8').trim();
  }
  const secret = crypto.randomBytes(48).toString('hex');
  fs.writeFileSync(secretPath, secret, 'utf8');
  return secret;
}

const SECRET = getSecret();

function signToken(payload) {
  return jwt.sign(payload, SECRET, { expiresIn: '12h' });
}

function verifyToken(token) {
  return jwt.verify(token, SECRET);
}

module.exports = { signToken, verifyToken };
