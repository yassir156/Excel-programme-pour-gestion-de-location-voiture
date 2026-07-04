const bcrypt = require('bcryptjs');
const {
  initDatabase,
  User,
  Vehicle,
  Client,
  Reservation,
  Contract,
  Payment,
  Return,
  Maintenance,
  AgencySettings,
} = require('../src/models');
const { generateNumber } = require('../src/utils/numbering');

function daysFromNow(n) {
  const d = new Date();
  d.setDate(d.getDate() + n);
  return d.toISOString().slice(0, 10);
}

async function seedDemoData() {
  await initDatabase();

  const userCount = await User.count();
  if (userCount === 0) {
    const password = await bcrypt.hash('admin123', 10);
    await User.bulkCreate([
      { username: 'admin', passwordHash: password, fullName: 'Administrateur Principal', role: 'administrateur' },
      { username: 'manager', passwordHash: await bcrypt.hash('manager123', 10), fullName: 'Sara Manager', role: 'manager' },
      { username: 'agent', passwordHash: await bcrypt.hash('agent123', 10), fullName: 'Youssef Agent', role: 'agent' },
    ]);
    console.log('Utilisateurs de démonstration créés (admin/admin123, manager/manager123, agent/agent123).');
  }

  let settings = await AgencySettings.findOne();
  if (!settings) settings = await AgencySettings.create({});
  await settings.update({
    name: 'AutoLoc Premium',
    address: '12 Avenue Mohammed V, Casablanca',
    phone: '+212 522 00 00 00',
    email: 'contact@autoloc-premium.ma',
    currency: 'MAD',
    taxRate: 20,
  });

  const vehicleCount = await Vehicle.count();
  if (vehicleCount === 0) {
    const vehicles = await Vehicle.bulkCreate([
      { brand: 'Dacia', model: 'Logan', year: 2022, plate: '12345-A-1', category: 'Économique', color: 'Blanc', mileage: 32000, fuelType: 'diesel', transmission: 'manuelle', dailyPrice: 250, deposit: 2000, status: 'disponible', insuranceExpiry: daysFromNow(20), technicalControlExpiry: daysFromNow(200) },
      { brand: 'Renault', model: 'Clio', year: 2023, plate: '54321-B-2', category: 'Compacte', color: 'Gris', mileage: 15000, fuelType: 'essence', transmission: 'manuelle', dailyPrice: 280, deposit: 2500, status: 'disponible', insuranceExpiry: daysFromNow(180), technicalControlExpiry: daysFromNow(15) },
      { brand: 'Volkswagen', model: 'Golf', year: 2021, plate: '67890-C-3', category: 'Berline', color: 'Noir', mileage: 48000, fuelType: 'diesel', transmission: 'automatique', dailyPrice: 400, deposit: 3500, status: 'louee', insuranceExpiry: daysFromNow(90), technicalControlExpiry: daysFromNow(120) },
      { brand: 'Hyundai', model: 'Tucson', year: 2023, plate: '11223-D-4', category: 'SUV', color: 'Bleu', mileage: 9000, fuelType: 'essence', transmission: 'automatique', dailyPrice: 550, deposit: 5000, status: 'disponible', insuranceExpiry: daysFromNow(300), technicalControlExpiry: daysFromNow(300) },
      { brand: 'Mercedes', model: 'Classe C', year: 2022, plate: '99887-E-5', category: 'Luxe', color: 'Noir', mileage: 22000, fuelType: 'diesel', transmission: 'automatique', dailyPrice: 900, deposit: 8000, status: 'maintenance', insuranceExpiry: daysFromNow(60), technicalControlExpiry: daysFromNow(60) },
      { brand: 'Fiat', model: '500', year: 2020, plate: '44556-F-6', category: 'Citadine', color: 'Rouge', mileage: 61000, fuelType: 'essence', transmission: 'manuelle', dailyPrice: 220, deposit: 1800, status: 'disponible', insuranceExpiry: daysFromNow(10), technicalControlExpiry: daysFromNow(400) },
    ]);
    console.log(`${vehicles.length} véhicules de démonstration créés.`);

    const clients = await Client.bulkCreate([
      { firstName: 'Karim', lastName: 'El Amrani', phone: '0612345678', email: 'karim.amrani@example.com', address: 'Rabat', cin: 'AB123456', licenseNumber: 'PL001122', licenseExpiry: daysFromNow(500) },
      { firstName: 'Sofia', lastName: 'Bennani', phone: '0623456789', email: 'sofia.bennani@example.com', address: 'Casablanca', cin: 'AC234567', licenseNumber: 'PL002233', licenseExpiry: daysFromNow(700) },
      { firstName: 'Yassine', lastName: 'Idrissi', phone: '0634567890', email: 'yassine.idrissi@example.com', address: 'Marrakech', cin: 'AD345678', licenseNumber: 'PL003344', licenseExpiry: daysFromNow(25) },
      { firstName: 'Nadia', lastName: 'Chraibi', phone: '0645678901', email: 'nadia.chraibi@example.com', address: 'Fès', cin: 'AE456789', licenseNumber: 'PL004455', licenseExpiry: daysFromNow(900) },
    ]);
    console.log(`${clients.length} clients de démonstration créés.`);

    const golf = vehicles[2];
    const reservation1 = await Reservation.create({
      clientId: clients[0].id,
      vehicleId: golf.id,
      startDate: daysFromNow(-2),
      endDate: daysFromNow(3),
      days: 5,
      totalPrice: 5 * golf.dailyPrice,
      status: 'confirmee',
    });

    await Contract.create({
      reservationId: reservation1.id,
      clientId: clients[0].id,
      vehicleId: golf.id,
      contractNumber: generateNumber('CTR'),
      startDate: reservation1.startDate,
      endDate: reservation1.endDate,
      totalPrice: reservation1.totalPrice,
      deposit: golf.deposit,
      terms: settings.contractTerms,
      signedClient: true,
      signedAgency: true,
    });

    await Payment.create({
      reservationId: reservation1.id,
      clientId: clients[0].id,
      receiptNumber: generateNumber('REC'),
      amount: 1000,
      method: 'especes',
      depositPaid: true,
      status: 'partiel',
      remaining: reservation1.totalPrice - 1000,
    });

    await Reservation.create({
      clientId: clients[1].id,
      vehicleId: vehicles[1].id,
      startDate: daysFromNow(5),
      endDate: daysFromNow(10),
      days: 5,
      totalPrice: 5 * vehicles[1].dailyPrice,
      status: 'en_attente',
    });

    const loganPastReservation = await Reservation.create({
      clientId: clients[2].id,
      vehicleId: vehicles[0].id,
      startDate: daysFromNow(-10),
      endDate: daysFromNow(-6),
      days: 4,
      totalPrice: 4 * vehicles[0].dailyPrice,
      status: 'terminee',
    });

    await Return.create({
      reservationId: loganPastReservation.id,
      vehicleId: vehicles[0].id,
      returnDate: new Date(daysFromNow(-6)),
      mileageReturn: vehicles[0].mileage + 350,
      fuelReturn: '3/4',
      condition: 'Bon état',
      lateFee: 0,
      fuelFee: 50,
      damageFee: 0,
      extraMileageFee: 0,
      totalExtra: 50,
    });

    await Payment.create({
      reservationId: loganPastReservation.id,
      clientId: clients[2].id,
      receiptNumber: generateNumber('REC'),
      amount: loganPastReservation.totalPrice,
      method: 'carte',
      depositPaid: true,
      status: 'paye',
      remaining: 0,
    });

    await Maintenance.create({
      vehicleId: vehicles[4].id,
      type: 'reparation',
      date: daysFromNow(-3),
      cost: 1200,
      description: 'Remplacement plaquettes de frein',
    });

    console.log('Réservations, contrat, paiements, retour et maintenance de démonstration créés.');
  }

  console.log('Seed terminé avec succès.');
}

module.exports = { seedDemoData };
