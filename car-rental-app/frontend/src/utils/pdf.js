import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { format } from 'date-fns';
import logoUrl from '../assets/logo.png';

const METHOD_LABELS = { especes: 'Espèces', carte: 'Carte bancaire', virement: 'Virement', cheque: 'Chèque' };

let cachedLogoDataUrl = null;

async function getLogoDataUrl() {
  if (cachedLogoDataUrl) return cachedLogoDataUrl;
  try {
    const res = await fetch(logoUrl);
    const blob = await res.blob();
    cachedLogoDataUrl = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  } catch {
    cachedLogoDataUrl = null;
  }
  return cachedLogoDataUrl;
}

async function drawHeader(doc, settings, title) {
  doc.setFillColor(15, 23, 42);
  doc.rect(0, 0, 210, 30, 'F');

  const logoDataUrl = await getLogoDataUrl();
  const textX = logoDataUrl ? 30 : 14;
  if (logoDataUrl) {
    try {
      doc.addImage(logoDataUrl, 'PNG', 12, 6, 18, 18);
    } catch {
      // ignore logo rendering issues, header text still renders
    }
  }

  doc.setTextColor(255, 255, 255);
  doc.setFontSize(16);
  doc.setFont('helvetica', 'bold');
  doc.text(settings?.name || 'Nova Motion Car', textX, 15);
  doc.setFontSize(9);
  doc.setFont('helvetica', 'normal');
  doc.text(
    [settings?.address, settings?.phone, settings?.email].filter(Boolean).join(' · '),
    textX,
    22
  );
  doc.setFontSize(13);
  doc.setFont('helvetica', 'bold');
  doc.text(title, 196, 17, { align: 'right' });
  doc.setTextColor(0, 0, 0);
}

export async function generateContractPDF(contract, settings) {
  const doc = new jsPDF();
  const currency = settings?.currency || 'MAD';
  await drawHeader(doc, settings, 'CONTRAT DE LOCATION');

  doc.setFontSize(10);
  doc.text(`Contrat N° : ${contract.contractNumber}`, 14, 40);
  doc.text(`Date d'émission : ${format(new Date(contract.createdAt || Date.now()), 'dd/MM/yyyy')}`, 14, 46);

  autoTable(doc, {
    startY: 54,
    head: [['Informations client', '']],
    body: [
      ['Nom complet', `${contract.client?.firstName} ${contract.client?.lastName}`],
      ['Téléphone', contract.client?.phone || '—'],
      ['CIN / Passeport', contract.client?.cin || '—'],
      ['N° Permis', contract.client?.licenseNumber || '—'],
    ],
    theme: 'grid',
    headStyles: { fillColor: [47, 95, 237] },
  });

  autoTable(doc, {
    startY: doc.lastAutoTable.finalY + 6,
    head: [['Informations véhicule', '']],
    body: [
      ['Véhicule', `${contract.vehicle?.brand} ${contract.vehicle?.model} (${contract.vehicle?.year || ''})`],
      ['Immatriculation', contract.vehicle?.plate || '—'],
      ['Catégorie', contract.vehicle?.category || '—'],
    ],
    theme: 'grid',
    headStyles: { fillColor: [47, 95, 237] },
  });

  autoTable(doc, {
    startY: doc.lastAutoTable.finalY + 6,
    head: [['Détails de la location', '']],
    body: [
      ['Date de début', format(new Date(contract.startDate), 'dd/MM/yyyy')],
      ['Date de fin', format(new Date(contract.endDate), 'dd/MM/yyyy')],
      ['Prix total', `${contract.totalPrice} ${currency}`],
      ['Caution', `${contract.deposit} ${currency}`],
    ],
    theme: 'grid',
    headStyles: { fillColor: [47, 95, 237] },
  });

  const termsY = doc.lastAutoTable.finalY + 10;
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(11);
  doc.text('Conditions générales', 14, termsY);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(9);
  const terms = doc.splitTextToSize(contract.terms || '', 182);
  doc.text(terms, 14, termsY + 6);

  const signY = termsY + 6 + terms.length * 4.2 + 20;
  doc.line(14, signY, 80, signY);
  doc.text('Signature du client', 14, signY + 5);
  doc.line(130, signY, 196, signY);
  doc.text("Signature de l'agence", 130, signY + 5);

  doc.save(`contrat-${contract.contractNumber}.pdf`);
}

export async function generateReceiptPDF(payment, settings) {
  const doc = new jsPDF();
  const currency = settings?.currency || 'MAD';
  await drawHeader(doc, settings, 'REÇU DE PAIEMENT');

  doc.setFontSize(10);
  doc.text(`Reçu N° : ${payment.receiptNumber}`, 14, 40);
  doc.text(`Date : ${format(new Date(payment.paidAt || Date.now()), 'dd/MM/yyyy HH:mm')}`, 14, 46);

  autoTable(doc, {
    startY: 54,
    body: [
      ['Client', `${payment.client?.firstName} ${payment.client?.lastName}`],
      ['Véhicule', payment.reservation?.vehicle ? `${payment.reservation.vehicle.brand} ${payment.reservation.vehicle.model}` : '—'],
      ['Mode de paiement', METHOD_LABELS[payment.method] || payment.method],
      ['Montant payé', `${payment.amount} ${currency}`],
      ['Reste à payer', `${payment.remaining} ${currency}`],
      ['Statut', payment.status === 'paye' ? 'Payé intégralement' : payment.status === 'partiel' ? 'Paiement partiel' : 'Impayé'],
    ],
    theme: 'grid',
    headStyles: { fillColor: [47, 95, 237] },
  });

  doc.setFontSize(9);
  doc.text('Merci pour votre confiance.', 14, doc.lastAutoTable.finalY + 12);

  doc.save(`recu-${payment.receiptNumber}.pdf`);
}
