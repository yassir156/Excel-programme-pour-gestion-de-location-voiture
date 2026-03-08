# Mini logiciel Excel/VBA — Gestion de location de voitures (Maroc)

Ce projet fournit une base **complète et légère** pour construire un mini logiciel de location de voitures sous Excel (Windows + VBA), en français, avec montants en DH et dates `jj/mm/aaaa`.

## Fonctionnalités couvertes

- Gestion des **véhicules** (CRUD).
- Gestion des **clients** (CRUD) avec CIN, permis, téléphone, etc.
- Gestion des **réservations / locations**:
  - réservation,
  - départ,
  - retour,
  - prolongation,
  - annulation.
- État véhicule au départ et au retour.
- Tarification au jour + réduction:
  - remise manuelle,
  - remise automatique long séjour (seuil configurable).
- Suivi des **paiements** (avances, soldes, reste à payer).
- Recherche multi-critères: CIN, nom client, immatriculation, numéro contrat, date.
- Dashboard: CA, payé, reste, nombre de locations, statut.
- Alertes visuelles (rouge) sur retards, impayés et expirations entretien.
- Génération automatique de toutes les feuilles via macro `Setup_InitialiserClasseur`.

## Structure des feuilles générées

- `CONFIG` : paramètres globaux.
- `VEHICULES` : base des voitures.
- `CLIENTS` : base clients.
- `LOCATIONS` : contrats et cycle de vie (réservation -> retour).
- `PAIEMENTS` : journal des paiements.
- `ENTRETIEN` : suivi maintenance (optionnel).
- `RECHERCHE` : moteur de recherche.
- `DASHBOARD` : indicateurs.
- `FORM_CLIENT` : formulaire CRUD client (sur feuille).
- `FORM_VEHICULE` : formulaire CRUD véhicule (sur feuille).
- `FORM_LOCATION` : formulaire CRUD location (sur feuille).

## Installation

1. Créer un fichier Excel macro-enabled: `LocationVoiture.xlsm`.
2. Ouvrir l'éditeur VBA (`ALT+F11`).
3. Importer tous les fichiers `.bas` du dossier `vba/`.
4. Lancer la macro: `Setup_InitialiserClasseur`.
5. Ajouter des boutons sur les feuilles de formulaires:
   - `Client_Ajouter`, `Client_Modifier`, `Client_Supprimer`,
   - `Vehicule_Ajouter`, `Vehicule_Modifier`, `Vehicule_Supprimer`,
   - `Location_Ajouter`, `Location_Modifier`, `Location_Annuler`, `Location_Depart`, `Location_Retour`, `Location_Prolonger`,
   - `Paiement_AjouterDepuisForm`,
   - `Recherche_Lancer`, `Dashboard_Refresh`.

## Paramètres à ajuster (`CONFIG`)

- `B2`: Agence (ex: Sefrou).
- `B3`: Devise (`DH`).
- `B4`: Format date (`jj/mm/aaaa`).
- `B5`: Réduction auto (%) longue durée (ex: `10`).
- `B6`: Seuil jours réduction auto (ex: `30`).

## Flux conseillé d'utilisation

1. Ajouter les véhicules dans `FORM_VEHICULE`.
2. Ajouter les clients dans `FORM_CLIENT`.
3. Créer une réservation/location depuis `FORM_LOCATION`.
4. Enregistrer les paiements à la création puis pendant le cycle.
5. Changer statut: départ -> retour/prolongation/annulation.
6. Faire `Dashboard_Refresh` pour tableau de bord et alertes.

## Remarques

- Le suivi `ENTRETIEN` est **optionnel**: aucun blocage sur les autres opérations.
- Le code est conçu pour un usage mono-utilisateur sur un poste moyen.
- Vous pouvez enrichir ensuite avec impression contrat PDF.
