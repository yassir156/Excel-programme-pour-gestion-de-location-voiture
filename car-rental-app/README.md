# Car Rental Manager — Application desktop de gestion de location de voitures

Application desktop complète (Windows/macOS/Linux) pour la gestion d'une agence de location de voitures : véhicules, clients, réservations, contrats, paiements, retours, maintenance, statistiques et administration.

## Stack technique

| Domaine | Technologie |
|---|---|
| Application desktop | Electron.js |
| Interface | React.js + Tailwind CSS |
| Backend local | Node.js / Express.js (démarré automatiquement par Electron) |
| Base de données | SQLite (fichier local, aucun serveur externe requis) |
| ORM | Sequelize |
| Génération PDF | jsPDF + jspdf-autotable (contrats et reçus) |
| Export Excel / CSV | SheetJS (xlsx) |
| Graphiques | Recharts |
| Packaging | Electron Builder (.exe NSIS pour Windows) |

## Architecture du projet

```
car-rental-app/
├── electron/                  # Processus principal Electron
│   ├── main.js                 # Démarre le serveur Express embarqué + la fenêtre
│   └── preload.js               # Pont sécurisé entre Electron et le renderer
├── backend/
│   ├── src/
│   │   ├── models/              # Modèles Sequelize (User, Vehicle, Client, Reservation, ...)
│   │   ├── routes/               # Routes REST Express (une par ressource)
│   │   ├── services/             # Logique métier (disponibilité, calcul de prix, alertes)
│   │   ├── middleware/           # Authentification JWT, gestion des rôles, erreurs
│   │   ├── utils/                 # JWT, numérotation contrats/reçus
│   │   ├── app.js                 # Application Express
│   │   ├── db.js                   # Connexion SQLite
│   │   └── server.js               # Démarrage du serveur HTTP local
│   └── seed/                    # Données de démonstration
├── frontend/                  # Application React (Vite + Tailwind)
│   └── src/
│       ├── api/                  # Client Axios
│       ├── context/               # Authentification, thème clair/sombre
│       ├── components/            # Sidebar, Header, DataTable, Modal, StatCard...
│       ├── pages/                  # Toutes les pages de l'application
│       └── utils/                  # Génération PDF, export Excel/CSV
└── package.json                # Configuration racine + Electron Builder
```

La base de données SQLite et les fichiers uploadés (photos, documents, logo) sont stockés dans le dossier de données utilisateur de l'application (`app.getPath('userData')` une fois installée), afin de survivre aux mises à jour.

## Fonctionnalités

- **Authentification & rôles** : administrateur, manager, agent — mots de passe chiffrés (bcrypt), sessions JWT.
- **Tableau de bord** : statistiques clés, graphiques (revenus, répartition de la flotte), prochaines réservations/retours, alertes.
- **Véhicules** : fiche complète, photos, statut, alertes assurance/contrôle technique, historique des locations.
- **Clients** : fiche complète, documents (CIN, permis), historique des locations.
- **Réservations** : vérification automatique de disponibilité (anti double-réservation), calcul automatique des jours/prix, calendrier des statuts.
- **Contrats** : génération automatique depuis une réservation, export PDF, suivi des signatures.
- **Paiements** : espèces / carte / virement / chèque, suivi du reste à payer, reçu PDF.
- **Retours** : kilométrage, carburant, état, frais supplémentaires calculés automatiquement, mise à jour du statut du véhicule.
- **Maintenance** : suivi des opérations, alertes d'échéance.
- **Statistiques** : revenus mensuels, véhicules/clients les plus actifs, taux d'occupation, exports PDF / Excel / CSV.
- **Paramètres** : informations de l'agence, devise, taxes, conditions générales, sauvegarde/restauration de la base de données.

## Installation (développement)

Prérequis : [Node.js](https://nodejs.org/) ≥ 18.

```bash
cd car-rental-app
npm install                 # installe les dépendances Electron + backend
npm run frontend:install    # installe les dépendances du frontend React
```

## Lancer l'application en mode développement

```bash
npm run dev
```

Cette commande démarre le serveur Vite (frontend) puis lance Electron, qui démarre lui-même le serveur Express local et charge l'interface. Au premier lancement, la base de données est créée et peuplée automatiquement avec des données de démonstration.

### Comptes de démonstration

| Identifiant | Mot de passe | Rôle |
|---|---|---|
| `admin` | `admin123` | Administrateur |
| `manager` | `manager123` | Manager |
| `agent` | `agent123` | Agent |

### Peupler manuellement la base de données de démonstration

```bash
npm run backend:seed
```

## Générer l'exécutable Windows (.exe)

```bash
npm run frontend:build   # build de production du frontend React
npm run dist              # génère l'installeur Windows (.exe) via electron-builder
```

L'installeur NSIS est généré dans le dossier `release/`. Il permet à l'utilisateur de choisir le dossier d'installation et crée des raccourcis Bureau / Menu Démarrer.

> **Remarque** : la génération d'un `.exe` signé nécessite d'exécuter `electron-builder` sur Windows (ou avec Wine sur Linux/macOS pour la génération croisée). Sur Linux, utilisez `npm run dist:linux` pour générer un AppImage de test ; sur macOS, `npm run dist:mac`.

Le module natif `sqlite3` est automatiquement recompilé pour la version d'Electron utilisée via le script `postinstall` (`electron-builder install-app-deps`).

## Sauvegarde et restauration

Dans **Paramètres → Sauvegarde et restauration** (réservé aux administrateurs), il est possible de télécharger une copie du fichier SQLite ou d'en restaurer une. Un redémarrage de l'application est nécessaire après une restauration.

## Sécurité

- Mots de passe hachés avec bcrypt (jamais stockés en clair).
- Authentification par jeton JWT signé avec une clé générée aléatoirement et stockée localement.
- Le serveur backend n'écoute que sur `127.0.0.1` (aucun accès réseau externe).
- Contrôle d'accès par rôle sur les actions sensibles (suppression, gestion des utilisateurs, paramètres).
