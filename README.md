# Mini logiciel Excel — Gestion de location de voitures (Maroc)

Ce dépôt contient **2 versions** du même projet:

1. **Version VBA légère** (`/vba`) pour automatiser setup + CRUD + dashboard.
2. **Version sans VBA** (`/excel_sans_vba`) 100% feuilles Excel + formules + TCD.

---

## A) Version VBA légère

Cette version fournit une base complète pour Excel Windows (mono-utilisateur):
- CRUD clients et véhicules,
- cycle réservation / départ / retour / prolongation / annulation,
- paiements, impayés, retards,
- recherche multi-critères,
- dashboard, alertes rouges,
- suivi entretien (optionnel).

### Installation rapide (VBA)

1. Créer `LocationVoiture.xlsm`.
2. Ouvrir VBA (`ALT+F11`).
3. Importer tous les `.bas` du dossier `vba/`.
4. Exécuter `Setup_InitialiserClasseur`.
5. Associer les macros aux boutons des feuilles formulaires.

---

## B) Version sans VBA (recommandée si vous voulez uniquement des feuilles)

Le guide complet est dans:

- `excel_sans_vba/README.md`

Ce guide inclut:
- structure des feuilles,
- noms de tableaux,
- colonnes,
- formules prêtes à coller,
- mise en forme conditionnelle (alertes rouges),
- dashboard,
- recherche avec `FILTRE`/`RECHERCHEX`.

Des modèles d'en-têtes CSV sont fournis dans:
- `excel_sans_vba/templates/`

---

## Paramètres métier (Maroc)

- Agence: configurable (`Sefrou` par défaut)
- Devise: `DH`
- Date: `jj/mm/aaaa`
- Remise auto longue durée: configurable (%) + seuil jours.

## Conseils

- Si vous êtes seul et souhaitez un fichier simple: commencez par la version **sans VBA**.
- Si vous voulez plus d'automatisation: activez la version **VBA légère**.

---

## Fichiers XLS livrés (prêts à ouvrir)

Des feuilles Excel compatibles sont fournies dans:
- `livrables_xls/modele_gestion_location_voiture.xls` (classeur multi-feuilles)
- `livrables_xls/PARAMETRES.xls`
- `livrables_xls/VEHICULES.xls`
- `livrables_xls/CLIENTS.xls`
- `livrables_xls/LOCATIONS.xls`
- `livrables_xls/PAIEMENTS.xls`
- `livrables_xls/ENTRETIEN.xls`
- `livrables_xls/RECHERCHE.xls`
- `livrables_xls/DASHBOARD.xls`
- `livrables_xls/RAPPORTS.xls`

> Ces fichiers sont fournis pour démarrage rapide "feuilles XLS" et s'ouvrent dans Excel Windows.
