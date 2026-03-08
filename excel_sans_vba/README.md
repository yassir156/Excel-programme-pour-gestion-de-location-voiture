# Version sans VBA (100% feuilles Excel)

Ce dossier décrit le même mini logiciel de location de voitures, **sans macro**, uniquement avec:
- feuilles structurées,
- tableaux Excel,
- validations de données,
- formules,
- mise en forme conditionnelle.

## 1) Feuilles à créer

1. `PARAMETRES`
2. `VEHICULES`
3. `CLIENTS`
4. `LOCATIONS`
5. `PAIEMENTS`
6. `ENTRETIEN` (optionnelle)
7. `RECHERCHE`
8. `DASHBOARD`
9. `RAPPORTS`

## 2) PARAMETRES

Créer ces cellules:
- `B2` = Agence (ex: `Sefrou`)
- `B3` = Devise (`DH`)
- `B4` = Format date (`jj/mm/aaaa`)
- `B5` = Remise auto longue durée % (ex: `10`)
- `B6` = Seuil jours remise auto (ex: `30`)

## 3) VEHICULES (Tableau `T_Vehicules`)

Colonnes:
- `VehiculeID`
- `Immatriculation`
- `Marque`
- `Modele`
- `Annee`
- `Km`
- `Carburant`
- `PrixJourDH`
- `Statut` (Disponible, Réservée, Louée, Maintenance)
- `DateAjout`
- `CA_Total_DH` (calcul)
- `Total_Paye_DH` (calcul)
- `Reste_DH` (calcul)

Formules (ligne 2 du tableau, puis recopier):
- `CA_Total_DH`:
  - `=SOMME.SI.ENS(T_Locations[MontantNet];T_Locations[VehiculeID];[@VehiculeID])`
- `Total_Paye_DH`:
  - `=SOMME.SI.ENS(T_Locations[TotalPaye];T_Locations[VehiculeID];[@VehiculeID])`
- `Reste_DH`:
  - `=[@CA_Total_DH]-[@Total_Paye_DH]`

## 4) CLIENTS (Tableau `T_Clients`)

Colonnes:
- `ClientID`
- `CIN`
- `Nom`
- `Prenom`
- `Telephone`
- `Adresse`
- `PermisNumero`
- `PermisExpiration`
- `DateAjout`

## 5) LOCATIONS (Tableau `T_Locations`)

Colonnes:
- `LocationID`
- `NumeroContrat`
- `ClientID`
- `VehiculeID`
- `DateDebut`
- `DateFinPrevue`
- `DateRetourReelle`
- `NbJours`
- `PrixJourDH`
- `RemisePctManuelle`
- `RemisePctAuto`
- `RemisePctFinale`
- `MontantBrut`
- `MontantRemise`
- `MontantNet`
- `TotalPaye`
- `ResteAPayer`
- `Statut` (RESERVATION, DEPART, RETOUR, PROLONGATION, ANNULATION)
- `EtatDepart`
- `EtatRetour`
- `DateCreation`

Formules:
- `NumeroContrat`:
  - `="CTR-"&TEXTE([@DateCreation];"aaaamm")&"-"&TEXTE([@LocationID];"0000")`
- `NbJours`:
  - `=MAX(1;[@DateFinPrevue]-[@DateDebut])`
- `RemisePctAuto`:
  - `=SI([@NbJours]>=PARAMETRES!$B$6;PARAMETRES!$B$5;0)`
- `RemisePctFinale`:
  - `=MAX([@RemisePctManuelle];[@RemisePctAuto])`
- `MontantBrut`:
  - `=[@NbJours]*[@PrixJourDH]`
- `MontantRemise`:
  - `=[@MontantBrut]*[@RemisePctFinale]/100`
- `MontantNet`:
  - `=[@MontantBrut]-[@MontantRemise]`
- `TotalPaye`:
  - `=SOMME.SI.ENS(T_Paiements[MontantDH];T_Paiements[LocationID];[@LocationID])`
- `ResteAPayer`:
  - `=[@MontantNet]-[@TotalPaye]`
- `DateCreation`:
  - `=SI([@LocationID]="";"";AUJOURDHUI())`

## 6) PAIEMENTS (Tableau `T_Paiements`)

Colonnes:
- `PaiementID`
- `LocationID`
- `DatePaiement`
- `ModePaiement` (Espèces, Virement, Carte, Chèque)
- `MontantDH`
- `Reference`

## 7) ENTRETIEN (Tableau `T_Entretien`, optionnel)

Colonnes:
- `EntretienID`
- `VehiculeID`
- `Type`
- `DateOperation`
- `DateProchaine`
- `KmOperation`
- `CoutDH`
- `Notes`
- `Alerte`

Formule `Alerte`:
- `=SI(ET([@DateProchaine]<>"";AUJOURDHUI()>=[@DateProchaine]);"ROUGE";"OK")`

## 8) RECHERCHE

Créer zone critères en `A3:B7`:
- CIN
- Nom client
- Immatriculation
- N° Contrat
- Date début

Puis afficher les résultats avec `FILTRE` (Excel 365) dans `D5`:

```excel
=FILTRE(T_Locations[[NumeroContrat]:[Statut]];
 (SI($B$3="";VRAI;ESTNUM(CHERCHE($B$3;RECHERCHEX(T_Locations[ClientID];T_Clients[ClientID];T_Clients[CIN];""))))) *
 (SI($B$4="";VRAI;ESTNUM(CHERCHE($B$4;RECHERCHEX(T_Locations[ClientID];T_Clients[ClientID];T_Clients[Nom];""))))) *
 (SI($B$5="";VRAI;ESTNUM(CHERCHE($B$5;RECHERCHEX(T_Locations[VehiculeID];T_Vehicules[VehiculeID];T_Vehicules[Immatriculation];""))))) *
 (SI($B$6="";VRAI;ESTNUM(CHERCHE($B$6;T_Locations[NumeroContrat])))) *
 (SI($B$7="";VRAI;T_Locations[DateDebut]=$B$7));
 "Aucun résultat")
```

## 9) DASHBOARD

Créer des cartes:
- CA total: `=SOMME(T_Locations[MontantNet])`
- Total payé: `=SOMME(T_Locations[TotalPaye])`
- Reste à payer: `=SOMME(T_Locations[ResteAPayer])`
- Locations actives: `=NB.SI(T_Locations[Statut];"DEPART")+NB.SI(T_Locations[Statut];"PROLONGATION")`
- Réservations: `=NB.SI(T_Locations[Statut];"RESERVATION")`
- Retards:
  - `=SOMMEPROD((T_Locations[Statut]="DEPART")*(T_Locations[DateFinPrevue]<AUJOURDHUI()))+SOMMEPROD((T_Locations[Statut]="PROLONGATION")*(T_Locations[DateFinPrevue]<AUJOURDHUI()))`

Ajouter 2 tableaux croisés dynamiques:
- CA par véhicule.
- Paiements mensuels.

## 10) Alertes rouges (mise en forme conditionnelle)

1. `LOCATIONS[ResteAPayer] > 0` -> fond rouge clair.
2. `LOCATIONS[DateFinPrevue] < AUJOURDHUI()` avec statut `DEPART`/`PROLONGATION` -> rouge.
3. `CLIENTS[PermisExpiration] <= AUJOURDHUI()+30` -> orange/rouge.
4. `ENTRETIEN[Alerte]="ROUGE"` -> rouge.

## 11) Flux de travail sans macro

1. Ajouter véhicules + clients.
2. Saisir location (ID, client, véhicule, dates, prix/jour, remise manuelle).
3. Mettre statut à jour manuellement (réservation -> départ -> retour / prolongation / annulation).
4. Saisir paiements dans `PAIEMENTS`.
5. Consulter recherche, dashboard et rapports.

## 12) Conseils performance (PC moyen)

- Éviter les colonnes entières dans les formules.
- Utiliser uniquement des tableaux structurés.
- Limiter les mises en forme lourdes.
- Archiver les anciennes années dans un autre fichier.
