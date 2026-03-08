Attribute VB_Name = "modSetup"
Option Explicit

Public Sub Setup_InitialiserClasseur()
    Application.ScreenUpdating = False

    Setup_Config
    Setup_Vehicules
    Setup_Clients
    Setup_Locations
    Setup_Paiements
    Setup_Entretien
    Setup_Recherche
    Setup_Dashboard
    Setup_FormsSheets
    Setup_ConditionalFormatting

    Application.ScreenUpdating = True
    MsgBox "Initialisation terminée.", vbInformation
End Sub

Private Sub Setup_Config()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SH_CONFIG)
    ws.Cells.Clear

    ws.Range("A1").Value = "Paramètre"
    ws.Range("B1").Value = "Valeur"

    ws.Range("A2").Value = "Agence"
    ws.Range("B2").Value = "Sefrou"

    ws.Range("A3").Value = "Devise"
    ws.Range("B3").Value = "DH"

    ws.Range("A4").Value = "Format date"
    ws.Range("B4").Value = "jj/mm/aaaa"

    ws.Range("A5").Value = "Remise auto longue durée (%)"
    ws.Range("B5").Value = 10

    ws.Range("A6").Value = "Seuil jours remise auto"
    ws.Range("B6").Value = 30

    ws.Columns("A:B").AutoFit
End Sub

Private Sub Setup_Vehicules()
    Dim ws As Worksheet, headers As Variant
    Set ws = GetOrCreateSheet(SH_VEHICULES)
    ws.Cells.Clear
    headers = Array("VehiculeID", "Immatriculation", "Marque", "Modele", "Annee", "Km", "Carburant", "PrixJourDH", "Statut", "DateAjout")
    EnsureTable ws, TB_VEHICULES, headers
End Sub

Private Sub Setup_Clients()
    Dim ws As Worksheet, headers As Variant
    Set ws = GetOrCreateSheet(SH_CLIENTS)
    ws.Cells.Clear
    headers = Array("ClientID", "CIN", "Nom", "Prenom", "Telephone", "Adresse", "PermisNumero", "PermisExpiration", "DateAjout")
    EnsureTable ws, TB_CLIENTS, headers
End Sub

Private Sub Setup_Locations()
    Dim ws As Worksheet, headers As Variant
    Set ws = GetOrCreateSheet(SH_LOCATIONS)
    ws.Cells.Clear
    headers = Array("LocationID", "NumeroContrat", "ClientID", "VehiculeID", "DateDebut", "DateFinPrevue", "DateRetourReelle", "NbJours", "PrixJourDH", "RemisePct", "MontantBrut", "MontantRemise", "MontantNet", "TotalPaye", "ResteAPayer", "Statut", "EtatDepart", "EtatRetour", "DateCreation")
    EnsureTable ws, TB_LOCATIONS, headers
End Sub

Private Sub Setup_Paiements()
    Dim ws As Worksheet, headers As Variant
    Set ws = GetOrCreateSheet(SH_PAIEMENTS)
    ws.Cells.Clear
    headers = Array("PaiementID", "LocationID", "DatePaiement", "ModePaiement", "MontantDH", "Reference")
    EnsureTable ws, TB_PAIEMENTS, headers
End Sub

Private Sub Setup_Entretien()
    Dim ws As Worksheet, headers As Variant
    Set ws = GetOrCreateSheet(SH_ENTRETIEN)
    ws.Cells.Clear
    headers = Array("EntretienID", "VehiculeID", "Type", "DateOperation", "DateProchaine", "KmOperation", "CoutDH", "Notes", "Alerte")
    EnsureTable ws, TB_ENTRETIEN, headers
End Sub

Private Sub Setup_Recherche()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SH_RECHERCHE)
    ws.Cells.Clear

    ws.Range("A1").Value = "Recherche Locations"
    ws.Range("A3").Value = "CIN"
    ws.Range("A4").Value = "Nom client"
    ws.Range("A5").Value = "Immatriculation"
    ws.Range("A6").Value = "N° contrat"
    ws.Range("A7").Value = "Date début"

    ws.Range("C3").Value = "Résultats"
    ws.Range("C4:K4").Value = Array("Contrat", "ClientID", "VehiculeID", "DateDébut", "DateFin", "MontantNet", "Payé", "Reste", "Statut")
    ws.Columns.AutoFit
End Sub

Private Sub Setup_Dashboard()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SH_DASHBOARD)
    ws.Cells.Clear

    ws.Range("A1").Value = "Tableau de bord"
    ws.Range("A3").Value = "CA total (DH)"
    ws.Range("A4").Value = "Total payé (DH)"
    ws.Range("A5").Value = "Reste à payer (DH)"
    ws.Range("A6").Value = "Locations actives"
    ws.Range("A7").Value = "Réservations"
    ws.Range("A8").Value = "Retours en retard"

    ws.Range("A10").Value = "Top véhicules (CA)"
    ws.Range("A11:D11").Value = Array("VehiculeID", "Immatriculation", "Nb locations", "CA")
    ws.Columns.AutoFit
End Sub

Private Sub Setup_FormsSheets()
    Setup_FormClient
    Setup_FormVehicule
    Setup_FormLocation
End Sub

Private Sub Setup_FormClient()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SH_FORM_CLIENT)
    ws.Cells.Clear

    ws.Range("A1").Value = "Formulaire Client"
    ws.Range("A3:A10").Value = Application.Transpose(Array("ClientID (si modification)", "CIN", "Nom", "Prénom", "Téléphone", "Adresse", "Permis n°", "Permis expiration"))
    ws.Columns("A:B").AutoFit
End Sub

Private Sub Setup_FormVehicule()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SH_FORM_VEHICULE)
    ws.Cells.Clear

    ws.Range("A1").Value = "Formulaire Véhicule"
    ws.Range("A3:A11").Value = Application.Transpose(Array("VehiculeID (si modification)", "Immatriculation", "Marque", "Modèle", "Année", "KM", "Carburant", "Prix/Jour DH", "Statut"))
    ws.Columns("A:B").AutoFit
End Sub

Private Sub Setup_FormLocation()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SH_FORM_LOCATION)
    ws.Cells.Clear

    ws.Range("A1").Value = "Formulaire Location"
    ws.Range("A3:A15").Value = Application.Transpose(Array("LocationID (si modification)", "ClientID", "VehiculeID", "Date début", "Date fin prévue", "Prix/jour DH", "Remise % (manuel)", "Etat départ", "Etat retour", "Mode paiement", "Montant paiement", "Référence paiement", "Action"))
    ws.Range("D3").Value = "Actions: RESERVATION / DEPART / RETOUR / PROLONGATION / ANNULATION"
    ws.Columns("A:D").AutoFit
End Sub

Private Sub Setup_ConditionalFormatting()
    Dim ws As Worksheet
    Dim lo As ListObject

    Set ws = ThisWorkbook.Worksheets(SH_LOCATIONS)
    Set lo = ws.ListObjects(TB_LOCATIONS)

    If Not lo.DataBodyRange Is Nothing Then
        lo.DataBodyRange.FormatConditions.Delete
        With lo.ListColumns("ResteAPayer").DataBodyRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
            .Interior.Color = RGB(255, 199, 206)
            .Font.Color = RGB(156, 0, 6)
        End With
    End If
End Sub
