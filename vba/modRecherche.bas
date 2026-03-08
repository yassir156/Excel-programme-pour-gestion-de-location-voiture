Attribute VB_Name = "modRecherche"
Option Explicit

Public Sub Recherche_Lancer()
    Dim wsR As Worksheet, lo As ListObject, outRow As Long, r As ListRow
    Dim fCin As String, fNom As String, fImmat As String, fContrat As String
    Dim fDate As Variant

    Set wsR = ThisWorkbook.Worksheets(SH_RECHERCHE)
    Set lo = GetTable(SH_LOCATIONS, TB_LOCATIONS)

    fCin = LCase$(Trim$(wsR.Range("B3").Value))
    fNom = LCase$(Trim$(wsR.Range("B4").Value))
    fImmat = LCase$(Trim$(wsR.Range("B5").Value))
    fContrat = LCase$(Trim$(wsR.Range("B6").Value))
    fDate = wsR.Range("B7").Value

    wsR.Range("C5:K5000").ClearContents
    outRow = 5

    For Each r In lo.ListRows
        If MatchLocationFilters(r, lo, fCin, fNom, fImmat, fContrat, fDate) Then
            wsR.Cells(outRow, "C").Value = r.Range.Cells(1, lo.ListColumns("NumeroContrat").Index).Value
            wsR.Cells(outRow, "D").Value = r.Range.Cells(1, lo.ListColumns("ClientID").Index).Value
            wsR.Cells(outRow, "E").Value = r.Range.Cells(1, lo.ListColumns("VehiculeID").Index).Value
            wsR.Cells(outRow, "F").Value = r.Range.Cells(1, lo.ListColumns("DateDebut").Index).Value
            wsR.Cells(outRow, "G").Value = r.Range.Cells(1, lo.ListColumns("DateFinPrevue").Index).Value
            wsR.Cells(outRow, "H").Value = r.Range.Cells(1, lo.ListColumns("MontantNet").Index).Value
            wsR.Cells(outRow, "I").Value = r.Range.Cells(1, lo.ListColumns("TotalPaye").Index).Value
            wsR.Cells(outRow, "J").Value = r.Range.Cells(1, lo.ListColumns("ResteAPayer").Index).Value
            wsR.Cells(outRow, "K").Value = r.Range.Cells(1, lo.ListColumns("Statut").Index).Value
            outRow = outRow + 1
        End If
    Next r

    MsgBox "Recherche terminée.", vbInformation
End Sub

Private Function MatchLocationFilters(ByVal r As ListRow, ByVal lo As ListObject, _
                                      ByVal cinFilter As String, ByVal nomFilter As String, _
                                      ByVal immatFilter As String, ByVal contratFilter As String, _
                                      ByVal dateFilter As Variant) As Boolean
    Dim clientId As Variant, vehiculeId As Variant
    Dim cinValue As String, nomValue As String, immatValue As String, contratValue As String
    Dim dateDebut As Variant

    clientId = r.Range.Cells(1, lo.ListColumns("ClientID").Index).Value
    vehiculeId = r.Range.Cells(1, lo.ListColumns("VehiculeID").Index).Value
    contratValue = LCase$(CStr(r.Range.Cells(1, lo.ListColumns("NumeroContrat").Index).Value))
    dateDebut = r.Range.Cells(1, lo.ListColumns("DateDebut").Index).Value

    cinValue = LCase$(GetClientValue(clientId, "CIN"))
    nomValue = LCase$(GetClientValue(clientId, "Nom"))
    immatValue = LCase$(GetVehiculeValue(vehiculeId, "Immatriculation"))

    MatchLocationFilters = True

    If cinFilter <> "" And InStr(1, cinValue, cinFilter, vbTextCompare) = 0 Then MatchLocationFilters = False
    If nomFilter <> "" And InStr(1, nomValue, nomFilter, vbTextCompare) = 0 Then MatchLocationFilters = False
    If immatFilter <> "" And InStr(1, immatValue, immatFilter, vbTextCompare) = 0 Then MatchLocationFilters = False
    If contratFilter <> "" And InStr(1, contratValue, contratFilter, vbTextCompare) = 0 Then MatchLocationFilters = False
    If IsDate(dateFilter) Then
        If CLng(DateValue(dateDebut)) <> CLng(DateValue(dateFilter)) Then MatchLocationFilters = False
    End If
End Function

Private Function GetClientValue(ByVal clientId As Variant, ByVal colName As String) As String
    Dim lo As ListObject, rr As ListRow
    Set lo = GetTable(SH_CLIENTS, TB_CLIENTS)
    Set rr = FindRowByValue(lo, "ClientID", clientId)
    If rr Is Nothing Then
        GetClientValue = ""
    Else
        GetClientValue = CStr(rr.Range.Cells(1, lo.ListColumns(colName).Index).Value)
    End If
End Function

Private Function GetVehiculeValue(ByVal vehiculeId As Variant, ByVal colName As String) As String
    Dim lo As ListObject, rr As ListRow
    Set lo = GetTable(SH_VEHICULES, TB_VEHICULES)
    Set rr = FindRowByValue(lo, "VehiculeID", vehiculeId)
    If rr Is Nothing Then
        GetVehiculeValue = ""
    Else
        GetVehiculeValue = CStr(rr.Range.Cells(1, lo.ListColumns(colName).Index).Value)
    End If
End Function
