Attribute VB_Name = "modClients"
Option Explicit

Public Sub Client_Ajouter()
    Dim wsForm As Worksheet, lo As ListObject, row As ListRow
    Set wsForm = ThisWorkbook.Worksheets(SH_FORM_CLIENT)
    Set lo = GetTable(SH_CLIENTS, TB_CLIENTS)

    If Trim$(wsForm.Range("B4").Value) = "" Then
        MsgBox "CIN obligatoire.", vbExclamation
        Exit Sub
    End If

    Set row = lo.ListRows.Add
    row.Range.Cells(1, lo.ListColumns("ClientID").Index).Value = NextId(lo, "ClientID")
    row.Range.Cells(1, lo.ListColumns("CIN").Index).Value = wsForm.Range("B4").Value
    row.Range.Cells(1, lo.ListColumns("Nom").Index).Value = wsForm.Range("B5").Value
    row.Range.Cells(1, lo.ListColumns("Prenom").Index).Value = wsForm.Range("B6").Value
    row.Range.Cells(1, lo.ListColumns("Telephone").Index).Value = wsForm.Range("B7").Value
    row.Range.Cells(1, lo.ListColumns("Adresse").Index).Value = wsForm.Range("B8").Value
    row.Range.Cells(1, lo.ListColumns("PermisNumero").Index).Value = wsForm.Range("B9").Value
    row.Range.Cells(1, lo.ListColumns("PermisExpiration").Index).Value = DateSafe(wsForm.Range("B10").Value)
    row.Range.Cells(1, lo.ListColumns("DateAjout").Index).Value = Date

    MsgBox "Client ajouté.", vbInformation
End Sub

Public Sub Client_Modifier()
    Dim wsForm As Worksheet, lo As ListObject, r As ListRow
    Set wsForm = ThisWorkbook.Worksheets(SH_FORM_CLIENT)
    Set lo = GetTable(SH_CLIENTS, TB_CLIENTS)

    If Trim$(wsForm.Range("B3").Value) = "" Then
        MsgBox "ClientID obligatoire pour modifier.", vbExclamation
        Exit Sub
    End If

    Set r = FindRowByValue(lo, "ClientID", wsForm.Range("B3").Value)
    If r Is Nothing Then
        MsgBox "Client introuvable.", vbCritical
        Exit Sub
    End If

    r.Range.Cells(1, lo.ListColumns("CIN").Index).Value = wsForm.Range("B4").Value
    r.Range.Cells(1, lo.ListColumns("Nom").Index).Value = wsForm.Range("B5").Value
    r.Range.Cells(1, lo.ListColumns("Prenom").Index).Value = wsForm.Range("B6").Value
    r.Range.Cells(1, lo.ListColumns("Telephone").Index).Value = wsForm.Range("B7").Value
    r.Range.Cells(1, lo.ListColumns("Adresse").Index).Value = wsForm.Range("B8").Value
    r.Range.Cells(1, lo.ListColumns("PermisNumero").Index).Value = wsForm.Range("B9").Value
    r.Range.Cells(1, lo.ListColumns("PermisExpiration").Index).Value = DateSafe(wsForm.Range("B10").Value)

    MsgBox "Client modifié.", vbInformation
End Sub

Public Sub Client_Supprimer()
    Dim wsForm As Worksheet, lo As ListObject, r As ListRow
    Set wsForm = ThisWorkbook.Worksheets(SH_FORM_CLIENT)
    Set lo = GetTable(SH_CLIENTS, TB_CLIENTS)

    If Trim$(wsForm.Range("B3").Value) = "" Then
        MsgBox "ClientID obligatoire pour supprimer.", vbExclamation
        Exit Sub
    End If

    Set r = FindRowByValue(lo, "ClientID", wsForm.Range("B3").Value)
    If r Is Nothing Then
        MsgBox "Client introuvable.", vbCritical
        Exit Sub
    End If

    If MsgBox("Supprimer ce client ?", vbQuestion + vbYesNo) = vbYes Then
        r.Delete
        MsgBox "Client supprimé.", vbInformation
    End If
End Sub
