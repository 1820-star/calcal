Attribute VB_Name = "modKalender"
Option Explicit

'=============================================================
'  KALENDER APP 2026  -  Hauptmodul
'=============================================================

Public gSelRow As Long
Public gSelCol As Long
Public gSelRef As String

Sub OeffneTerminDialog(ByVal Target As Range)
    If Target.Row < 2 Or Target.Column < 2 Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub
    
    gSelRow = Target.Row
    gSelCol = Target.Column
    gSelRef = Target.Address(False, False)
    
    ' Reset form fields
    With frmTermin
        .txtName.Value = ""
        .txtKG.Value = ""
        .txtEZ.Value = ""
        .txtTel.Value = ""
        .txtAnliegen.Value = ""
        .cmdLoeschen.Visible = False
        
        ' Pre-fill if already has content
        If Trim(Target.Value) <> "" Then
            Dim lines() As String
            lines = Split(Target.Value, Chr(10))
            Dim i As Integer
            For i = 0 To UBound(lines)
                If InStr(lines(i), ": ") > 0 Then
                    Dim kv() As String
                    kv = Split(lines(i), ": ", 2)
                    Select Case Trim(kv(0))
                        Case "Name":     .txtName.Value = Trim(kv(1))
                        Case "KG":       .txtKG.Value = Trim(kv(1))
                        Case "EZ":       .txtEZ.Value = Trim(kv(1))
                        Case "Tel":      .txtTel.Value = Trim(kv(1))
                        Case "Anliegen": .txtAnliegen.Value = Trim(kv(1))
                    End Select
                End If
            Next i
            .cmdLoeschen.Visible = True
        End If
        
        ' Update title with time + date
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets("Kalender")
        Dim timeStr As String: timeStr = ws.Cells(gSelRow, 1).Value
        Dim dateStr As String: dateStr = ws.Cells(1, gSelCol).Value
        .lblTitle.Caption = timeStr & "  |  " & dateStr
        
        .Show
    End With
End Sub

Sub SpeichereTermin()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Kalender")
    Dim cell As Range
    Set cell = ws.Range(gSelRef)
    
    Dim sName As String: sName = Trim(frmTermin.txtName.Value)
    Dim sKG As String:   sKG   = Trim(frmTermin.txtKG.Value)
    Dim sEZ As String:   sEZ   = Trim(frmTermin.txtEZ.Value)
    Dim sTel As String:  sTel  = Trim(frmTermin.txtTel.Value)
    Dim sAnl As String:  sAnl  = Trim(frmTermin.txtAnliegen.Value)
    
    If sName = "" And sKG = "" And sEZ = "" And sTel = "" And sAnl = "" Then
        MsgBox "Bitte mindestens ein Feld ausfüllen.", vbInformation, "Hinweis"
        Exit Sub
    End If
    
    ' Build cell content
    Dim content As String
    If sName <> "" Then content = content & "Name: " & sName & Chr(10)
    If sKG   <> "" Then content = content & "KG: "   & sKG   & Chr(10)
    If sEZ   <> "" Then content = content & "EZ: "   & sEZ   & Chr(10)
    If sTel  <> "" Then content = content & "Tel: "  & sTel  & Chr(10)
    If sAnl  <> "" Then content = content & sAnl
    
    cell.Value = Trim(content)
    
    With cell
        .Font.Size = 8
        .Font.Name = "Calibri"
        .Font.Color = RGB(30, 58, 95)
        .Interior.Color = RGB(214, 228, 247)
        .WrapText = True
    End With
    
    ' Save to data sheet
    SaveToDataSheet sName, sKG, sEZ, sTel, sAnl
    RefreshUebersicht
    RefreshTagesliste
    
    Unload frmTermin
End Sub

Sub LoescheTermin()
    If MsgBox("Termin wirklich löschen?", vbYesNo + vbQuestion, "Termin löschen") = vbNo Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Kalender")
    Dim cell As Range
    Set cell = ws.Range(gSelRef)
    
    cell.Value = ""
    cell.Interior.Color = RGB(255, 255, 255)
    
    DeleteFromDataSheet
    RefreshUebersicht
    RefreshTagesliste
    
    Unload frmTermin
End Sub

Private Sub SaveToDataSheet(sName As String, sKG As String, sEZ As String, _
                             sTel As String, sAnl As String)
    Dim wsd As Worksheet, wsc As Worksheet
    Set wsd = ThisWorkbook.Sheets("_Daten")
    Set wsc = ThisWorkbook.Sheets("Kalender")
    
    Dim lastRow As Long
    lastRow = wsd.Cells(wsd.Rows.Count, 1).End(xlUp).Row
    
    ' Find existing or new row
    Dim targetRow As Long: targetRow = lastRow + 1
    Dim i As Long
    For i = 2 To lastRow
        If wsd.Cells(i, 9).Value = gSelRow And wsd.Cells(i, 10).Value = gSelCol Then
            targetRow = i
            Exit For
        End If
    Next i
    
    wsd.Cells(targetRow, 1).Value = wsc.Cells(1, gSelCol).Value
    wsd.Cells(targetRow, 3).Value = wsc.Cells(gSelRow, 1).Value
    wsd.Cells(targetRow, 4).Value = sName
    wsd.Cells(targetRow, 5).Value = sKG
    wsd.Cells(targetRow, 6).Value = sEZ
    wsd.Cells(targetRow, 7).Value = sTel
    wsd.Cells(targetRow, 8).Value = sAnl
    wsd.Cells(targetRow, 9).Value = gSelRow
    wsd.Cells(targetRow, 10).Value = gSelCol
End Sub

Private Sub DeleteFromDataSheet()
    Dim wsd As Worksheet
    Set wsd = ThisWorkbook.Sheets("_Daten")
    Dim lastRow As Long
    lastRow = wsd.Cells(wsd.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = lastRow To 2 Step -1
        If wsd.Cells(i, 9).Value = gSelRow And wsd.Cells(i, 10).Value = gSelCol Then
            wsd.Rows(i).Delete
        End If
    Next i
End Sub

Sub RefreshUebersicht()
    Dim wsd As Worksheet, wso As Worksheet
    Set wsd = ThisWorkbook.Sheets("_Daten")
    Set wso = ThisWorkbook.Sheets("Terminübersicht")
    
    Dim lastOut As Long
    lastOut = wso.Cells(wso.Rows.Count, 1).End(xlUp).Row
    If lastOut > 2 Then wso.Rows("3:" & lastOut).Delete
    
    Dim lastData As Long
    lastData = wsd.Cells(wsd.Rows.Count, 1).End(xlUp).Row
    If lastData < 2 Then Exit Sub
    
    Dim outRow As Long: outRow = 3
    Dim i As Long
    For i = 2 To lastData
        If Trim(wsd.Cells(i, 4).Value) <> "" Then
            wso.Cells(outRow, 1).Value = wsd.Cells(i, 1).Value
            wso.Cells(outRow, 3).Value = wsd.Cells(i, 3).Value
            wso.Cells(outRow, 4).Value = wsd.Cells(i, 4).Value
            wso.Cells(outRow, 5).Value = wsd.Cells(i, 5).Value
            wso.Cells(outRow, 6).Value = wsd.Cells(i, 6).Value
            wso.Cells(outRow, 7).Value = wsd.Cells(i, 7).Value
            wso.Cells(outRow, 8).Value = wsd.Cells(i, 8).Value
            Dim bgC As Long
            If outRow Mod 2 = 0 Then bgC = RGB(238, 242, 248) Else bgC = RGB(250, 252, 255)
            With wso.Range(wso.Cells(outRow, 1), wso.Cells(outRow, 8))
                .Interior.Color = bgC
                .Font.Name = "Calibri"
                .Font.Size = 9
                .Font.Color = RGB(44, 62, 80)
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Color = RGB(200, 214, 232)
                .RowHeight = 18
            End With
            outRow = outRow + 1
        End If
    Next i
End Sub

Sub RefreshTagesliste()
    Dim wsd As Worksheet, wst As Worksheet, wsc As Worksheet
    Set wsd = ThisWorkbook.Sheets("_Daten")
    Set wst = ThisWorkbook.Sheets("Tagesliste")
    Set wsc = ThisWorkbook.Sheets("Kalender")
    
    Dim lastOut As Long
    lastOut = wst.Cells(wst.Rows.Count, 1).End(xlUp).Row
    If lastOut > 2 Then wst.Rows("3:" & lastOut).Delete
    
    ' Get selected column's date header
    Dim selCol As Long: selCol = gSelCol
    If selCol < 2 Then selCol = 2
    Dim dateHdr As String: dateHdr = wsc.Cells(1, selCol).Value
    
    ' Update sheet title
    wst.Range("A1:F1").Merge
    With wst.Cells(1, 1)
        .Value = "Tagesliste  —  " & dateHdr
        .Font.Name = "Calibri"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Font.Size = 13
        .Interior.Color = RGB(30, 58, 95)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Dim lastData As Long
    lastData = wsd.Cells(wsd.Rows.Count, 1).End(xlUp).Row
    If lastData < 2 Then Exit Sub
    
    Dim outRow As Long: outRow = 3
    Dim i As Long
    For i = 2 To lastData
        If wsd.Cells(i, 1).Value = dateHdr And Trim(wsd.Cells(i, 4).Value) <> "" Then
            wst.Cells(outRow, 1).Value = wsd.Cells(i, 3).Value
            wst.Cells(outRow, 2).Value = wsd.Cells(i, 4).Value
            wst.Cells(outRow, 3).Value = wsd.Cells(i, 5).Value
            wst.Cells(outRow, 4).Value = wsd.Cells(i, 6).Value
            wst.Cells(outRow, 5).Value = wsd.Cells(i, 7).Value
            wst.Cells(outRow, 6).Value = wsd.Cells(i, 8).Value
            Dim bgC As Long
            If outRow Mod 2 = 0 Then bgC = RGB(238, 242, 248) Else bgC = RGB(250, 252, 255)
            With wst.Range(wst.Cells(outRow, 1), wst.Cells(outRow, 6))
                .Interior.Color = bgC
                .Font.Name = "Calibri"
                .Font.Size = 9
                .Font.Color = RGB(44, 62, 80)
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Color = RGB(200, 214, 232)
                .RowHeight = 18
            End With
            outRow = outRow + 1
        End If
    Next i
End Sub
