Option Explicit

Sub UpdateHistoryLog()
    Dim rawSheet As Worksheet
    Dim histSheet As Worksheet
    Dim statusCol As Long
    Dim lastRawRow As Long
    Dim lastHistRow As Long
    Dim histMap As Object
    Dim rawMap As Object
    Dim i As Long
    Dim histRow As Long
    Dim paymentID As String
    Dim status As String

    Set rawSheet = ThisWorkbook.Worksheets("Raw report")
    Set histSheet = ThisWorkbook.Worksheets("History Log")

    ' Find next available dynamic status column starting from column N
    statusCol = 14 ' Column N
    Do While histSheet.Cells(1, statusCol).Value <> ""
        statusCol = statusCol + 1
    Loop
    histSheet.Cells(1, statusCol).Value = rawSheet.Range("A5").Value
    histSheet.Cells(2, statusCol).Value = Now

    ' Build dictionary of existing Payment IDs in History Log
    Set histMap = CreateObject("Scripting.Dictionary")
    lastHistRow = histSheet.Cells(histSheet.Rows.Count, "M").End(xlUp).Row
    For i = 3 To lastHistRow
        paymentID = CStr(histSheet.Cells(i, "M").Value)
        If paymentID <> "" And Not histMap.Exists(paymentID) Then
            histMap.Add paymentID, i
        End If
    Next i

    ' Dictionary to track Payment IDs currently in Raw report
    Set rawMap = CreateObject("Scripting.Dictionary")
    lastRawRow = rawSheet.Cells(rawSheet.Rows.Count, "M").End(xlUp).Row
    For i = 8 To lastRawRow
        paymentID = CStr(rawSheet.Cells(i, "M").Value)
        If paymentID <> "" Then
            status = CStr(rawSheet.Cells(i, "B").Value)
            rawMap(paymentID) = True
            If histMap.Exists(paymentID) Then
                ' Existing Payment ID - update status columns only
                histRow = histMap(paymentID)
                histSheet.Cells(histRow, "B").Value = status
                histSheet.Cells(histRow, statusCol).Value = status
            Else
                ' New Payment ID - append to bottom of History Log
                lastHistRow = lastHistRow + 1
                histSheet.Cells(lastHistRow, "A").Value = rawSheet.Cells(i, "A").Value
                histSheet.Cells(lastHistRow, "B").Value = status
                histSheet.Cells(lastHistRow, "C").Value = rawSheet.Cells(i, "C").Value
                histSheet.Cells(lastHistRow, "D").Value = rawSheet.Cells(i, "D").Value
                histSheet.Cells(lastHistRow, "E").Value = rawSheet.Cells(i, "E").Value
                histSheet.Cells(lastHistRow, "F").Value = rawSheet.Cells(i, "H").Value
                histSheet.Cells(lastHistRow, "G").Value = rawSheet.Cells(i, "I").Value
                histSheet.Cells(lastHistRow, "H").Value = rawSheet.Cells(i, "K").Value
                histSheet.Cells(lastHistRow, "I").Value = rawSheet.Cells(i, "M").Value
                histSheet.Cells(lastHistRow, "J").Value = rawSheet.Cells(i, "N").Value
                histSheet.Cells(lastHistRow, "M").Value = rawSheet.Cells(i, "M").Value
                histSheet.Cells(lastHistRow, statusCol).Value = status
            End If
        End If
    Next i

    ' Mark Payment IDs not present in Raw report as Cleared
    For Each paymentID In histMap.Keys
        If Not rawMap.Exists(paymentID) Then
            histRow = histMap(paymentID)
            histSheet.Cells(histRow, "B").Value = "Cleared"
            histSheet.Cells(histRow, statusCol).Value = "Cleared"
        End If
    Next paymentID
End Sub
