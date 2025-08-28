Sub Button_Click()
    If CheckFilePath Then
        Call DownloadDataFromIQY
        Call CountPreviousRecords
        Call ProcessVisitorsData
        Call CountCurrentRecords
        If CheckIfNewRecords Then
            Call BringNewRecords
            MsgBox "New records found"
        Else
            MsgBox "Nothing to update"
        End If
            Call PrintLastConsultedDate
    Else
        MsgBox "The query file path is invalid or the file doesn't exist.", vbExclamation
    End If
End Sub


Function CheckFilePath() As Boolean
    Dim iqyPath As String
    
    ' Read and clean the path
    iqyPath = Trim(ThisWorkbook.Sheets("Information").Range("F4").Value)
    
    ' Default to False
    CheckFilePath = False
    
    ' Validate: must not be empty
    If iqyPath = "" Then Exit Function
    
    ' Validate: must have .iqy extension
    If LCase(Right(iqyPath, 4)) <> ".iqy" Then Exit Function
    
    ' Validate: must be an existing file (not folder)
    If Dir(iqyPath, vbNormal) = "" Then Exit Function
    
    ' If all checks pass
    CheckFilePath = True
End Function


Sub DownloadDataFromIQY()
    Dim iqyPath As String
    
    ' Read the path from cell
    iqyPath = ThisWorkbook.Sheets("Information").Range("F4").Value
    
    ' Add the query using the dynamic path
    With ThisWorkbook.Sheets("VisitorsRawData").QueryTables.Add( _
        Connection:="FINDER;" & iqyPath, _
        Destination:=ThisWorkbook.Sheets("VisitorsRawData").Range("A1"))
        .Refresh BackgroundQuery:=False
    End With
End Sub


Sub CountPreviousRecords()
    Set wsSrc = ThisWorkbook.Sheets("ProcessedVisitors")
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row - 1
    ThisWorkbook.Sheets("Information").Range("C4").Value = lastRow
End Sub


Sub ProcessVisitorsData()
    Dim wsSrc As Worksheet, wsDst As Worksheet
    Dim lastRow As Long, r As Long
    Dim dataArr As Variant, outArr() As Variant
    Dim dict As Object, groupDict As Object
    Dim key As String
    Dim arrival As String, departure As String, reason As String
    Dim classification As String, fullName As String, company As String
    Dim diffDates As String, visitName As String, customField As String
    Dim grp As Variant
    Dim counter As Long
    
    Set wsSrc = ThisWorkbook.Sheets("VisitorsRawData")
    Set wsDst = ThisWorkbook.Sheets("ProcessedVisitors")
    
    ' Find last row in source
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    
    ' Read data into array
    dataArr = wsSrc.Range("A1:I" & lastRow).Value
    
    ' Main dictionary for grouping
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    ' Loop over rows
    For r = 2 To UBound(dataArr, 1)
        arrival = Trim(dataArr(r, 5))
        departure = Trim(dataArr(r, 6))
        reason = Trim(dataArr(r, 7))
        classification = Trim(dataArr(r, 3))
        fullName = Trim(dataArr(r, 2))
        company = Trim(dataArr(r, 4))
        
        ' Grouping key
        key = arrival & "|" & departure & "|" & reason
        
        ' Create new group if it doesn't exist
        If Not dict.Exists(key) Then
            Set groupDict = CreateObject("Scripting.Dictionary")
            groupDict.CompareMode = vbTextCompare
            
            groupDict("Arrival") = arrival
            groupDict("Departure") = departure
            groupDict("Reason") = reason
            groupDict("Classification") = classification
            
            Set groupDict("Names") = CreateObject("Scripting.Dictionary")
            groupDict("Names").CompareMode = vbTextCompare
            Set groupDict("Companies") = CreateObject("Scripting.Dictionary")
            groupDict("Companies").CompareMode = vbTextCompare
            
            Set dict(key) = groupDict
        End If
        
        ' Add names and companies
        If Not dict(key)("Names").Exists(fullName) Then dict(key)("Names")(fullName) = True
        If Not dict(key)("Companies").Exists(company) Then dict(key)("Companies")(company) = True
    Next r
    
    ' Prepare output array
    ReDim outArr(0 To dict.Count, 1 To 6)
    outArr(0, 1) = "Visitor Classification"
    outArr(0, 2) = "Company Name"
    outArr(0, 3) = "Expected Arrival Date"
    outArr(0, 4) = "Expected Departure Date"
    outArr(0, 5) = "Reason for Visit"
    outArr(0, 6) = "Custom Field"
    
    ' Fill output array
    counter = 1
    For Each grp In dict.Keys
        arrival = dict(grp)("Arrival")
        departure = dict(grp)("Departure")
        reason = dict(grp)("Reason")
        classification = dict(grp)("Classification")
        
        ' --- Diff Dates
        If arrival <> departure Then
            diffDates = Format(CDate(arrival), "m/d") & "-" & Format(CDate(departure), "d")
        Else
            diffDates = ""
        End If
        
        ' --- Visit Name
        If InStr(1, classification, "Internal", vbTextCompare) > 0 Then
            If dict(grp)("Names").Count = 1 Then
                visitName = dict(grp)("Names").Keys()(0)
            Else
                visitName = dict(grp)("Names").Keys()(0) & " (+" & (dict(grp)("Names").Count - 1) & ")"
            End If
        Else
            visitName = dict(grp)("Companies").Keys()(0)
        End If
        
        ' --- Custom Field
        customField = ""
        If diffDates <> "" Then customField = diffDates & " "
        customField = customField & visitName & " " & classification & " Visit"
        
        ' Write to array
        outArr(counter, 1) = classification
        outArr(counter, 2) = Join(dict(grp)("Companies").Keys, ", ")
        outArr(counter, 3) = arrival
        outArr(counter, 4) = departure
        outArr(counter, 5) = reason
        outArr(counter, 6) = customField
        
        counter = counter + 1
    Next grp
    
    ' --- Clear previous output and write new one
    wsDst.Cells.Clear
    wsDst.Range("A1").Resize(UBound(outArr, 1) + 1, UBound(outArr, 2)).Value = outArr
    
    ' MsgBox "Processing completed. " & (counter - 1) & " grouped events created."
    
End Sub


Sub CountCurrentRecords()
    Set wsSrc = ThisWorkbook.Sheets("ProcessedVisitors")
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row - 1
    ThisWorkbook.Sheets("Information").Range("C5").Value = lastRow
End Sub

Function CheckIfNewRecords() As Boolean
    Dim Previous, Current As Integer
    
    Previous = Trim(ThisWorkbook.Sheets("Information").Range("C4").Value)
    Current = Trim(ThisWorkbook.Sheets("Information").Range("C5").Value)
    
    ThisWorkbook.Sheets("Information").Range("C6").Value = Current - Previous
    
    If Current - Previous = 0 Then
        CheckIfNewRecords = False
    Else
        CheckIfNewRecords = True
    End If
End Function


Sub BringNewRecords()
    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim X As Integer
    Dim dateAsLong As Long
    
    Set wsSrc = ThisWorkbook.Sheets("ProcessedVisitors")
    X = ThisWorkbook.Sheets("Information").Range("C6").Value
        
    Set wsSrc = ThisWorkbook.Sheets("ProcessedVisitors")
    Set wsDst = ThisWorkbook.Sheets("Information")
    
    ' Clean Previous Data
    wsDst.Range("B10:K1000").ClearContents
    
    ' Last row with data
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    
    ' Copy last registers to Information Sheet
    For i = 0 To X - 1
        dateAsLong = DateDiff("d", wsSrc.Cells(lastRow - i, 3).Value, wsSrc.Cells(lastRow - i, 4).Value) + 1
        wsDst.Cells(10 + i, 2).Value = wsSrc.Cells(lastRow - i, 3).Value ' Expected Arrival Date
        wsDst.Cells(10 + i, 3).Value = wsSrc.Cells(lastRow - i, 6).Value ' Custom Field
        
        wsDst.Cells(10 + i, 5).Value = wsSrc.Cells(lastRow - i, 3).Value ' Expected Arrival Date
        wsDst.Cells(10 + i, 6).Value = wsSrc.Cells(lastRow - i, 4).Value ' Expected Departure Date
        wsDst.Cells(10 + i, 7).Value = dateAsLong ' Expected Count Days
        wsDst.Cells(10 + i, 8).Value = wsSrc.Cells(lastRow - i, 1).Value ' Visitor Classification
        wsDst.Cells(10 + i, 9).Value = wsSrc.Cells(lastRow - i, 2).Value ' Company Name
        wsDst.Cells(10 + i, 10).Value = wsSrc.Cells(lastRow - i, 5).Value ' Reason for Visit
    Next i
End Sub


Sub PrintLastConsultedDate()
    ThisWorkbook.Sheets("Information").Range("F5").Value = Now
End Sub