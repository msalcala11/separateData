Sub separateData()

    Dim mainWB As Workbook
    Set mainWB = ActiveWorkbook

    ' 'Lets grab the EDC names
    ' Dim EDCnameRange As Range
    ' Set EDCnameRange = ActiveWorkbook.Sheets("Index").Range("E2:K2")

    ' Dim edcNames(6) As String
    ' arrInd = 0
    ' For Each cell In EDCnameRange
    '       edcNames(arrInd) = cell.Value
    '       Debug.print edcNames(arrInd)
    '       arrInd = arrInd + 1
    ' Next cell

    Dim edcNames As Variant
    'Dim arr(6) As String
    'edcNames = arr
    edcNames = getEDCnames()

    '62 Blue Tab color '
    '-4142 Clear Tab color '
    blueTabColor = 62
    clearTabColor = -4142

    For Each sht In mainWB.Sheets

        If sht.Tab.ColorIndex = 62 Or sht.Tab.ColorIndex = -4142 Then
            ' Debug.Print sht.Name
            ' Debug.Print sht.Tab.ColorIndex
            ' Debug.Print vbnewline

            If sht.Name = "InteriorLightings" Then Call filterSheetByEDCname(sht.Name, edcNames(0))

        End If
        sht.AutoFilterMode = False
    Next sht

End Sub

Function getEDCnames()

    'Lets grab the EDC names from the Index Sheet
    Dim EDCnameRange As Range
    Set EDCnameRange = ActiveWorkbook.Sheets("Index").Range("E2:K2")

    'Now lets convert into an array
    Dim edcNames As Variant
    Dim arr(6) As String
    edcNames = arr
    arrInd = 0
    'Loop through each value in the range and add to array
    For Each cell In EDCnameRange
        edcNames(arrInd) = cell.Value
        Debug.Print edcNames(arrInd)
        arrInd = arrInd + 1
    Next cell
    
    'Return the Array
    getEDCnames = edcNames

End Function

Sub filterSheetByEDCname(ByVal sheetName As String, ByVal edcName As String)

    'Lets select the entire usable range'
    Dim r As Range
    Set r = ActiveWorkbook.Sheets(sheetName).UsedRange

    Debug.Print sheetName
    Debug.Print edcName
    Debug.Print r.Rows.Count

End Sub


