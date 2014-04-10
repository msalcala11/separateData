Sub separateData()

    Dim mainWB As Workbook
    Set mainWB = ActiveWorkbook

    Dim edcNames As Variant
    edcNames = getEDCnames()

    'Define the tab color indices of the sheets we will need to use
    blueTabColor = 62
    clearTabColor = -4142

    For Each edc In edcNames
        If edc = edcNames(0) Then
            AddNew (edc)

            For Each sht In mainWB.Sheets

                If sht.Tab.ColorIndex = 62 Or sht.Tab.ColorIndex = -4142 Then

                    If sht.Name = "InteriorLightings" Then Call filterSheetByEDCname(mainWB, sht.Name, edcNames(0))

                End If
                'sht.AutoFilterMode = False
            Next sht
        End If

    Next edc

End Sub

Function getEDCnames()
    'Grabs the EDC names we will need to filter by from a row in the Index Sheet and adds the names to an array for easy retrieval

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

Sub filterSheetByEDCname(mainWB As Workbook, ByVal sheetName As String, ByVal edcName As String)

    'Lets select the entire usable range'
    Dim r As Range
    Set r = mainWB.Sheets(sheetName).UsedRange

    Debug.Print sheetName
    Debug.Print edcName
    Debug.Print r.Rows.Count

    'Dim filteredRange As Range


    r.AutoFilter _
         field:=13, _
         Criteria1:=edcName, _
         VisibleDropDown:=False
End Sub


Sub AddNew(ByVal bookName As String)
    ' For creating new Workbooks
    Set NewBook = Workbooks.Add
        With NewBook
            .Title = bookName
            .SaveAs Filename:=ThisWorkbook.Path & "\" & bookName & ".xls"
        End With
End Sub

