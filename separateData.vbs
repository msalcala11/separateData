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

                If sht.Tab.ColorIndex = 62 Or sht.Tab.ColorIndex = 62 Then

                    'If sht.Name = "InteriorLightings" Then Call filterSheetByEDCname(mainWB, sht.Name, edcNames(0))
                    Call filterSheetByEDCname(mainWB, sht.Name, edcNames(0))

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

    Dim edcWorkbook As Workbook
    Set edcWorkbook = Workbooks(edcName & ".xls")

    Debug.Print edcWorkbook.Sheets.Count

    ' Debug.Print sheetName
    ' Debug.Print edcName
    ' Debug.Print r.Rows.Count

    'Lets find the header cell that contains "EDC Name"
    'mainWB.Sheets(sheetName).Range("A1").Select

    Dim edcNameCell As Range
    Set edcNameCell = mainWB.Sheets(sheetName).Cells.find(What:="EDC Name", After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
    xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True _
    , SearchFormat:=False)

    Debug.Print "EDC Name Column: " & edcNameCell.Column

    'Dim filteredRange As Range
    r.AutoFilter _
         field:=edcNameCell.Column, _
         Criteria1:=edcName, _
         VisibleDropDown:=False

    r.SpecialCells(xlCellTypeVisible).Copy

    If edcWorkbook.Sheets(1).Name = "Sheet1" Then
        ' If the first Sheet is named Sheet1 then we are pasting into the EDC workbook for the first time and do not have
        ' to create a new sheet. Instead we paste into Sheet1 and rename it to the sheet name we are copying from
        'edcWorkbook.Sheets("Sheet1").Paste
        edcWorkbook.Sheets("Sheet1").Name = sheetName
    Else 
        ' If the first sheet is not Sheet1 then we need to create a new sheet to hold our copied data 
        edcWorkbook.Sheets.Add(After:=edcWorkbook.Sheets(edcWorkbook.Sheets.Count)).Name = sheetName
    End If

    edcWorkbook.Sheets(sheetName).Paste
    edcWorkbook.Sheets(sheetName).Columns.AutoFit
    edcWorkbook.Sheets(sheetName).Range("A1").Select

End Sub


Sub AddNew(ByVal bookName As String)
    ' For creating new Workbooks
    Set NewBook = Workbooks.Add
        With NewBook
            .Title = bookName
            .SaveAs Filename:=ThisWorkbook.Path & "\" & bookName & ".xls"
        End With
    Application.DisplayAlerts = False
    NewBook.Sheets("Sheet2").Delete
    NewBook.Sheets("Sheet3").Delete
    Application.DisplayAlerts = True
End Sub


Sub unfilterSheets()

For Each sht In ActiveWorkbook.Sheets
    sht.AutoFilterMode = False
Next sht

End Sub