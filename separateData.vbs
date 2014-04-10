Sub separateData()

    Dim mainWB As Workbook
    Set mainWB = ActiveWorkbook

    '62 Blue Tab color '
	'-4142 Clear Tab color '
	blueTabColor = 62
	clearTabColor = -4142

    For Each sht In mainWB.Sheets

    	If sht.Tab.ColorIndex = 62 Or sht.Tab.ColorIndex = -4142 Then 
	    	Debug.Print sht.Name
	    	Debug.Print sht.Tab.ColorIndex
	    	Debug.Print vbnewline
    	End If
        sht.AutoFilterMode = False
    Next sht

End Sub


Sub filterByEDCname()




End Sub