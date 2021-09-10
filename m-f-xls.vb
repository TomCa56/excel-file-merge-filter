Sub mergeFilterXls()
'module to merge excel sheets removing headers and filtering for specific values

	Dim sheet_name As String
	Dim value_remove As String
	Dim search_col As Integer
	Dim int_remove As Integer

	Dim last_row As Long
    Dim last_col As Long
	
	'example values
	sheet_name = "Name"
	value_remove = "Value"
	search_col = 0
	search_row = 0
	int_remove = 100
	

    'Loop through worksheets and copy them to your new worksheet
	'Function DeleteHeaders()
		For Each ws In Worksheets
			ws.Activate
			If ws.Index <> 1 Then
			   Rows(1).Select
			   Selection.Delete
			End If
		Next
	'End Function
    
	'Insert a new worksheet
	'Function NewSheet(name As String)
		Worksheets(1).Select
		Sheets.Add.Name = sheet_name
	'End Function
    
	
	'Function MergeToAndDelete(name As String)
		'Select and copy
		For Each ws In Worksheets
			ws.Activate
			'copy sheet
			ws.UsedRange.Select
			Selection.Copy
			'to merge sheet
			Sheets(sheet_name).Activate
			'Select the last filled cell
			ActiveSheet.Range("A1048576").Select 'maximum nr of rows in excel 
			Selection.End(xlUp).Select
			'paste
			ActiveSheet.Paste
		Next

		'Delete other sheets
		For Each ws In Worksheets
			ws.Activate
			If ws.Name <> sheet_name Then
				Worksheets(ws.Name).Delete
			End If
		Next
	'End Function
   
   
   
	'select main sheet and determine last row and col with values
	'Function SheetRange(index As Integer)
		Worksheets(1).Activate
		With ActiveSheet
			last_row = .Range("A1").SpecialCells(xlCellTypeLastCell).Row
			last_col = .Range("A1").SpecialCells(xlCellTypeLastCell).Column
		End With
	'End Function
	
	'find and delete rows columns for specific cell values
	'Function SearchAndDelete(row As Integer, col As Integer, val_remove As String)
		For i = 1 To last_col
			For j = 1 To last_row
				If Cells(j, i).Value = value_remove Then
					Rows(j).Select
					'Columns(i).Select
					Selection.Delete
				End If
			Next
		Next
	'End Function
   

   'Function SearchAndDeleteRowOnCol(col As Integer, row_val As Integer)
		'filter for a specific value in a specific col
		
		'For i = 1 To last_row
			'If Cells(i, search_col).Value = int_remove Then
				'Rows(i).Select
				'Selection.Delete
			'End If
		'Next
	'End Function
   

   'Function SearchAndDeleteColOnRow(row As Integer, col_val As Integer)
		'filter for a specific value in a specific row
		'For i = 1 To last_col
			'If Cells(i, search_row).Value = int_remove Then
				'Rows(i).Select
				'Selection.Delete
			'End If
		'Next
	'End Function
   
   
End Sub