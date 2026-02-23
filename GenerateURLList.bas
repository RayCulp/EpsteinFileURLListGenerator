Option Explicit
'____________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________

Sub GenerateURLList 

    ' Declarations
    	
    	Dim oSheets As Object 
        Dim oUnresolvedSheet As Object
        Dim oExtensionsSheet As Object
        Dim oMergedSheet As Object
        Dim lUnresolvedSheetLen As Long
        Dim lExtensionsSheetLen As Long 
        Dim lTargetRow As Long 
        Dim sUnresolvedBaseUrl As String 
        Dim sExtension As String 
        Dim sMergedUrl As String 
        Dim oCell As Object
        Dim i As Long
        Dim j As Long 
        
    ' Get references to sheets

    	oSheets = ThisComponent.Sheets
		oUnresolvedSheet = oSheets.getByName ("unresolved")
        oExtensionsSheet = oSheets.getByName ("extensions")
        oMergedSheet = oSheets.getByName ("merged")
        
	' Get table lengths
	
		lUnresolvedSheetLen = GetTableLength (oUnresolvedSheet, False)
		lExtensionsSheetLen = GetTableLength (oExtensionsSheet, False) 
        
    ' Loop through the cells in column A of the "unresolved" sheet

        For i = 0 To lUnresolvedSheetLen - 1 
        
            oCell = oUnresolvedSheet.getCellByPosition(0, i) ' Column 1 (A) is at position 0
            ThisComponent.CurrentController.Select(oCell)
            sUnresolvedBaseUrl = oUnresolvedSheet.getCellByPosition(0, i).Formula
            
		' Loop through cells in column A of the "extensions" sheet
            
	        For j = 0 To lExtensionsSheetLen - 1 
	        
	            oCell = oExtensionsSheet.getCellByPosition(0, j) ' Column 1 (A) is at position 0
	            ThisComponent.CurrentController.Select(oCell)
	            sExtension = oExtensionsSheet.getCellByPosition(0, j).Formula
	            sMergedUrl = sUnresolvedBaseUrl & sExtension
	            oCell = oMergedSheet.getCellByPosition(0, lTargetRow) ' Column 1 (A) is at position 0
	            ThisComponent.CurrentController.Select(oCell)
	            oMergedSheet.getCellByPosition (0, lTargetRow).Formula = "<a href=""" & sMergedUrl & """>" & sMergedUrl & "</a>"
				lTargetRow = lTargetRow + 1
	            
	        Next j
            
        Next i
        
End Sub
'____________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________

Public Function GetTableLength (ByVal oSheet As Object, ByVal blAnyModification As Boolean) As Long
'************************************************************************************************************
' Procedure Name: GetTableLength
' Purpose: Gets the real length of a sheet
' Procedure Type: Function
' Procedure Access: Public
' Parameter oSheet (Object): The sheet object to process
' Parameter blAnyModification (Boolean): Sets whether to check only for cells with contents,
'     or also with modifications such as background color or outlines
' Return value (Long): The length of the sheet (NOTICE: THIS NOT THE ROW NUMBER! SHEET
'     INDEX IS ZERO-BASED!
' Usage example:
' Author: Ray Culp
' Date: 05.06.2024
' More information: https://ask.libreoffice.org/t/basic-calc-how-to-get-address-of-last-cell-used-with-content/46656
'************************************************************************************************************

    ' Declarations
    
        Dim objCursor As Object
        Dim objRange As Object
        Dim objUsedRange As Object
        Dim lngLastRow As Long
        Dim objRangeAddress As Object 
        
        
    ' Check whether the function should check for cells with any modification or only those that contain text
    
        If blAnyModification = True Then
    
        ' Find the last cell with any modification, including:
        ' Background color
        ' Border
        ' Create a cursor
    
            objCursor = oSheet.createCursor
        
        ' Move objCursor to the end of the used area
        
            objCursor.gotoEndOfUsedArea(True)
            
        ' Get the row of the cursor position
            
            lngLastRow = objCursor.RangeAddress.EndRow + 1

        ElseIf blAnyModification = False Then
        
        ' Find the last cell containing any VALUE, DATETIME, STRING or FORMULA
        
        ' 1023 = com.sun.star.sheet.CellFlags.VALUE + com.sun.star.sheet.CellFlags.DATETIME + 
        ' com.sun.star.sheet.CellFlags.STRING + com.sun.star.sheet.CellFlags.FORMULA
        
        ' Get the range with any of the above content
        
            objRange = oSheet.queryContentCells(1023)
            
        ' Loop through all RangeAddresses in objRange. The highest row number will give us the actual length.
            
            For Each objRangeAddress In objRange.RangeAddresses
            
                If objRangeAddress.EndRow > lngLastRow Then 
                
                    lngLastRow = objRangeAddress.EndRow
                    
                End If
                
            Next objRangeAddress
            
            lngLastRow = lngLastRow + 1

        End If
       
    ' Return the table width
    
        GetTableLength = lngLastRow
       
End Function

