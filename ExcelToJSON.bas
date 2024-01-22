Attribute VB_Name = "ExcelToJSON"

' Excel Export Json For Azure Active Directory
' by Enrico Rossini - PureSourceCode.com
Option Explicit

'Stores the names of tables that the user wants to convert to JSON.
'The names are added to the array during the function SubmitBtn_Click() in ExcelToJSONForm.
Public usrSlctdTblsNameArray() As String

'Variables for iterating through loops
Public i As Integer, j As Integer, k As Integer

Public outputFileFQPN As String

Public numIndentationSpaces As Integer
Public currentIndentation As Integer

Public Function IsInArrayIndex(stringToFind As String, arr As Variant) As Long

    IsInArrayIndex = -1

    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i) <> "" And arr(i) = stringToFind Then
            IsInArrayIndex = i
            Exit Function
        End If
    Next i

End Function

Sub ExcelToJSON()
    
    numIndentationSpaces = 4
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    
    Dim originallySelectedSheet As String: originallySelectedSheet = ActiveSheet.Name
    
    Dim FALSEInLocalLang As String: FALSEInLocalLang = getLocalTranslationOfFALSE()
    
    Dim table As ListObject, sheet As Worksheet
    
    Dim numFormCtrlsNotCountingCheckBoxes As Integer: numFormCtrlsNotCountingCheckBoxes = ExcelToJSONForm.Controls.Count
    
    Dim tableNamesInCurrentWorkbook()
    
    Dim tblNameToChBxName As String
    
    Dim selectedSheetsAndTables As Object
    Set selectedSheetsAndTables = CreateObject("Scripting.Dictionary")
    
    'Loop through all of the tables in the workbook and generate checkboxes with corresponding names.
    'The checkboxes lets the user select which tables to export to JSON in ExcelToJSONForm.
    For Each sheet In Worksheets
        
        selectedSheetsAndTables.Add sheet.Name, "Test"
        
        For Each table In sheet.ListObjects
            i = i + 1
            ReDim Preserve tableNamesInCurrentWorkbook(0 To i + 1)
            tableNamesInCurrentWorkbook(i) = table.Name
            
            tblNameToChBxName = table.Name & "ChBx"
            Dim addNewCheckbox As Boolean: addNewCheckbox = ExcelToJSONForm.Controls.Add("forms.checkbox.1", tblNameToChBxName, True)
            
        Next table
    Next sheet
    
    i = 0
    Dim userFormControl As Object
    For Each userFormControl In ExcelToJSONForm.Controls
        i = i + 1
        
        If i > numFormCtrlsNotCountingCheckBoxes Then
            With userFormControl
                .Top = (i * 30) - 80
                .Left = 108
                .Caption = tableNamesInCurrentWorkbook(i - numFormCtrlsNotCountingCheckBoxes)
                .AutoSize = True
            End With
        End If
        
    Next userFormControl
    
    Dim UsrFormWindowHeight As Integer: UsrFormWindowHeight = ExcelToJSONForm.Controls.Count * 30
    
    ExcelToJSONForm.SubmitBtn.Top = UsrFormWindowHeight - 60
    ExcelToJSONForm.CancelBtn.Top = UsrFormWindowHeight - 60
    
    With ExcelToJSONForm
        .Height = UsrFormWindowHeight
        .Width = 400
        .Show
    End With
    
    Do While ExcelToJSONForm.Visible = True
    Loop
    
    outputFileFQPN = Application.GetSaveAsFilename(FileFilter:="JSON Files (*.json), *.json")
    If outputFileFQPN <> vbNullString And outputFileFQPN <> FALSEInLocalLang Then
        
        Open outputFileFQPN For Output As #1
        
        'Print the opening bracket for the JSON object as well as a key/value for a string of the current file name and
        'the JSON key and opening bracket for the object representing all worksheets in the workbook
        Print #1, "{"
        Print #1, createJsonKeyValuePair(numIndentationSpaces * 1, "Groups", "[", False)
        
        Worksheets(1).Activate
        
        Dim numTablesInSheet As Integer
        
        Dim sheetContainsSelectedTables As Boolean
        
        Dim printCommaUnlessLastIteration As Boolean
        
        Dim printCommaUnlessLastTable As Boolean
        
        Dim numSheetsToLoopThrough As Integer
        
        Dim currentSheetHasASelectedTable As Boolean: currentSheetHasASelectedTable = False
        
        For Each sheet In Worksheets
            
            For Each table In sheet.ListObjects
                For i = 0 To UBound(usrSlctdTblsNameArray)
                    If table.Name = usrSlctdTblsNameArray(i) Then
                        currentSheetHasASelectedTable = True
                    End If
                Next i
            Next table
            
            If currentSheetHasASelectedTable Then
                numSheetsToLoopThrough = numSheetsToLoopThrough + 1
                currentSheetHasASelectedTable = False
            End If
            
        Next sheet
        
        Dim numSelectedTablesInSheet As Integer
        
        For Each sheet In Worksheets
        
            numSelectedTablesInSheet = countNumSelectedTablesInSheet(sheet.Name, usrSlctdTblsNameArray)
        
            sheetContainsSelectedTables = False
            
            numTablesInSheet = 0
            
            Dim openingBracketsPrintedForCurrentSheet As Boolean: openingBracketsPrintedForCurrentSheet = False
            
            For Each table In sheet.ListObjects
                If IsInArrayIndex(table.Name, usrSlctdTblsNameArray) > 0 Then
                    numTablesInSheet = numTablesInSheet + 1
                    
                    printCommaUnlessLastTable = False
                    
                    For i = 0 To numSelectedTablesInSheet
                            
                            sheetContainsSelectedTables = True
                            numSelectedTablesInSheet = numSelectedTablesInSheet - 1
                            printCommaUnlessLastTable = numSelectedTablesInSheet > 0
                            openingBracketsPrintedForCurrentSheet = True
                        
                            Dim groupName As String
                            Dim appRegName As String
                            Dim appRoles As String
                            Dim json As String
                            
                            For j = 1 To table.ListRows.Count
                                table.ListRows(j).Range.Select
                                
                                Dim tableRowIndexCellValue As String: tableRowIndexCellValue = ActiveCell.Value
                                Dim tableNamePlusIterationNumber As String: tableNamePlusIterationNumber = WorksheetFunction.Concat(table.Name, j)
                                Dim tableRowKey As String: tableRowKey = IIf((tableRowIndexCellValue = vbNullString), tableNamePlusIterationNumber, tableRowIndexCellValue)
                            
                                If groupName <> tableRowIndexCellValue Then
                                    If Len(appRoles) > 0 Then
                                        Print #1, createJsonKeyValuePair(numIndentationSpaces * 4, "AppRoles", "[", False)
                                        Print #1, createJsonArray(numIndentationSpaces * 5, appRoles)
                                        Print #1, createJsonArrayBracket(numIndentationSpaces * 4, "]", False)
                                        Print #1, createJsonArrayBracket(numIndentationSpaces * 3, "}", False)
                                        
                                        appRegName = ""
                                        appRoles = ""
                                    End If
                                    If Len(groupName) > 0 Then
                                        Print #1, createJsonArrayBracket(numIndentationSpaces * 2, "]", False)
                                        Print #1, createJsonArrayBracket(numIndentationSpaces * 1, "}", True)
                                    End If
                                    
                                    groupName = tableRowIndexCellValue
                                    Print #1, createJsonArrayBracket(numIndentationSpaces * 1, "{", False)
                                    Print #1, createJsonKeyValuePair(numIndentationSpaces * 2, "GroupName", groupName, True)
                                    Print #1, createJsonKeyValuePair(numIndentationSpaces * 2, "AppRegs", "[", False)
                                    
                                    appRegName = ""
                                    appRoles = ""
                                End If
                                
                                'Loop through all cells in the current row, starting with the 2nd cell from the left.
                                'The loop starts with the 2nd cell from the left because the 1st cell from the left is
                                'converted to a JSON Key for the rest of the cells in the same table row and the cells in the same row are converted to JSON Values.
                                For k = 2 To table.ListColumns.Count
                                    ActiveCell.Offset(, 1).Activate
                                    
                                    Dim vl As String
                                    vl = CStr(ActiveCell.Value)
                                    
                                    If k = 2 Then
                                        If vl <> appRegName Then
                                            If Len(appRoles) > 0 Then
                                                Print #1, createJsonKeyValuePair(numIndentationSpaces * 4, "AppRoles", "[", False)
                                                Print #1, createJsonArray(numIndentationSpaces * 5, appRoles)
                                                Print #1, createJsonArrayBracket(numIndentationSpaces * 4, "]", False)
                                                Print #1, createJsonArrayBracket(numIndentationSpaces * 3, "}", True)
                                            End If
                                            Print #1, createJsonArrayBracket(numIndentationSpaces * 3, "{", False)
                                            Print #1, createJsonKeyValuePair(numIndentationSpaces * 4, "AppRegName", vl, True)
                                            
                                            appRegName = vl
                                            appRoles = ""
                                        End If
                                    End If
                                    If k = 3 Then
                                        If Len(appRoles) > 0 Then appRoles = appRoles + ", "
                                        appRoles = appRoles & " """ & vl & """"
                                    End If
                                Next k
                                
                                'Reselect the index cell of the current row
                                ActiveCell.Offset(, (table.ListColumns.Count * -1) + 1).Activate
                                
                                printCommaUnlessLastIteration = j < table.ListRows.Count
                                
                                If table.ListRows.Count = j Then
                                    Print #1, createJsonKeyValuePair(numIndentationSpaces * 4, "AppRoles", "[", False)
                                    Print #1, createJsonArray(numIndentationSpaces * 5, appRoles)
                                    Print #1, createJsonArrayBracket(numIndentationSpaces * 4, "]", False)
                                    Print #1, createJsonArrayBracket(numIndentationSpaces * 3, "}", False)
                                    Print #1, createJsonArrayBracket(numIndentationSpaces * 2, "]", False)
                                    Print #1, createJsonArrayBracket(numIndentationSpaces * 1, "}", False)
                                    Print #1, createJsonArrayBracket(numIndentationSpaces * 0, "]", False)
                                    Print #1, createJsonArrayBracket(numIndentationSpaces * 0, "}", False)
                                End If
                            Next j
                    Next
                End If
            Next table
            
            'If the loop is not on its last iteration, select the next sheet in the workbook
            If ActiveSheet.Index < Worksheets.Count Then
                Worksheets(ActiveSheet.Index + 1).Activate
            End If
            
        Next sheet
        
        Close #1
        End
    Else
        End
    End If
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Public Function createJsonArray(numSpacesToIndent As Integer, arrayString As String) As String
    Dim output As String: output = vbNullString
    
    Dim spaces As String
    spaces = ""
    For i = 1 To numSpacesToIndent
        spaces = spaces & " "
    Next i
    
    Dim arrSplitStrings() As String
    arrSplitStrings = Split(arrayString, ",")

    For i = LBound(arrSplitStrings, 1) To UBound(arrSplitStrings, 1)
        output = output & spaces & LTrim(RTrim(arrSplitStrings(i)))
        If i < UBound(arrSplitStrings, 1) Then output = output & "," & Chr(13)
    Next i
        
    createJsonArray = output
End Function

Public Function createJsonKeyValuePair(numSpacesToIndent As Integer, keyString As String, valueString As String, showComma As Boolean) As String
    Dim output As String: output = vbNullString
    
    For i = 1 To numSpacesToIndent
        output = output & " "
    Next i
    
    output = output & Chr$(34) & Replace(keyString, Chr$(34), "\" & Chr$(34)) & Chr$(34)
    output = output & ": "
    
    valueString = Replace(valueString, Chr$(34), "\" & Chr$(34))
    output = IIf(valueString = "{", output & "{", IIf(valueString = "[", output & "[", output & Chr$(34) & valueString & Chr$(34)))
    
    output = IIf(showComma, output & ",", output)
    
    createJsonKeyValuePair = output
End Function

Public Function createJsonClosingBracket(numSpacesToIndent As Integer, showComma As Boolean) As String
    Dim output As String: output = vbNullString
    
    For i = 1 To numSpacesToIndent
        output = output & " "
    Next i
    
    output = output & "}"
    output = IIf(showComma, output & ",", output)
    
    createJsonClosingBracket = output
End Function

Public Function createJsonArrayBracket(numSpacesToIndent As Integer, valueString As String, showComma As Boolean) As String
    Dim output As String: output = vbNullString
    
    For i = 1 To numSpacesToIndent
        output = output & " "
    Next i
    
    output = output & valueString
    output = IIf(showComma, output & ",", output)
    
    createJsonArrayBracket = output
End Function

Public Function getLocalTranslationOfFALSE()
    'The last possible cell "XFD1048576" is selected because of its low risk of containing sensitive iformation and thus
    'can be used to paste the local translation of "false"
    Dim originalValue As Variant: originalValue = Range("XFD1048576").Value
    
    Range("XFD1048576").Value = False
    Dim localTranslation As String
    localTranslation = Range("XFD1048576").Value
    
    Range("XFD1048576").Value = originalValue
    
    getLocalTranslationOfFALSE = localTranslation
End Function

Public Function isInArray(stringToSearchFor As String, arrayName() As String)
    isInArray = UBound(Filter(arrayName, stringToSearchFor)) > -1
End Function

Public Function countNumSelectedTablesInSheet(sheetName As String, tableNamesArray() As String)
    Dim table As ListObject
    Dim numTables As Integer: numTables = 0
    For Each table In Sheets(sheetName).ListObjects
        For i = 0 To UBound(tableNamesArray)
            If table.Name = tableNamesArray(i) Then
                numTables = numTables + 1
            End If
        Next i
    Next table
    countNumSelectedTablesInSheet = numTables
End Function
