Option Explicit

'================================================================================
' EXCEL VBA - COLUMN MANAGER
' Insert and Remove Columns Based on List Values
'================================================================================
' Author: Created for Azure DevOps Project Management
' Date: November 5, 2025
' Purpose: Dynamically insert and remove columns in Excel based on a list
'================================================================================

'================================================================================
' MAIN PROCEDURES
'================================================================================

'--------------------------------------------------------------------------------
' InsertColumnsFromList
' Inserts new columns after a specified column based on list of column names
'
' Parameters:
'   targetSheet: Worksheet where columns will be inserted
'   listSheet: Worksheet containing the list of column names
'   listRange: Range containing column names (e.g., "A1:A10")
'   afterColumn: Column letter or number after which to insert (e.g., "C" or 3)
'   copyFormat: Optional - Copy format from adjacent column (default: True)
'
' Example Usage:
'   Call InsertColumnsFromList(Sheet1, Sheet2, "A2:A10", "C", True)
'--------------------------------------------------------------------------------
Sub InsertColumnsFromList(targetSheet As Worksheet, _
                         listSheet As Worksheet, _
                         listRange As String, _
                         afterColumn As Variant, _
                         Optional copyFormat As Boolean = True)
    
    On Error GoTo ErrorHandler
    
    Dim columnList As Range
    Dim cell As Range
    Dim columnName As String
    Dim insertPosition As Long
    Dim newColumn As Range
    Dim formatColumn As Range
    Dim counter As Long
    
    ' Convert column reference to number if it's a letter
    If IsNumeric(afterColumn) Then
        insertPosition = CLng(afterColumn)
    Else
        insertPosition = Range(afterColumn & "1").Column
    End If
    
    ' Get the list of column names
    Set columnList = listSheet.Range(listRange)
    
    ' Remove any blank cells from the list
    Dim validColumns As Collection
    Set validColumns = New Collection
    
    For Each cell In columnList
        If Trim(cell.Value) <> "" Then
            validColumns.Add Trim(cell.Value)
        End If
    Next cell
    
    ' Disable screen updating for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Insert columns from the list
    counter = 0
    Dim colName As Variant
    For Each colName In validColumns
        columnName = CStr(colName)
        
        ' Check if column already exists
        If Not ColumnExists(targetSheet, columnName) Then
            ' Insert new column
            targetSheet.Columns(insertPosition + counter + 1).Insert Shift:=xlToRight
            
            ' Set column header
            targetSheet.Cells(1, insertPosition + counter + 1).Value = columnName
            
            ' Copy format from adjacent column if requested
            If copyFormat Then
                Set formatColumn = targetSheet.Columns(insertPosition + counter)
                Set newColumn = targetSheet.Columns(insertPosition + counter + 1)
                
                formatColumn.Copy
                newColumn.PasteSpecial Paste:=xlPasteFormats
                Application.CutCopyMode = False
            End If
            
            counter = counter + 1
        Else
            Debug.Print "Column '" & columnName & "' already exists. Skipping."
        End If
    Next colName
    
    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox counter & " column(s) inserted successfully!", vbInformation, "Insert Columns"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error inserting columns: " & Err.Description, vbCritical, "Error"
End Sub

'--------------------------------------------------------------------------------
' RemoveColumnsFromList
' Removes columns from worksheet based on list of column names
'
' Parameters:
'   targetSheet: Worksheet where columns will be removed
'   listSheet: Worksheet containing the list of column names to remove
'   listRange: Range containing column names (e.g., "A1:A10")
'   confirmDelete: Optional - Prompt for confirmation before deleting (default: True)
'
' Example Usage:
'   Call RemoveColumnsFromList(Sheet1, Sheet2, "A2:A10", True)
'--------------------------------------------------------------------------------
Sub RemoveColumnsFromList(targetSheet As Worksheet, _
                         listSheet As Worksheet, _
                         listRange As String, _
                         Optional confirmDelete As Boolean = True)
    
    On Error GoTo ErrorHandler
    
    Dim columnList As Range
    Dim cell As Range
    Dim columnName As String
    Dim columnNumber As Long
    Dim counter As Long
    Dim deletedColumns As String
    Dim response As VbMsgBoxResult
    
    ' Get the list of column names
    Set columnList = listSheet.Range(listRange)
    
    ' Count valid columns to delete
    Dim validColumns As Collection
    Set validColumns = New Collection
    
    For Each cell In columnList
        If Trim(cell.Value) <> "" Then
            validColumns.Add Trim(cell.Value)
        End If
    Next cell
    
    ' Confirm deletion if requested
    If confirmDelete Then
        response = MsgBox("Are you sure you want to delete " & validColumns.Count & _
                         " column(s)?" & vbCrLf & vbCrLf & _
                         "This action cannot be undone.", _
                         vbYesNo + vbQuestion, "Confirm Delete")
        If response = vbNo Then Exit Sub
    End If
    
    ' Disable screen updating for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Remove columns from the list (delete in reverse order to maintain positions)
    counter = 0
    deletedColumns = ""
    
    Dim colName As Variant
    For Each colName In validColumns
        columnName = CStr(colName)
        columnNumber = GetColumnNumber(targetSheet, columnName)
        
        If columnNumber > 0 Then
            targetSheet.Columns(columnNumber).Delete
            counter = counter + 1
            deletedColumns = deletedColumns & columnName & vbCrLf
        Else
            Debug.Print "Column '" & columnName & "' not found. Skipping."
        End If
    Next colName
    
    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    If counter > 0 Then
        MsgBox counter & " column(s) deleted successfully!" & vbCrLf & vbCrLf & _
               "Deleted columns:" & vbCrLf & deletedColumns, _
               vbInformation, "Delete Columns"
    Else
        MsgBox "No columns were deleted.", vbInformation, "Delete Columns"
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error removing columns: " & Err.Description, vbCritical, "Error"
End Sub

'--------------------------------------------------------------------------------
' InsertColumnsAtPosition
' Inserts columns at specific positions with names from a list
'
' Parameters:
'   targetSheet: Worksheet where columns will be inserted
'   columnPositions: Dictionary of column names and positions (e.g., "Name":3, "Age":5)
'   copyFormat: Optional - Copy format from previous column (default: True)
'
' Example Usage:
'   Dim positions As Object
'   Set positions = CreateObject("Scripting.Dictionary")
'   positions.Add "New Column 1", 3
'   positions.Add "New Column 2", 5
'   Call InsertColumnsAtPosition(Sheet1, positions, True)
'--------------------------------------------------------------------------------
Sub InsertColumnsAtPosition(targetSheet As Worksheet, _
                           columnPositions As Object, _
                           Optional copyFormat As Boolean = True)
    
    On Error GoTo ErrorHandler
    
    Dim key As Variant
    Dim columnName As String
    Dim position As Long
    Dim counter As Long
    
    ' Disable screen updating
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Sort positions in descending order to insert from right to left
    Dim sortedKeys() As Variant
    Dim sortedPositions() As Long
    Dim i As Long, j As Long
    Dim tempKey As Variant
    Dim tempPos As Long
    
    ReDim sortedKeys(0 To columnPositions.Count - 1)
    ReDim sortedPositions(0 To columnPositions.Count - 1)
    
    i = 0
    For Each key In columnPositions.Keys
        sortedKeys(i) = key
        sortedPositions(i) = columnPositions(key)
        i = i + 1
    Next key
    
    ' Bubble sort (descending)
    For i = 0 To UBound(sortedPositions) - 1
        For j = i + 1 To UBound(sortedPositions)
            If sortedPositions(i) < sortedPositions(j) Then
                tempPos = sortedPositions(i)
                sortedPositions(i) = sortedPositions(j)
                sortedPositions(j) = tempPos
                
                tempKey = sortedKeys(i)
                sortedKeys(i) = sortedKeys(j)
                sortedKeys(j) = tempKey
            End If
        Next j
    Next i
    
    ' Insert columns at specified positions
    For i = 0 To UBound(sortedKeys)
        columnName = CStr(sortedKeys(i))
        position = sortedPositions(i)
        
        If Not ColumnExists(targetSheet, columnName) Then
            targetSheet.Columns(position).Insert Shift:=xlToRight
            targetSheet.Cells(1, position).Value = columnName
            
            If copyFormat And position > 1 Then
                targetSheet.Columns(position - 1).Copy
                targetSheet.Columns(position).PasteSpecial Paste:=xlPasteFormats
                Application.CutCopyMode = False
            End If
            
            counter = counter + 1
        End If
    Next i
    
    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox counter & " column(s) inserted at specified positions!", vbInformation, "Insert Columns"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error inserting columns: " & Err.Description, vbCritical, "Error"
End Sub

'================================================================================
' HELPER FUNCTIONS
'================================================================================

'--------------------------------------------------------------------------------
' ColumnExists
' Check if a column with specified header name exists in worksheet
'
' Returns: True if column exists, False otherwise
'--------------------------------------------------------------------------------
Function ColumnExists(ws As Worksheet, columnName As String) As Boolean
    On Error Resume Next
    
    Dim headerRow As Range
    Dim cell As Range
    
    ' Assume headers are in row 1
    Set headerRow = ws.Rows(1)
    
    For Each cell In headerRow.Cells
        If Trim(UCase(cell.Value)) = Trim(UCase(columnName)) Then
            ColumnExists = True
            Exit Function
        End If
        ' Stop at first empty cell to improve performance
        If cell.Column > 1 And Trim(cell.Value) = "" And Trim(cell.Offset(0, -1).Value) = "" Then
            Exit For
        End If
    Next cell
    
    ColumnExists = False
End Function

'--------------------------------------------------------------------------------
' GetColumnNumber
' Get the column number for a column with specified header name
'
' Returns: Column number if found, 0 if not found
'--------------------------------------------------------------------------------
Function GetColumnNumber(ws As Worksheet, columnName As String) As Long
    On Error Resume Next
    
    Dim headerRow As Range
    Dim cell As Range
    
    ' Assume headers are in row 1
    Set headerRow = ws.Rows(1)
    
    For Each cell In headerRow.Cells
        If Trim(UCase(cell.Value)) = Trim(UCase(columnName)) Then
            GetColumnNumber = cell.Column
            Exit Function
        End If
        ' Stop at first empty cell
        If cell.Column > 1 And Trim(cell.Value) = "" And Trim(cell.Offset(0, -1).Value) = "" Then
            Exit For
        End If
    Next cell
    
    GetColumnNumber = 0
End Function

'--------------------------------------------------------------------------------
' GetColumnLetter
' Convert column number to column letter
'
' Returns: Column letter (e.g., "A", "AA", "XFD")
'--------------------------------------------------------------------------------
Function GetColumnLetter(columnNumber As Long) As String
    GetColumnLetter = Split(Cells(1, columnNumber).Address, "$")(1)
End Function

'================================================================================
' ADVANCED PROCEDURES
'================================================================================

'--------------------------------------------------------------------------------
' SyncColumnsWithList
' Synchronize worksheet columns with a master list
' - Adds missing columns
' - Removes extra columns not in list
' - Reorders columns to match list order
'
' Parameters:
'   targetSheet: Worksheet to synchronize
'   listSheet: Worksheet containing master column list
'   listRange: Range containing column names in desired order
'   removeExtra: Optional - Remove columns not in list (default: False)
'--------------------------------------------------------------------------------
Sub SyncColumnsWithList(targetSheet As Worksheet, _
                       listSheet As Worksheet, _
                       listRange As String, _
                       Optional removeExtra As Boolean = False)
    
    On Error GoTo ErrorHandler
    
    Dim masterList As Range
    Dim cell As Range
    Dim columnName As String
    Dim currentPosition As Long
    Dim targetPosition As Long
    Dim addedCount As Long
    Dim removedCount As Long
    Dim movedCount As Long
    
    ' Get master list
    Set masterList = listSheet.Range(listRange)
    
    ' Disable screen updating
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Build collection of desired columns
    Dim desiredColumns As Collection
    Set desiredColumns = New Collection
    
    For Each cell In masterList
        If Trim(cell.Value) <> "" Then
            desiredColumns.Add Trim(cell.Value)
        End If
    Next cell
    
    ' Step 1: Add missing columns and reorder
    targetPosition = 1
    Dim colName As Variant
    For Each colName In desiredColumns
        columnName = CStr(colName)
        currentPosition = GetColumnNumber(targetSheet, columnName)
        
        If currentPosition = 0 Then
            ' Column doesn't exist, insert it
            targetSheet.Columns(targetPosition).Insert Shift:=xlToRight
            targetSheet.Cells(1, targetPosition).Value = columnName
            addedCount = addedCount + 1
        ElseIf currentPosition <> targetPosition Then
            ' Column exists but in wrong position, move it
            targetSheet.Columns(currentPosition).Cut
            targetSheet.Columns(targetPosition).Insert Shift:=xlToRight
            movedCount = movedCount + 1
        End If
        
        targetPosition = targetPosition + 1
    Next colName
    
    ' Step 2: Remove extra columns if requested
    If removeExtra Then
        Dim lastColumn As Long
        lastColumn = targetSheet.Cells(1, targetSheet.Columns.Count).End(xlToLeft).Column
        
        Dim col As Long
        For col = lastColumn To targetPosition Step -1
            columnName = targetSheet.Cells(1, col).Value
            
            If Not IsInCollection(desiredColumns, columnName) Then
                targetSheet.Columns(col).Delete
                removedCount = removedCount + 1
            End If
        Next col
    End If
    
    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Synchronization complete!" & vbCrLf & vbCrLf & _
           "Added: " & addedCount & vbCrLf & _
           "Moved: " & movedCount & vbCrLf & _
           "Removed: " & removedCount, _
           vbInformation, "Sync Columns"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error synchronizing columns: " & Err.Description, vbCritical, "Error"
End Sub

'--------------------------------------------------------------------------------
' IsInCollection
' Helper function to check if value exists in collection
'--------------------------------------------------------------------------------
Function IsInCollection(col As Collection, value As String) As Boolean
    On Error Resume Next
    Dim item As Variant
    For Each item In col
        If UCase(Trim(item)) = UCase(Trim(value)) Then
            IsInCollection = True
            Exit Function
        End If
    Next item
    IsInCollection = False
End Function

'--------------------------------------------------------------------------------
' HideColumnsFromList
' Hide columns based on list (instead of deleting)
'
' Parameters:
'   targetSheet: Worksheet where columns will be hidden
'   listSheet: Worksheet containing list of column names to hide
'   listRange: Range containing column names
'--------------------------------------------------------------------------------
Sub HideColumnsFromList(targetSheet As Worksheet, _
                       listSheet As Worksheet, _
                       listRange As String)
    
    On Error GoTo ErrorHandler
    
    Dim columnList As Range
    Dim cell As Range
    Dim columnName As String
    Dim columnNumber As Long
    Dim counter As Long
    
    Set columnList = listSheet.Range(listRange)
    
    Application.ScreenUpdating = False
    
    For Each cell In columnList
        If Trim(cell.Value) <> "" Then
            columnName = Trim(cell.Value)
            columnNumber = GetColumnNumber(targetSheet, columnName)
            
            If columnNumber > 0 Then
                targetSheet.Columns(columnNumber).Hidden = True
                counter = counter + 1
            End If
        End If
    Next cell
    
    Application.ScreenUpdating = True
    
    MsgBox counter & " column(s) hidden successfully!", vbInformation, "Hide Columns"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error hiding columns: " & Err.Description, vbCritical, "Error"
End Sub

'--------------------------------------------------------------------------------
' ShowColumnsFromList
' Show (unhide) columns based on list
'
' Parameters:
'   targetSheet: Worksheet where columns will be shown
'   listSheet: Worksheet containing list of column names to show
'   listRange: Range containing column names
'--------------------------------------------------------------------------------
Sub ShowColumnsFromList(targetSheet As Worksheet, _
                       listSheet As Worksheet, _
                       listRange As String)
    
    On Error GoTo ErrorHandler
    
    Dim columnList As Range
    Dim cell As Range
    Dim columnName As String
    Dim columnNumber As Long
    Dim counter As Long
    
    Set columnList = listSheet.Range(listRange)
    
    Application.ScreenUpdating = False
    
    For Each cell In columnList
        If Trim(cell.Value) <> "" Then
            columnName = Trim(cell.Value)
            columnNumber = GetColumnNumber(targetSheet, columnName)
            
            If columnNumber > 0 Then
                targetSheet.Columns(columnNumber).Hidden = False
                counter = counter + 1
            End If
        End If
    Next cell
    
    Application.ScreenUpdating = True
    
    MsgBox counter & " column(s) shown successfully!", vbInformation, "Show Columns"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error showing columns: " & Err.Description, vbCritical, "Error"
End Sub

'================================================================================
' EXAMPLE USAGE PROCEDURES
'================================================================================

'--------------------------------------------------------------------------------
' Example1_InsertColumns
' Example showing how to insert columns after column C
'--------------------------------------------------------------------------------
Sub Example1_InsertColumns()
    ' Insert columns from list in Sheet2 (range A2:A10) after column C in Sheet1
    Call InsertColumnsFromList(Sheet1, Sheet2, "A2:A10", "C", True)
End Sub

'--------------------------------------------------------------------------------
' Example2_RemoveColumns
' Example showing how to remove columns
'--------------------------------------------------------------------------------
Sub Example2_RemoveColumns()
    ' Remove columns listed in Sheet2 (range A2:A10) from Sheet1
    Call RemoveColumnsFromList(Sheet1, Sheet2, "A2:A10", True)
End Sub

'--------------------------------------------------------------------------------
' Example3_InsertAtSpecificPositions
' Example showing how to insert columns at specific positions
'--------------------------------------------------------------------------------
Sub Example3_InsertAtSpecificPositions()
    ' Create dictionary of column names and positions
    Dim positions As Object
    Set positions = CreateObject("Scripting.Dictionary")
    
    positions.Add "New Column 1", 3
    positions.Add "New Column 2", 5
    positions.Add "New Column 3", 7
    
    Call InsertColumnsAtPosition(Sheet1, positions, True)
End Sub

'--------------------------------------------------------------------------------
' Example4_SyncWithMasterList
' Example showing how to synchronize columns with master list
'--------------------------------------------------------------------------------
Sub Example4_SyncWithMasterList()
    ' Sync Sheet1 columns with master list in Sheet2 (A2:A20)
    ' This will add missing columns, reorder existing ones, and remove extras
    Call SyncColumnsWithList(Sheet1, Sheet2, "A2:A20", True)
End Sub

'--------------------------------------------------------------------------------
' Example5_HideUnhideColumns
' Example showing how to hide/unhide columns
'--------------------------------------------------------------------------------
Sub Example5_HideUnhideColumns()
    ' Hide columns listed in Sheet2 (range B2:B10)
    Call HideColumnsFromList(Sheet1, Sheet2, "B2:B10")
    
    ' Show columns listed in Sheet2 (range C2:C10)
    Call ShowColumnsFromList(Sheet1, Sheet2, "C2:C10")
End Sub

'================================================================================
' USER INTERFACE PROCEDURES
'================================================================================

'--------------------------------------------------------------------------------
' UI_InsertColumnsWithDialog
' User-friendly procedure with input dialogs
'--------------------------------------------------------------------------------
Sub UI_InsertColumnsWithDialog()
    On Error GoTo ErrorHandler
    
    Dim targetSheetName As String
    Dim listSheetName As String
    Dim listRange As String
    Dim afterColumn As String
    Dim targetSheet As Worksheet
    Dim listSheet As Worksheet
    
    ' Get target sheet name
    targetSheetName = InputBox("Enter the name of the worksheet where columns will be inserted:", _
                               "Target Worksheet", ActiveSheet.Name)
    If targetSheetName = "" Then Exit Sub
    
    ' Get list sheet name
    listSheetName = InputBox("Enter the name of the worksheet containing the column list:", _
                            "List Worksheet", "Sheet2")
    If listSheetName = "" Then Exit Sub
    
    ' Get list range
    listRange = InputBox("Enter the range containing column names (e.g., A2:A10):", _
                        "Column List Range", "A2:A10")
    If listRange = "" Then Exit Sub
    
    ' Get after column
    afterColumn = InputBox("Enter the column letter after which to insert (e.g., C):", _
                          "Insert After Column", "C")
    If afterColumn = "" Then Exit Sub
    
    ' Validate sheets exist
    On Error Resume Next
    Set targetSheet = ThisWorkbook.Worksheets(targetSheetName)
    If targetSheet Is Nothing Then
        MsgBox "Target worksheet '" & targetSheetName & "' not found!", vbExclamation
        Exit Sub
    End If
    
    Set listSheet = ThisWorkbook.Worksheets(listSheetName)
    If listSheet Is Nothing Then
        MsgBox "List worksheet '" & listSheetName & "' not found!", vbExclamation
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    ' Execute
    Call InsertColumnsFromList(targetSheet, listSheet, listRange, afterColumn, True)
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

'--------------------------------------------------------------------------------
' UI_RemoveColumnsWithDialog
' User-friendly procedure with input dialogs for removing columns
'--------------------------------------------------------------------------------
Sub UI_RemoveColumnsWithDialog()
    On Error GoTo ErrorHandler
    
    Dim targetSheetName As String
    Dim listSheetName As String
    Dim listRange As String
    Dim targetSheet As Worksheet
    Dim listSheet As Worksheet
    
    ' Get target sheet name
    targetSheetName = InputBox("Enter the name of the worksheet where columns will be removed:", _
                               "Target Worksheet", ActiveSheet.Name)
    If targetSheetName = "" Then Exit Sub
    
    ' Get list sheet name
    listSheetName = InputBox("Enter the name of the worksheet containing the column list:", _
                            "List Worksheet", "Sheet2")
    If listSheetName = "" Then Exit Sub
    
    ' Get list range
    listRange = InputBox("Enter the range containing column names to remove (e.g., A2:A10):", _
                        "Column List Range", "A2:A10")
    If listRange = "" Then Exit Sub
    
    ' Validate sheets exist
    On Error Resume Next
    Set targetSheet = ThisWorkbook.Worksheets(targetSheetName)
    If targetSheet Is Nothing Then
        MsgBox "Target worksheet '" & targetSheetName & "' not found!", vbExclamation
        Exit Sub
    End If
    
    Set listSheet = ThisWorkbook.Worksheets(listSheetName)
    If listSheet Is Nothing Then
        MsgBox "List worksheet '" & listSheetName & "' not found!", vbExclamation
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    ' Execute
    Call RemoveColumnsFromList(targetSheet, listSheet, listRange, True)
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub
