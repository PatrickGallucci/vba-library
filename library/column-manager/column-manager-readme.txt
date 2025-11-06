EXCEL VBA COLUMN MANAGER - USER GUIDE
================================================================================
Created: November 5, 2025
Purpose: Insert and Remove Columns Based on List Values
================================================================================

TABLE OF CONTENTS
-----------------
1. Quick Start Guide
2. Main Procedures Overview
3. Detailed Usage Instructions
4. Example Scenarios
5. Setup Instructions
6. Troubleshooting
7. Advanced Features
8. Best Practices

================================================================================
1. QUICK START GUIDE
================================================================================

STEP 1: INSTALL THE VBA CODE
-----------------------------
1. Open your Excel workbook
2. Press ALT + F11 to open VBA Editor
3. Click Insert → Module
4. Copy and paste all code from Column_Manager_VBA.bas
5. Press CTRL + S to save
6. Close VBA Editor (ALT + Q)

STEP 2: PREPARE YOUR WORKBOOK
------------------------------
Your workbook should have:
- A TARGET SHEET: Where you want to insert/remove columns
- A LIST SHEET: Contains the list of column names

Example Structure:
Sheet1 (Target) - Your data with columns: ID, Name, Age, Status
Sheet2 (List)   - List of new columns to add:
                  A1: Column Name
                  A2: Email
                  A3: Phone
                  A4: Department
                  A5: Hire Date

STEP 3: RUN THE MACRO
----------------------
Option A - Use Dialog Interface (Recommended for beginners):
1. Press ALT + F8
2. Select "UI_InsertColumnsWithDialog" or "UI_RemoveColumnsWithDialog"
3. Click Run
4. Follow the prompts

Option B - Use Direct Procedures:
1. Customize the Example procedures
2. Press ALT + F8
3. Select your example (e.g., "Example1_InsertColumns")
4. Click Run

================================================================================
2. MAIN PROCEDURES OVERVIEW
================================================================================

CORE PROCEDURES:
----------------

InsertColumnsFromList
  Purpose: Insert new columns after a specified column
  Parameters:
    - targetSheet: Where to insert
    - listSheet: Where the list is
    - listRange: Range with column names (e.g., "A2:A10")
    - afterColumn: Insert after this column (e.g., "C" or 3)
    - copyFormat: Copy formatting from adjacent column (True/False)

RemoveColumnsFromList
  Purpose: Delete columns based on a list
  Parameters:
    - targetSheet: Where to delete from
    - listSheet: Where the list is
    - listRange: Range with column names to delete
    - confirmDelete: Show confirmation dialog (True/False)

InsertColumnsAtPosition
  Purpose: Insert columns at specific positions
  Parameters:
    - targetSheet: Where to insert
    - columnPositions: Dictionary with {ColumnName: Position}
    - copyFormat: Copy formatting (True/False)

SyncColumnsWithList
  Purpose: Synchronize columns with master list
  Parameters:
    - targetSheet: Sheet to sync
    - listSheet: Master list location
    - listRange: Range with desired columns
    - removeExtra: Delete columns not in list (True/False)

HELPER FUNCTIONS:
-----------------
ColumnExists(ws, columnName) - Check if column exists
GetColumnNumber(ws, columnName) - Get column number by name
GetColumnLetter(columnNumber) - Convert number to letter
IsInCollection(collection, value) - Check if value in collection

VISIBILITY PROCEDURES:
----------------------
HideColumnsFromList - Hide columns (don't delete)
ShowColumnsFromList - Unhide columns

================================================================================
3. DETAILED USAGE INSTRUCTIONS
================================================================================

SCENARIO A: INSERT COLUMNS AFTER COLUMN C
------------------------------------------

Setup:
  Sheet1: Your data sheet with existing columns
  Sheet2: List of new columns in range A2:A10

Code:
  Sub MyInsertColumns()
      Call InsertColumnsFromList(Sheet1, Sheet2, "A2:A10", "C", True)
  End Sub

What Happens:
1. Reads column names from Sheet2, cells A2:A10
2. Inserts each column after column C in Sheet1
3. Copies formatting from column C
4. Skips columns that already exist
5. Shows success message

SCENARIO B: REMOVE SPECIFIC COLUMNS
------------------------------------

Setup:
  Sheet1: Your data sheet
  Sheet2: List of columns to delete in range B2:B5

Code:
  Sub MyRemoveColumns()
      Call RemoveColumnsFromList(Sheet1, Sheet2, "B2:B5", True)
  End Sub

What Happens:
1. Reads column names from Sheet2, cells B2:B5
2. Shows confirmation dialog
3. Deletes matching columns from Sheet1
4. Shows summary of deleted columns

SCENARIO C: INSERT COLUMNS AT SPECIFIC POSITIONS
-------------------------------------------------

Setup:
  Sheet1: Your data sheet

Code:
  Sub MyInsertAtPositions()
      Dim positions As Object
      Set positions = CreateObject("Scripting.Dictionary")
      
      ' Add columns: {Name: Position}
      positions.Add "Email", 3        ' Insert at column C
      positions.Add "Phone", 5        ' Insert at column E
      positions.Add "Department", 7   ' Insert at column G
      
      Call InsertColumnsAtPosition(Sheet1, positions, True)
  End Sub

What Happens:
1. Creates dictionary with column names and positions
2. Inserts columns at exact positions specified
3. Copies formatting from previous column
4. Shows success message

SCENARIO D: SYNCHRONIZE WITH MASTER LIST
-----------------------------------------

Setup:
  Sheet1: Your current data sheet (may have extra/missing columns)
  Sheet2: Master list of desired columns in A1:A20 (in order)

Code:
  Sub MySyncColumns()
      Call SyncColumnsWithList(Sheet1, Sheet2, "A1:A20", True)
  End Sub

What Happens:
1. Reads desired column order from Sheet2
2. Adds any missing columns
3. Reorders existing columns to match list
4. Removes columns not in list (if removeExtra = True)
5. Shows summary: Added, Moved, Removed

SCENARIO E: HIDE/UNHIDE COLUMNS
--------------------------------

Setup:
  Sheet1: Your data sheet
  Sheet2: Columns to hide in B2:B10
  Sheet2: Columns to show in C2:C10

Code:
  Sub MyHideShowColumns()
      ' Hide columns
      Call HideColumnsFromList(Sheet1, Sheet2, "B2:B10")
      
      ' Show columns
      Call ShowColumnsFromList(Sheet1, Sheet2, "C2:C10")
  End Sub

What Happens:
1. Hides columns listed in B2:B10
2. Unhides columns listed in C2:C10
3. Columns are not deleted, just hidden

================================================================================
4. EXAMPLE SCENARIOS WITH FULL CODE
================================================================================

EXAMPLE 1: ADDING PROJECT MANAGEMENT COLUMNS
---------------------------------------------

Context: You have a task list and want to add project management fields

Sheet1 (Tasks): ID, Task Name, Status
Sheet2 (NewFields): 
  A2: Assigned To
  A3: Due Date
  A4: Priority
  A5: Estimated Hours
  A6: Actual Hours

Code:
  Sub AddProjectFields()
      ' Insert new PM fields after "Task Name" column (column B)
      Call InsertColumnsFromList(Sheet1, Sheet2, "A2:A6", "B", True)
  End Sub

Result:
  Sheet1 now has: ID, Task Name, Assigned To, Due Date, Priority, 
                  Estimated Hours, Actual Hours, Status

EXAMPLE 2: REMOVING TEMPORARY COLUMNS
--------------------------------------

Context: You added temporary calculation columns and want to clean up

Sheet1 (Data): Has columns including Temp1, Temp2, Calc1, Calc2
Sheet2 (ToRemove):
  A2: Temp1
  A3: Temp2
  A4: Calc1
  A5: Calc2

Code:
  Sub RemoveTemporaryColumns()
      Call RemoveColumnsFromList(Sheet1, Sheet2, "A2:A5", True)
  End Sub

Result:
  All temporary columns deleted from Sheet1

EXAMPLE 3: QUARTERLY REPORT - HIDE DETAIL COLUMNS
--------------------------------------------------

Context: Creating executive summary by hiding detailed columns

Code:
  Sub PrepareExecutiveSummary()
      ' Hide detail columns
      Dim detailColumns As Object
      Set detailColumns = CreateObject("Scripting.Dictionary")
      
      ' Create list in Sheet2
      Worksheets("Sheet2").Range("A1").Value = "Detail1"
      Worksheets("Sheet2").Range("A2").Value = "Detail2"
      Worksheets("Sheet2").Range("A3").Value = "Notes"
      Worksheets("Sheet2").Range("A4").Value = "Calculations"
      
      Call HideColumnsFromList(Sheet1, Sheet2, "A1:A4")
      
      MsgBox "Executive summary ready!", vbInformation
  End Sub

EXAMPLE 4: STANDARDIZE COLUMN ORDER
------------------------------------

Context: Multiple files with columns in different orders

Sheet1 (YourData): Columns in random order
Sheet2 (Standard): Desired column order in A1:A15

Code:
  Sub StandardizeColumns()
      ' This will add missing, reorder existing, remove extra
      Call SyncColumnsWithList(Sheet1, Sheet2, "A1:A15", True)
      
      ' Auto-fit columns after standardization
      Sheet1.Columns.AutoFit
  End Sub

EXAMPLE 5: CONDITIONAL COLUMN MANAGEMENT
-----------------------------------------

Context: Add columns based on user type

Code:
  Sub AddColumnsBasedOnUserType()
      Dim userType As String
      Dim columnRange As String
      
      userType = InputBox("Enter user type (Admin/Manager/User):", _
                         "User Type", "User")
      
      Select Case UCase(userType)
          Case "ADMIN"
              columnRange = "A2:A10"  ' All columns
          Case "MANAGER"
              columnRange = "A2:A6"   ' Some columns
          Case "USER"
              columnRange = "A2:A4"   ' Basic columns
          Case Else
              MsgBox "Invalid user type!", vbExclamation
              Exit Sub
      End Select
      
      Call InsertColumnsFromList(Sheet1, Sheet2, columnRange, "B", True)
  End Sub

================================================================================
5. SETUP INSTRUCTIONS
================================================================================

WORKBOOK STRUCTURE SETUP
-------------------------

1. CREATE TARGET SHEET (e.g., Sheet1)
   - This is your main data sheet
   - Should have header row in Row 1
   - Example:
     A1: ID
     B1: Name
     C1: Age
     D1: Status

2. CREATE LIST SHEET (e.g., Sheet2)
   - This contains your column lists
   - Can have multiple lists in different columns
   - Example:
     Column A - Columns to Add:
       A1: Column Name
       A2: Email
       A3: Phone
       A4: Department
     
     Column B - Columns to Remove:
       B1: Old Columns
       B2: TempField1
       B3: TempField2
     
     Column C - Standard Order:
       C1: ID
       C2: Name
       C3: Email
       C4: Phone
       C5: Department
       C6: Age
       C7: Status

3. INSTALL VBA CODE
   - Press ALT + F11
   - Insert → Module
   - Paste the VBA code
   - Save the workbook as .xlsm (macro-enabled)

4. CREATE CUSTOM BUTTONS (OPTIONAL)
   - Developer tab → Insert → Button
   - Draw button on sheet
   - Assign macro (e.g., UI_InsertColumnsWithDialog)
   - Label button appropriately

NAMING CONVENTIONS
------------------
- Use descriptive column names
- Avoid special characters in column names
- Keep column names unique
- Consistent capitalization helps matching

================================================================================
6. TROUBLESHOOTING
================================================================================

PROBLEM: "Column already exists" message in Debug window
SOLUTION: This is normal - the macro skips existing columns. Check Debug 
          window (CTRL + G) to see which columns were skipped.

PROBLEM: "Error inserting columns: Invalid procedure call or argument"
SOLUTION: 
  - Check that sheet names are correct
  - Verify list range exists (e.g., "A2:A10")
  - Ensure afterColumn is valid (e.g., "C" not "C1")

PROBLEM: "Target worksheet not found"
SOLUTION:
  - Check exact spelling of sheet name (case-sensitive)
  - Verify sheet exists in current workbook
  - Use exact name, including spaces

PROBLEM: Columns inserted in wrong position
SOLUTION:
  - Verify afterColumn parameter
  - Remember: "C" inserts after column C (becomes new column D)
  - Use column letters, not cell addresses

PROBLEM: Formatting not copied
SOLUTION:
  - Check that copyFormat parameter is True
  - Ensure source column has formatting to copy
  - Column width must be set manually after if needed

PROBLEM: "Permission denied" or "Method failed" error
SOLUTION:
  - Ensure worksheet is not protected
  - Close any open dialogs
  - Check if another process is using the workbook

PROBLEM: Macro runs but nothing happens
SOLUTION:
  - Check Application.ScreenUpdating wasn't left False
  - Add this line at end: Application.ScreenUpdating = True
  - Check if error handler was triggered (check Immediate window)

PROBLEM: "Type mismatch" error
SOLUTION:
  - Ensure all parameters are correct type
  - Sheet references should be Worksheet objects
  - Strings should be in quotes
  - Numbers should not be in quotes

PROBLEM: Columns not found when they exist
SOLUTION:
  - Column names must match exactly (case-insensitive but spelling matters)
  - Check for extra spaces in column names
  - Verify headers are in Row 1
  - Use Trim() on your column names in list

================================================================================
7. ADVANCED FEATURES
================================================================================

FEATURE 1: BATCH OPERATIONS
----------------------------

Insert multiple column sets with one click:

Sub BatchInsertColumns()
    ' Insert customer fields
    Call InsertColumnsFromList(Sheet1, Sheet2, "A2:A5", "B", True)
    
    ' Insert sales fields
    Call InsertColumnsFromList(Sheet1, Sheet2, "B2:B8", "E", True)
    
    ' Insert admin fields
    Call InsertColumnsFromList(Sheet1, Sheet2, "C2:C4", "K", True)
    
    MsgBox "All column sets inserted!", vbInformation
End Sub

FEATURE 2: CONDITIONAL SYNCHRONIZATION
---------------------------------------

Sync only if user confirms:

Sub ConditionalSync()
    Dim response As VbMsgBoxResult
    
    response = MsgBox("This will reorder all columns. Continue?", _
                     vbYesNo + vbQuestion, "Confirm Sync")
    
    If response = vbYes Then
        Call SyncColumnsWithList(Sheet1, Sheet2, "A1:A20", True)
    Else
        MsgBox "Operation cancelled.", vbInformation
    End If
End Sub

FEATURE 3: LOOP THROUGH MULTIPLE SHEETS
----------------------------------------

Apply same column changes to multiple sheets:

Sub ApplyToAllSheets()
    Dim ws As Worksheet
    Dim counter As Long
    
    For Each ws In ThisWorkbook.Worksheets
        ' Skip the list sheet
        If ws.Name <> "Sheet2" Then
            Call InsertColumnsFromList(ws, Sheet2, "A2:A10", "C", True)
            counter = counter + 1
        End If
    Next ws
    
    MsgBox "Columns inserted in " & counter & " sheets!", vbInformation
End Sub

FEATURE 4: DYNAMIC RANGE DETECTION
-----------------------------------

Automatically detect the range of column names:

Sub InsertWithAutoRange()
    Dim listSheet As Worksheet
    Dim lastRow As Long
    Dim listRange As String
    
    Set listSheet = Sheet2
    
    ' Find last row with data in column A
    lastRow = listSheet.Cells(listSheet.Rows.Count, "A").End(xlUp).Row
    
    ' Build range string
    listRange = "A2:A" & lastRow
    
    ' Insert columns
    Call InsertColumnsFromList(Sheet1, listSheet, listRange, "C", True)
End Sub

FEATURE 5: ERROR LOGGING
-------------------------

Log errors to a separate sheet:

Sub InsertWithErrorLog()
    On Error GoTo ErrorHandler
    
    Call InsertColumnsFromList(Sheet1, Sheet2, "A2:A10", "C", True)
    Exit Sub
    
ErrorHandler:
    Dim logSheet As Worksheet
    Dim nextRow As Long
    
    ' Create or get error log sheet
    On Error Resume Next
    Set logSheet = ThisWorkbook.Worksheets("ErrorLog")
    If logSheet Is Nothing Then
        Set logSheet = ThisWorkbook.Worksheets.Add
        logSheet.Name = "ErrorLog"
        logSheet.Range("A1").Value = "Timestamp"
        logSheet.Range("B1").Value = "Error Description"
        logSheet.Range("C1").Value = "Error Number"
    End If
    On Error GoTo 0
    
    ' Log error
    nextRow = logSheet.Cells(logSheet.Rows.Count, 1).End(xlUp).Row + 1
    logSheet.Cells(nextRow, 1).Value = Now
    logSheet.Cells(nextRow, 2).Value = Err.Description
    logSheet.Cells(nextRow, 3).Value = Err.Number
    
    MsgBox "Error occurred. See ErrorLog sheet.", vbCritical
End Sub

FEATURE 6: UNDO FUNCTIONALITY
------------------------------

Track changes and provide undo:

Dim undoColumns As Collection

Sub InsertWithUndo()
    Set undoColumns = New Collection
    
    ' Store current columns before insert
    Dim cell As Range
    For Each cell In Sheet1.Rows(1).Cells
        If Trim(cell.Value) <> "" Then
            undoColumns.Add cell.Value
        Else
            Exit For
        End If
    Next cell
    
    ' Perform insert
    Call InsertColumnsFromList(Sheet1, Sheet2, "A2:A10", "C", True)
    
    MsgBox "Columns inserted. Run UndoColumnInsert to revert.", vbInformation
End Sub

Sub UndoColumnInsert()
    ' Implementation would remove columns not in original list
    ' This is a simplified example
    MsgBox "Undo functionality - restore original column set", vbInformation
End Sub

================================================================================
8. BEST PRACTICES
================================================================================

PLANNING
--------
✓ Plan your column structure before making changes
✓ Create a master list of all desired columns
✓ Document the purpose of each column
✓ Consider future needs when designing structure

TESTING
-------
✓ Always test on a copy of your data first
✓ Use small test datasets before applying to production
✓ Verify results after each operation
✓ Keep backup copies of important workbooks

ORGANIZATION
------------
✓ Keep column lists in a dedicated "Config" sheet
✓ Use meaningful names for lists (e.g., "StandardColumns")
✓ Document what each list is for
✓ Version control your column lists

CODE MANAGEMENT
---------------
✓ Customize the example procedures for your use cases
✓ Add comments to explain your customizations
✓ Use descriptive variable names
✓ Keep related procedures together

ERROR HANDLING
--------------
✓ Always include error handling in custom procedures
✓ Test with invalid inputs to ensure graceful failures
✓ Log errors for troubleshooting
✓ Provide clear error messages to users

PERFORMANCE
-----------
✓ Turn off screen updating for large operations
✓ Turn off calculation during column changes
✓ Restore settings after completion
✓ Process multiple changes in one operation when possible

MAINTENANCE
-----------
✓ Review and update column lists regularly
✓ Remove obsolete columns periodically
✓ Document any manual changes made outside macros
✓ Keep procedures updated with business changes

SECURITY
--------
✓ Save workbooks as .xlsm (macro-enabled)
✓ Use macro security settings appropriately
✓ Document what each macro does
✓ Only run macros from trusted sources

================================================================================
KEYBOARD SHORTCUTS REFERENCE
================================================================================

VBA Editor:
ALT + F11          - Open VBA Editor
ALT + Q            - Close VBA Editor
F5                 - Run macro
F8                 - Step through code (debugging)
CTRL + G           - Open Immediate window (for Debug.Print)
CTRL + R           - Show Project Explorer
F7                 - View code window

Excel:
ALT + F8           - Macro dialog
CTRL + S           - Save workbook
F9                 - Recalculate workbook

================================================================================
SUPPORT & RESOURCES
================================================================================

Microsoft VBA Documentation:
https://learn.microsoft.com/en-us/office/vba/api/overview/excel

Excel VBA Reference:
https://learn.microsoft.com/en-us/office/vba/api/overview/excel/object-model

VBA Fundamentals:
https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/visual-basic-language-reference

================================================================================
VERSION HISTORY
================================================================================

Version 1.0 (November 5, 2025)
- Initial release
- Core insert/remove functionality
- Sync capabilities
- Hide/show features
- User interface dialogs
- Comprehensive examples

================================================================================
END OF USER GUIDE
================================================================================
