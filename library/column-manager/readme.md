# Excel VBA Column Manager 📊

A powerful and flexible VBA library for managing Excel columns programmatically. Insert, remove, reorder, hide, and synchronize columns based on lists with just a few lines of code.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Excel Version](https://img.shields.io/badge/Excel-2010%2B-green.svg)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![VBA](https://img.shields.io/badge/VBA-Compatible-blue.svg)](https://docs.microsoft.com/en-us/office/vba/api/overview/excel)

---

## Features

- **Insert Columns**: Add new columns from a list after any specified column
- **Remove Columns**: Delete multiple columns based on a list of names
- **Insert at Positions**: Place columns at exact positions using a dictionary
- **Synchronize Columns**: Align worksheet structure with a master column list
- **Hide/Show Columns**: Toggle column visibility without deleting data
- **Format Preservation**: Automatically copy formatting from adjacent columns
- **Smart Detection**: Prevents duplicates and handles missing columns gracefully
- **User-Friendly Dialogs**: Built-in UI for non-technical users
- **Performance Optimized**: Handles large datasets efficiently

---

## Table of Contents

- [Installation](#-installation)
- [Quick Start](#-quick-start)
- [Core Functions](#-core-functions)
- [Usage Examples](#-usage-examples)
- [Advanced Features](#-advanced-features)
- [API Reference](#-api-reference)
- [Best Practices](#-best-practices)
- [Troubleshooting](#-troubleshooting)
- [Contributing](#-contributing)
- [License](#-license)

---

## Installation

### Method 1: Import Module (Recommended)

1. Download `Column_Manager_VBA.bas` from this repository
2. Open your Excel workbook
3. Press `ALT + F11` to open the VBA Editor
4. Go to `File → Import File`
5. Select the downloaded `.bas` file
6. Save your workbook as `.xlsm` (macro-enabled)

### Method 2: Manual Copy

1. Download `Column_Manager_VBA.bas` and open it in a text editor
2. Open your Excel workbook
3. Press `ALT + F11` to open the VBA Editor
4. Go to `Insert → Module`
5. Copy and paste the entire code
6. Save your workbook as `.xlsm`

### Requirements

- Microsoft Excel 2010 or later
- Windows or Mac with VBA support
- Macro security settings must allow VBA execution

---

## Quick Start

### Basic Setup

Your workbook needs two sheets:

- **Target Sheet**: Contains your data with headers in Row 1
- **List Sheet**: Contains column names in a range

```
Sheet1 (Target):
  A1: ID | B1: Name | C1: Age

Sheet2 (List):
  A2: Email
  A3: Phone
  A4: Department
```

### Simple Insert Example

```vba
Sub QuickInsert()
    ' Insert columns from Sheet2 (A2:A4) after column C in Sheet1
    Call InsertColumnsFromList(Sheet1, Sheet2, "A2:A4", "C", True)
End Sub
```

### Simple Remove Example

```vba
Sub QuickRemove()
    ' Remove columns listed in Sheet2 (B2:B5) from Sheet1
    Call RemoveColumnsFromList(Sheet1, Sheet2, "B2:B5", True)
End Sub
```

### User-Friendly Dialog

```vba
Sub UseDialog()
    ' Run the built-in dialog for inserting columns
    Call UI_InsertColumnsWithDialog
End Sub
```

Press `ALT + F8`, select the procedure, and click **Run**.

---

## Core Functions

### InsertColumnsFromList

Insert columns from a list after a specified column.

```vba
Call InsertColumnsFromList(targetSheet, listSheet, listRange, afterColumn, copyFormat)
```

**Parameters:**

- `targetSheet` (Worksheet) - Sheet where columns will be inserted
- `listSheet` (Worksheet) - Sheet containing the list of column names
- `listRange` (String) - Range containing column names (e.g., "A2:A10")
- `afterColumn` (Variant) - Column after which to insert (e.g., "C" or 3)
- `copyFormat` (Boolean) - Copy formatting from adjacent column [Optional, default: True]

**Example:**

```vba
Call InsertColumnsFromList(Sheet1, Sheet2, "A2:A10", "C", True)
```

---

### RemoveColumnsFromList

Remove columns based on a list of names.

```vba
Call RemoveColumnsFromList(targetSheet, listSheet, listRange, confirmDelete)
```

**Parameters:**

- `targetSheet` (Worksheet) - Sheet where columns will be removed
- `listSheet` (Worksheet) - Sheet containing the list of column names
- `listRange` (String) - Range containing column names to delete
- `confirmDelete` (Boolean) - Show confirmation dialog [Optional, default: True]

**Example:**

```vba
Call RemoveColumnsFromList(Sheet1, Sheet2, "A2:A5", True)
```

---

### InsertColumnsAtPosition

Insert columns at specific positions using a dictionary.

```vba
Call InsertColumnsAtPosition(targetSheet, columnPositions, copyFormat)
```

**Parameters:**

- `targetSheet` (Worksheet) - Sheet where columns will be inserted
- `columnPositions` (Dictionary) - Dictionary object with {ColumnName: Position}
- `copyFormat` (Boolean) - Copy formatting [Optional, default: True]

**Example:**

```vba
Dim positions As Object
Set positions = CreateObject("Scripting.Dictionary")
positions.Add "Email", 3
positions.Add "Phone", 5
positions.Add "Department", 7

Call InsertColumnsAtPosition(Sheet1, positions, True)
```

---

### SyncColumnsWithList

Synchronize worksheet columns with a master list.

```vba
Call SyncColumnsWithList(targetSheet, listSheet, listRange, removeExtra)
```

**Parameters:**

- `targetSheet` (Worksheet) - Sheet to synchronize
- `listSheet` (Worksheet) - Sheet containing master column list
- `listRange` (String) - Range with desired columns in order
- `removeExtra` (Boolean) - Remove columns not in list [Optional, default: False]

**What it does:**

-  Adds missing columns
-  Reorders columns to match list
-  Removes extra columns (if removeExtra = True)

**Example:**

```vba
Call SyncColumnsWithList(Sheet1, Sheet2, "A1:A20", True)
```

---

### HideColumnsFromList / ShowColumnsFromList

Hide or show columns without deleting them.

```vba
Call HideColumnsFromList(targetSheet, listSheet, listRange)
Call ShowColumnsFromList(targetSheet, listSheet, listRange)
```

**Example:**

```vba
' Hide detail columns
Call HideColumnsFromList(Sheet1, Sheet2, "C2:C10")

' Show summary columns
Call ShowColumnsFromList(Sheet1, Sheet2, "D2:D5")
```

---

## Usage Examples

### Example 1: Add Contact Fields to Customer Database

```vba
Sub AddContactFields()
    ' Setup: Sheet1 has customer data
    '        Sheet2 (A2:A4) contains: Email, Phone, Address

    Call InsertColumnsFromList(Sheet1, Sheet2, "A2:A4", "B", True)

    MsgBox "Contact fields added successfully!"
End Sub
```

### Example 2: Remove Temporary Calculation Columns

```vba
Sub RemoveTempColumns()
    ' Setup: Sheet2 (B2:B5) contains: TempCalc1, TempCalc2, Helper1, Helper2

    Call RemoveColumnsFromList(Sheet1, Sheet2, "B2:B5", True)
End Sub
```

### Example 3: Standardize Multiple Files

```vba
Sub StandardizeAllSheets()
    ' Apply standard column structure to all sheets
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "MasterList" Then
            Call SyncColumnsWithList(ws, Worksheets("MasterList"), "A1:A15", True)
        End If
    Next ws

    MsgBox "All sheets standardized!", vbInformation
End Sub
```

### Example 4: Dynamic Column Addition Based on User Role

```vba
Sub AddColumnsBasedOnRole()
    Dim userRole As String
    Dim columnRange As String

    userRole = InputBox("Enter role (Admin/Manager/User):", "User Role", "User")

    Select Case UCase(userRole)
        Case "ADMIN"
            columnRange = "A2:A10"  ' All columns
        Case "MANAGER"
            columnRange = "A2:A6"   ' Manager columns
        Case "USER"
            columnRange = "A2:A4"   ' Basic columns
        Case Else
            MsgBox "Invalid role!", vbExclamation
            Exit Sub
    End Select

    Call InsertColumnsFromList(Sheet1, Sheet2, columnRange, "C", True)
End Sub
```

### Example 5: Hide Details for Executive Summary

```vba
Sub PrepareExecutiveSummary()
    ' Hide detailed columns for executive presentation

    ' List detail columns to hide
    With Sheet2
        .Range("C1").Value = "RawData"
        .Range("C2").Value = "InternalNotes"
        .Range("C3").Value = "Calculations"
        .Range("C4").Value = "TempFields"
    End With

    Call HideColumnsFromList(Sheet1, Sheet2, "C1:C4")

    MsgBox "Executive summary view ready!", vbInformation
End Sub
```

### Example 6: Insert Project Management Columns

```vba
Sub AddProjectManagementColumns()
    ' Add complete PM column set at specific positions

    Dim positions As Object
    Set positions = CreateObject("Scripting.Dictionary")

    ' Define where each column should be inserted
    positions.Add "Priority", 3
    positions.Add "Assigned To", 4
    positions.Add "Status", 5
    positions.Add "Due Date", 6
    positions.Add "% Complete", 7
    positions.Add "Notes", 8

    Call InsertColumnsAtPosition(Sheet1, positions, True)

    ' Format the new columns
    Sheet1.Columns(6).NumberFormat = "mm/dd/yyyy"  ' Due Date
    Sheet1.Columns(7).NumberFormat = "0%"          ' % Complete

    MsgBox "Project management columns added!", vbInformation
End Sub
```

---

## Advanced Features

### Auto-Detect List Range

```vba
Sub InsertWithAutoRange()
    Dim listSheet As Worksheet
    Dim lastRow As Long
    Dim listRange As String

    Set listSheet = Sheet2

    ' Automatically detect the last row with data
    lastRow = listSheet.Cells(listSheet.Rows.Count, "A").End(xlUp).Row
    listRange = "A2:A" & lastRow

    Call InsertColumnsFromList(Sheet1, listSheet, listRange, "C", True)
End Sub
```

### Conditional Synchronization

```vba
Sub ConditionalSync()
    Dim response As VbMsgBoxResult

    ' Ask for confirmation
    response = MsgBox("This will reorder all columns to match the master list." & vbCrLf & _
                     "Continue?", vbYesNo + vbQuestion, "Confirm Sync")

    If response = vbYes Then
        Call SyncColumnsWithList(Sheet1, Sheet2, "A1:A20", True)
    Else
        MsgBox "Synchronization cancelled.", vbInformation
    End If
End Sub
```

### Batch Operations Across Workbook

```vba
Sub BatchInsertToAllSheets()
    Dim ws As Worksheet
    Dim counter As Long

    For Each ws In ThisWorkbook.Worksheets
        ' Skip the list sheet
        If ws.Name <> "ColumnList" Then
            Call InsertColumnsFromList(ws, Worksheets("ColumnList"), _
                                      "A2:A10", "C", True)
            counter = counter + 1
        End If
    Next ws

    MsgBox "Columns inserted in " & counter & " sheets!", vbInformation
End Sub
```

### Error Logging

```vba
Sub InsertWithErrorLog()
    On Error GoTo ErrorHandler

    Call InsertColumnsFromList(Sheet1, Sheet2, "A2:A10", "C", True)
    Exit Sub

ErrorHandler:
    ' Log error to ErrorLog sheet
    Dim logSheet As Worksheet
    Dim nextRow As Long

    On Error Resume Next
    Set logSheet = ThisWorkbook.Worksheets("ErrorLog")
    If logSheet Is Nothing Then
        Set logSheet = ThisWorkbook.Worksheets.Add
        logSheet.Name = "ErrorLog"
        logSheet.Range("A1:C1").Value = Array("Timestamp", "Error", "Number")
    End If

    nextRow = logSheet.Cells(logSheet.Rows.Count, 1).End(xlUp).Row + 1
    logSheet.Cells(nextRow, 1).Value = Now
    logSheet.Cells(nextRow, 2).Value = Err.Description
    logSheet.Cells(nextRow, 3).Value = Err.Number

    MsgBox "Error logged. See ErrorLog sheet.", vbCritical
End Sub
```

---

## API Reference

### Helper Functions

#### ColumnExists

Check if a column with specified name exists.

```vba
Function ColumnExists(ws As Worksheet, columnName As String) As Boolean
```

**Returns:** `True` if column exists, `False` otherwise

**Example:**

```vba
If ColumnExists(Sheet1, "Email") Then
    MsgBox "Email column exists!"
End If
```

---

#### GetColumnNumber

Get the column number for a column with specified header name.

```vba
Function GetColumnNumber(ws As Worksheet, columnName As String) As Long
```

**Returns:** Column number if found, `0` if not found

**Example:**

```vba
Dim colNum As Long
colNum = GetColumnNumber(Sheet1, "Email")
If colNum > 0 Then
    MsgBox "Email is in column " & colNum
End If
```

---

#### GetColumnLetter

Convert column number to column letter.

```vba
Function GetColumnLetter(columnNumber As Long) As String
```

**Returns:** Column letter (e.g., "A", "AA", "XFD")

**Example:**

```vba
Dim letter As String
letter = GetColumnLetter(27)  ' Returns "AA"
```

---

#### IsInCollection

Check if a value exists in a collection.

```vba
Function IsInCollection(col As Collection, value As String) As Boolean
```

**Returns:** `True` if value exists in collection, `False` otherwise

---

### User Interface Procedures

#### UI_InsertColumnsWithDialog

User-friendly procedure with input dialogs for inserting columns.

```vba
Sub UI_InsertColumnsWithDialog()
```

**Prompts for:**

- Target worksheet name
- List worksheet name
- List range
- After which column to insert

**Usage:** Press `ALT + F8`, select procedure, click Run

---

#### UI_RemoveColumnsWithDialog

User-friendly procedure with input dialogs for removing columns.

```vba
Sub UI_RemoveColumnsWithDialog()
```

**Prompts for:**

- Target worksheet name
- List worksheet name
- List range

**Usage:** Press `ALT + F8`, select procedure, click Run

---

## Best Practices

### 1. Planning

-  Create a master list of all desired columns before making changes
-  Document the purpose of each column
-  Consider future needs when designing column structure
-  Test on a copy of your data first

### 2. Organization

-  Keep column lists in a dedicated "Config" or "Settings" sheet
-  Use meaningful names for your lists
-  Add comments explaining what each list is for
-  Version control your column structures

### 3. Error Prevention

-  Always save a backup before running macros
-  Test procedures on sample data first
-  Use `confirmDelete=True` when removing columns
-  Verify results before saving

### 4. Performance

-  The library automatically handles screen updating and calculation mode
-  For very large datasets (100K+ rows), consider processing in batches
-  Use specific ranges rather than entire columns (e.g., "A2:A10" not "A:A")

### 5. Naming Conventions

-  Use consistent capitalization for column names
-  Avoid special characters in column names
-  Keep column names unique
-  Use descriptive but concise names

### 6. Code Maintenance

-  Customize the example procedures for your specific needs
-  Add comments explaining your customizations
-  Keep related procedures together
-  Use descriptive variable names

---

## Troubleshooting

### Common Issues and Solutions

#### Issue: "Column already exists" message in Debug window

**Solution:** This is normal behavior - the macro skips existing columns to prevent duplicates. Check the Immediate window (`CTRL + G`) to see which columns were skipped.

---

#### Issue: "Error inserting columns: Invalid procedure call or argument"

**Possible causes:**

- Sheet names are misspelled (check exact spelling and case)
- List range doesn't exist or is invalid
- `afterColumn` parameter is incorrect format

**Solution:** 

- Verify sheet names match exactly
- Ensure range format is correct: "A2:A10" not "A2-A10"
- Use column letter without row number: "C" not "C1"

---

#### Issue: "Target worksheet not found"

**Solution:**

- Check exact spelling of sheet name (case-sensitive in some versions)
- Verify sheet exists in current workbook
- Use sheet's actual name, including any spaces

---

#### Issue: Columns inserted in wrong position

**Solution:**

- Remember: `afterColumn="C"` inserts AFTER column C (becomes new column D)
- Use column letters ("C") or numbers (3), not cell addresses ("C1")
- Verify the column you're inserting after exists

---

#### Issue: Formatting not copied to new columns

**Solution:**

- Ensure `copyFormat` parameter is `True`
- Verify source column has formatting to copy
- Note: Column width must be set manually afterward if needed

---

#### Issue: "Permission denied" or "Method failed" error

**Solution:**

- Unprotect the worksheet before running macros
- Close any open dialog boxes
- Ensure Excel isn't in edit mode (press ESC)
- Check if another process is accessing the workbook

---

#### Issue: Columns not found when they clearly exist

**Solution:**

- Column names must match exactly (case-insensitive but spelling matters)
- Check for extra spaces in column names
- Verify headers are in Row 1
- Try trimming spaces: Use "Email" not " Email "

---

#### Issue: Macro runs but nothing happens

**Solution:**

- Check if an error occurred silently
- Open Immediate window (`CTRL + G`) to see debug messages
- Verify `Application.ScreenUpdating` wasn't left as `False`
- Add manual refresh: `Application.ScreenUpdating = True`

---

### Debug Tips

```vba
' Enable debug messages
Debug.Print "Starting column insert..."
Debug.Print "Target: " & targetSheet.Name
Debug.Print "Range: " & listRange

' View messages in Immediate window (CTRL + G in VBA Editor)
```

---

## Contributing

Contributions are welcome! Here's how you can help:

### Reporting Bugs

1. Check existing issues to avoid duplicates
2. Include Excel version and Windows/Mac OS version
3. Provide sample code that reproduces the issue
4. Include error messages and screenshots if applicable

### Suggesting Enhancements

1. Open an issue describing the enhancement
2. Explain the use case and benefits
3. Provide examples of how it would work
4. Be open to discussion and feedback

### Pull Requests

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Make your changes
4. Add comments and documentation
5. Test thoroughly
6. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
7. Push to the branch (`git push origin feature/AmazingFeature`)
8. Open a Pull Request

### Development Guidelines

- Follow existing code style and naming conventions
- Add comments for complex logic
- Update documentation for new features
- Include example usage for new procedures
- Ensure backward compatibility when possible
- Test on multiple Excel versions if available

---

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

### MIT License Summary

```
Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
```

---

## Support

### Getting Help

- **Documentation**: Check the [User Guide](Column_Manager_VBA_User_Guide.txt) for detailed instructions
- **Quick Reference**: See [Quick Reference Card](Column_Manager_Quick_Reference.txt) for common patterns
- **Issues**: Open an issue on GitHub for bugs or questions
- **Discussions**: Use GitHub Discussions for general questions and ideas

### Resources

- [Microsoft Excel VBA Reference](https://docs.microsoft.com/en-us/office/vba/api/overview/excel)
- [VBA Language Reference](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/visual-basic-language-reference)
- [Excel VBA Tutorial](https://www.excel-easy.com/vba.html)

---

## Examples Repository

Check out the `examples` folder for additional real-world scenarios:

- `Example_CustomerDatabase.xlsm` - Customer database with contact fields
- `Example_ProjectTracking.xlsm` - Project management column setup
- `Example_FinancialReport.xlsm` - Hide/show columns for different audiences
- `Example_DataStandardization.xlsm` - Sync multiple files to standard structure

---

## Roadmap

### Current Version: 1.0

### Planned Features

- [ ] Support for column groups (outline levels)
- [ ] Column width preservation and copying
- [ ] Data validation copying to new columns
- [ ] Conditional formatting migration
- [ ] Support for filtered ranges
- [ ] Column reordering by drag-and-drop from list
- [ ] Export/import column configurations
- [ ] Integration with Power Query
- [ ] Multi-workbook synchronization
- [ ] Undo/Redo functionality

### Community Requests

Vote on features in GitHub Issues with 👍 reactions!

---

## Acknowledgments

- Inspired by common Excel automation needs in enterprise environments
- Built with best practices from the VBA developer community
- Thanks to all contributors and users providing feedback

---

## Stats

![GitHub stars](https://img.shields.io/github/stars/yourusername/excel-vba-column-manager?style=social)
![GitHub forks](https://img.shields.io/github/forks/yourusername/excel-vba-column-manager?style=social)
![GitHub issues](https://img.shields.io/github/issues/yourusername/excel-vba-column-manager)
![GitHub pull requests](https://img.shields.io/github/issues-pr/yourusername/excel-vba-column-manager)

---

## 🔗 Related Projects

- [Excel VBA Data Manager](https://github.com/example/excel-data-manager) - Comprehensive data manipulation library
- [VBA Developer Tools](https://github.com/example/vba-dev-tools) - Development utilities for VBA
- [Excel Automation Suite](https://github.com/example/excel-automation) - Enterprise automation tools

---

## 📬 Contact

- **Author**: Patrick Gallucci
- ](https://linkedin.com/in/yourprofile)

---

## ⭐ Star This Repository

If you find this library useful, please consider giving it a star! It helps others discover the project.

---

<div align="center">

**[Documentation](Column_Manager_VBA_User_Guide.txt)** • 
**[Quick Reference](Column_Manager_Quick_Reference.txt)** • 
**[Examples](#-usage-examples)** • 
**[API](#-api-reference)** • 
**[Contributing](#-contributing)**

Made with ❤️ for the Excel VBA community

</div>
