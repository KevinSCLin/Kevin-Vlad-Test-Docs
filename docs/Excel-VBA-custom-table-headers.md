---
layout: default
title: Build a Table with Custom Headers Using VBA
nav_order: 3
---

# Customize Headers in a Large Table using Excel VBA
{: .no_toc }


This instruction will guide you on how to build a table with customized pattern table headers.
{: .fs-6 .fw-300 }

## Table of contents
{: .no_toc .text-delta }

1. TOC
{:toc}

---
## Purpose of this instruction
Sometimes you will need to build a very large table with atypical column and/or row labels.
Each label corresponds to an item in real world, such as a specimen, case number, or a serial number in increasing order. 
Excel's built-in autofill function may not recognize the pattern of the labels.
Instead of manually entering the names of each column or row, it is easier to use VBA to generate the row and column
labels for us. This table builder can serve as a reusable template.

![Note][NOTE] NOTE: Excel's row and column widths are 1048576 and 16384 respectively. You cannot build a table
larger than 1048576 by 16384.

## What is VBA
VBA is an acronym for Visual Basic for Applications, a dialect of Visual Basic embedded in Microsoft Office Suite.
It is widely used in Microsoft Excel as a simple programming language for building macros, forms, and procedures to automate repeating tasks.  

![Note][NOTE] NOTE: web-based Excel does not have VBA. You must have a local copy of Microsoft Excel installed on your computer to proceed with this instruction.
## Enable Developer Tab in Excel
If your Microsoft Excel already has developer mode enabled, skip to [Create a module in VBA](#create-a-module-in-vba).

### Windows
1. Open Microsoft Excel
2. Go to [File] > [Options] This opens the Excel Options window. This opens the **Excel Options** window.
On the sidebar at the left side, click **Customize Ribbon**.
3. Check the box labeled **Developer**. Click **OK** to save and exit.

    You should see a new tab named **Developer** at the ribbon located at the top.

    ![ExcelOptions](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/ExcelOptions.PNG?raw=true)

4. Open a new Excel spreadsheet. Save in your preferred location and save the file as Excel-Macro-Enabled Workbook (*.xlsm)

    ![warning][ALERT] WARNING: Saving the file as regular spreadsheet will result in the loss of ALL VBA codes!

### MacOS

---
## Create a module in VBA
1. Go to [Developer] > [Visual Basic] to open VBA Integrated Development Environment.

    ![VBA_IDE](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/VBA_IDE.PNG?raw=true)

2. Insert a new module by going to [Insert] > [Module]. A new module named **Module1** appears in the **Modules** folder in the [Project] window.

3. Double click on Module1. Change the name of the module by going to the [Properties] window, and replace the name with CustomRowColumnLabels.

    ![Note][NOTE] NOTE: Module names cannot contain spaces and cannot begin with a numeric character or symbols.

4. In the VBA editor window, paste the following codes.

```
Sub createRowLabels()
    Dim prefix, postfix As String
    prefix = getPrefix
    postfix = getPostfix
    Dim idStr As String: idStr = getId
    Dim idNumeric As Long
    If IsNumeric(id) Then
        idNumeric = CInt(id)
    Else
        MsgBox ("id is not a number")
        Exit Sub
    End If
    Dim countStr As String: countStr = getCount
    Dim countNumeric As String
    If IsNumeric(countStr) Then
        countNumeric = CInt(countStr)
    Else
        MsgBox ("number of rows must be a number")
        Exit Sub
    End If
    Dim rowCount As Integer: rowCount = 2
    For i = 1 To countNumeric
        ThisWorkbook.Sheets(1).Range("A" & rowCount).Value = prefix & idNumeric & postfix
        rowCount = rowCount + 1
        idNumeric = idNumeric + 1
    Next i
End Sub
Sub createColumnLabels()
    Dim prefix, postfix As String
    prefix = getPrefix
    postfix = getPostfix
    Dim idStr As String: idStr = getId
    Dim idNumeric As Long
    If IsNumeric(id) Then
        idNumeric = CInt(id)
    Else
        MsgBox ("id is not a number")
        Exit Sub
    End If
    Dim countStr As String: countStr = getCount
    Dim countNumeric As String
    If IsNumeric(countStr) Then
        countNumeric = CInt(countStr)
    Else
        MsgBox ("number of rows must be a number")
        Exit Sub
    End If
    Dim ColumnCount As Integer: ColumnCount = 2
    For i = 1 To countNumeric
        ThisWorkbook.Sheets(1).Cells(1, ColumnCount).Value = prefix & idNumeric & postfix
        ColumnCount = ColumnCount + 1
        idNumeric = idNumeric + 1
    Next i
End Sub
Function getPrefix() As String
    getPrefix = InputBox("Enter prefix", "Please enter the prefix (Leave empty if none)")
End Function
Function getPostfix() As String
    getPostfix = InputBox("Enter postfix", "Please enter the postfix (Leave empty if none)")
End Function
Function getId() As String
    getId = InputBox("Enter starting number", "Please enter the starting number")
End Function
Function getCount() As String
    getCount = InputBox("Enter number of rows to create", "Please enter the number of rows to create")
End Function
```
Here is a screenshot snippet of how the VBA editor should look like.
![VBA IDE with codes](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/VBA_codes_pasted?raw=true)

---
## Implement the row & column label macros in the spreadsheet
The fastest method is to add buttons and mapped them to the macros we created above.
1. Go to [Developer] > [Insert] > [Form Controls] > **Button**. Left-click and drag a small distance to create a new button.

    An **Assign Macro** window appears to connect the macro to this button.
    ![Assign Macro](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/assignMacro.PNG?raw=true)
2. Assign the createColumnLabels macro by double clicking on it.
3. Rename Button 1 to a different name by right-clicking on it, and edit the text inside the button
4. Repeat steps 1, 2, and 3 above to implement the row label macro

---
## How to use the macro
1. Activate the **create column labels** macro by click on the button named after it.

    ![column button](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/column_labels_button.png?raw=true)

2. Enter the prefix in the message box titled _Please enter the prefix (Leave empty if none)_. Click **OK** to continue.

    ![prefix window](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/prefix_popup.PNG?raw=true)

3. Enter the postfix in the message box titled _Please enter the postfix (Leave empty if none)_. Click **OK** to continue.

    ![postfix window](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/postfix_popup.PNG?raw=true)

4. Enter the first number in the series of columns in the message box titled _Please enter the starting number_.
You must enter a integer number and cannot leave this box empty. Click **OK** to continue.

    ![starting number window](https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/rows_popup.PNG?raw=true)

5. Enter the number of columns to create in the message box titled _Please enter the number of rows to create_.
You must enter a integer number and cannot leave this box empty. Click **OK** to continue.

Excel will generate the column labels based on the information you entered.
Below is an example of a table created using the macro above.

## Benefits of using macro
You can save a lot of time.

---

## Troubleshooting
1. Did Excel disable macros upon opening a macro-enabled spreadsheet?
2. Are you using Excel Online?
3. Did you save the file as a macro-enabled spreadsheet?
4. 
5.

[NOTE]: https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/note_icon.png?raw=true
[ALERT]: https://github.com/KevinSCLin/Kevin-Vlad-Test-Docs/blob/gh-pages/assets/images/alert_icon.png?raw=true