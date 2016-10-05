'------------- Check if the sheet is already cleaned or not --------------
Sub A_StartProcessing_run()

    Worksheets("Combined").Select
    Range("A1").Select
    Dim cSht As Worksheet
    Dim cLastRow As Long
    Set cSht = ThisWorkbook.Worksheets("Combined")
    cLastRow = cSht.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

    If cLastRow - 1 = 0 Then
        MsgBox "The 'Combined' sheet is already clean. Check 'New' sheet to see if there is any information other than header in it. If there is none, Please proceed to combining data."
    Else
        If cLastRow - 1 < 0 Then
            MsgBox "There is no header in 'Combined' sheet."
        End If
        If cLastRow - 1 > 0 Then
            Call A_StartProcessing
        End If
    End If
End Sub

'---------------------- if the sheet is not clean, clean ----------------
Sub A_StartProcessing()

    ' Delete "New" & "Existing" without alerts & make new sheets with same names
    Application.DisplayAlerts = False
    Sheets("New").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("Existing").Select
    ActiveWindow.SelectedSheets.Delete

    ' Change "Combined" to be "Existing"
    Sheets("Combined").Select
    Sheets("Combined").Name = "Existing"

    'Chang Update or No update colum in Existing to be Existing
    Sheets("Existing").Select
    Range("V2").Select
    ActiveCell.FormulaR1C1 = "Existing"

    Range("A2").Select
    Selection.End(xlDown).Select
    Dim acc As Long
    acc = ActiveCell.Row

    Range("V2:V" & acc).Select
    Selection.DataSeries Rowcol:=xlColumns, Type:=xlAutoFill, Date:=xlDay, _
        Trend:=True
    Selection.FillDown

    'Generate "New" & "Combined"
    Sheets.Add.Name = "New"
    Sheets.Add.Name = "Combined"

    ' Change headder of the two sheets newly generated
    Sheets("Existing").Select
    Rows("1:1").Select
    Selection.Copy
    Sheets("New").Select
    Rows("1:1").Select
    ActiveSheet.Paste

    Sheets("Existing").Select
    Rows("1:1").Select
    Selection.Copy
    Sheets("Combined").Select
    Rows("1:1").Select
    ActiveSheet.Paste

    ' Clean up all data from the "Original" data
    Sheets("Original").Select
    Range("A:AZ").Select
    Selection.Delete
    Range("A1").Select

    'message Please copy the new original here
    MsgBox ("Copy & Paste new data here." & vbNewLine & "" & vbNewLine & "1. Open the file with new data" & vbNewLine & "2. Ctrl + A, Ctrl + C to copy all" & vbNewLine & "3. Come back to this sheet & Ctl +V at the A1 of Original Sheet")

    ' Clear selection on current sheet
    Application.CutCopyMode = False
    With ActiveSheet
       .EnableSelection = xlNoSelection
    End With

    ' Clear autofilter
    Dim i   As Long
    For i = 1 To Worksheets.count
        With Sheets(i)
            .AutoFilterMode = False
        End With
    Next
End Sub

'___________________________________________

Sub B_Port_Original_to_New()

    Sheets("Original").Select
    Range("A1").Select
    Dim oSht As Worksheet
    Dim oLastRow As Long
    Set oSht = ThisWorkbook.Worksheets("Original")
    oLastRow = oSht.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

    If oLastRow - 1 = 0 Then
        MsgBox "Original does not contain any info."
    Else
        If oLastRow - 1 < 0 Then
            MsgBox "Original does not contain any info nor a header."
        End If
        If oLastRow - 1 > 0 Then
            Call B_Port_Original_to_New_sub
        End If
    End If
End Sub



Sub B_Port_Original_to_New_sub()

'''''''''' enter code if the Original is empty then return msgbox

    Worksheets("Original").Select
    Range(Range("A1"), Range("A1").End(xlDown)).Select
    ActiveSheet.Range(Selection, Selection.End(xlToRight)).Select

    ExecuteExcel4Macro "PATTERNS(0,0,0,TRUE,2,1,0,0)"
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    Selection.Borders.LineStyle = xlNone
    Sheets("Original").Select (False)

    ' Apply formulas to the first row of "New" Sheet
    Sheets("New").Select
    Range("A2").Formula = "=Original!C2"
    Range("B2").Formula = "=Original!A2"
    Range("C2").Formula = "=LEFT(Original!AG2,8)"
    Range("D2").Formula = "=MID(Original!AG2, 10, 30)"
    Range("E2").Formula = "=Original!B2"
    Range("F2").Formula = "=Original!AL2"
    Range("G2").Formula = "=Original!G2"
    Range("H2").Formula = "=Original!H2"
    Range("I2").Formula = "=Original!I2"
    Range("J2").Formula = "=Original!K2"
    Range("K2").Formula = "=Original!L2"
    Range("L2").Formula = "=Original!P2"
    Range("M2").Formula = "=L2*Original!P2"
    Range("N2").Formula = "=Original!R2"
    Range("O2").Formula = "=Original!T2"
    Range("P2").Formula = "=Original!W2"
    Range("Q2").Formula = "=Original!Y2"
    Range("R2").Formula = "=Original!AM2"
    Range("S2").Formula = "=Original!AN2"
    Range("T2").Formula = "=VLOOKUP(Original!AH2,Ports!$A$1:$B$397,2,0)"
    Range("U2").Formula = "=Original!AO2"

    Rows("2:2").Select
    Selection.NumberFormat = "General"

    Sheets("Original").Select
    Range("A1").Select
    Selection.End(xlDown).Select
    Dim aRow As Long
    aRow = ActiveCell.Row

    Sheets("New").Select
    Range("A2", Range("U" & aRow)).Select
    Selection.FillDown
    ActiveSheet.Calculate

    ' Trim the paranthsis at the end of the customer PO (D2)
    Sheets("New").Select

    Range("A2").Select
    Selection.End(xlDown).Select
    Dim acd As Long
    acd = ActiveCell.Row

    Dim Cel As Range
    Dim Area As Range
    Dim Word As String
    Set Area = Range("D2:D" & acd)
    Word = ")"

    Application.ScreenUpdating = False
    For Each Cel In Area
        If Cel Like "*" & Word & "*" Then
            Cel = Replace(Cel, Word, "")
           'To remove the double space that follows ..
            Cel = Replace(Cel, "  ", " ")
        End If
    Next Cel
    Application.ScreenUpdating = True


    ' select the entire thing, and copy-paste value only to eliminate formula
    Sheets("New").Select
    Range("A2").Select
    Selection.End(xlDown).Select

    Dim rSource As Excel.Range
    Dim rDestination As Excel.Range
    Set rSource = ActiveSheet.Range("A2:W" & acd)
    Set rDestination = ActiveSheet.Range("A2")

    rSource.Copy
    rDestination.Select

    Selection.PasteSpecial Paste:=xlPasteValues, _
    Operation:=xlNone, _
    SkipBlanks:=False, _
    Transpose:=False

    Range("A1").Select

    ' Clear selection on current sheet
    Application.CutCopyMode = False
    With ActiveSheet
   .EnableSelection = xlNoSelection
    End With

    ' Clear autofilter
    Dim i   As Long
    For i = 1 To Worksheets.count
        With Sheets(i)
            .AutoFilterMode = False
        End With
    Next




End Sub

'___________________________________________

Sub C_Paste_formula_and_port_filtered_data_run()

'__________________ define variables_____________________
    'Define last row of Existing
    Worksheets("Existing").Select
    Range("A1").Select
    Dim eSht As Worksheet
    Dim eLastRow As Long
    Set eSht = ThisWorkbook.Worksheets("Existing")
    eLastRow = eSht.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

    'Define last row of New
    Worksheets("New").Select
    Range("A1").Select
    Dim nSht As Worksheet
    Dim nLastRow As Long
    Set nSht = ThisWorkbook.Worksheets("New")
    nLastRow = nSht.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row


'__________________ fill formula in "Existing"_____________________

    If eLastRow - 1 > 0 Then
        Worksheets("Existing").Select
        ' Paste a formula into the first cell of formula column in "Existing"
        Range("V2").Select
        Selection.FormulaArray = _
            "=IF(ISERROR(INDEX(New!R2C1:R" & nLastRow & "C23,MATCH(1,(New!R2C3:R" & nLastRow & "C3=Existing!RC3)*(New!R2C5:R" & aca & "C5=Existing!RC5),0),19)),""No Update"",""Updated"")"

        ' Fill data from V2 - V active cell 2
        Range("V2:V" & eLastRow).Select
        Selection.DataSeries Rowcol:=xlColumns, Type:=xlAutoFill, Date:=xlDay, _
            Trend:=True
        Selection.FillDown
        Sheets("Existing").Calculate
    End If

'__________________ fill formula in "New"_____________________
    If nLastRow - 1 > 0 Then
        Worksheets("New").Select
        ' Paste a formula into the first cell of formula column in "New"
        Range("V2").Select
        Selection.FormulaArray = _
            "=IF(ISERROR(INDEX(Existing!R2C1:R" & eLastRow & "C23,MATCH(1,(Existing!R2C3:R" & eLastRow & "C3=New!RC3)*(Existing!R2C5:R" & eLastRow & "C5=New!RC5),0),19)),""New Record"",""Existing Record"")"

        ' Fill data from V2 - V active cell 2
        Range("V2:V" & nLastRow).Select
        Selection.DataSeries Rowcol:=xlColumns, Type:=xlAutoFill, Date:=xlDay, _
            Trend:=True
        Selection.FillDown
        Sheets("New").Calculate
    End If
'__________________ repeat_____________________


    Call C_port_filtered_data("Existing", "No Update", "Combined") 'part of Sub D_Port_NoUpdate_from_ExistingSheet()
    Call C_port_filtered_data("Existing", "Updated", "Combined") 'Sub E_Port_Updated_From_NewSheet()
    Call C_port_filtered_data("New", "New Record", "Combined") 'part of Sub Sub C_Port_BrandNew_from_NewSheet()

End Sub


'___________________________________________

Sub D_Compare_ready_sheet_temp()
    ' Substitute the status cells of the "New Record" with comaprison formula

    ' Generate "temp" sheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:= _
             ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    ws.Name = "temp"

    ' Change headder of the two sheets newly generated
    Sheets("Existing").Select
    Rows("1:1").Select
    Selection.Copy
    Sheets("temp").Select
    Rows("1:1").Select
    ActiveSheet.Paste

    ' Put ledger on W - AC
    Cells(1, 23) = "Old:Forming"
    Cells(1, 24) = "Old:Packing"
    Cells(1, 25) = "Old:Shipped"
    Cells(1, 26) = "Old:Sailed Date"
    Cells(1, 27) = "Old:ETA"
    Cells(1, 28) = "Old:Port"
    Cells(1, 29) = "Old:Vessel"
    Cells(1, 30) = "Forming Comparison"
    Cells(1, 31) = "Packing Comparison"
    Cells(1, 32) = "Shipped Comparison"
    Cells(1, 33) = "Sailed Date Comparison"
    Cells(1, 34) = "ETA Comparison"
    Cells(1, 35) = "Port Comparison"
    Cells(1, 36) = "Vessel Comparison"
    Cells(1, 37) = "Changes, Combined"

End Sub

'__________________________________________

Sub E_Bring_updated_lines_from_combined_run()

    Call H_Bring_updated_lines_from_combined("Combined", "Updated", "temp")

End Sub

'___________________________________________
Sub Z_MessageBox_after_date_combined()

    'message Please copy the new original here
    MsgBox ("Data Combined.")

End Sub


Sub I_Compare_Bring_comparison_data_from_existing()

    ' Define the last row number for "Existing"
    Dim rowNumExisting As Long
    Sheets("Existing").Select
    Worksheets("Existing").AutoFilterMode = False
    Range("A1").Select
    Selection.End(xlDown).Select
    rowNumExisting = ActiveCell.Row

    ' Define the last row number for "temp"
    Dim rowNumTemp As Long
    Sheets("temp").Select
    Worksheets("temp").AutoFilterMode = False
    Range("A1").Select
    Selection.End(xlDown).Select
    rowNumTemp = ActiveCell.Row

    ' Enter the matching formula to the first line of temp
    Sheets("temp").Select
    Range("W2").Select
       Selection.FormulaArray = _
           "=IF(ISERROR(INDEX(Existing!R2C1:R" & rowNumExisting & "C22,MATCH(1,(Existing!R2C3:R" & rowNumExisting & "C3=temp!RC3)*(Existing!R2C5:R" & rowNumExisting & "C5=temp!RC5),0),22)),Existing!RC[-8],""Reload"")"

    ' Fill right & down the formula
    Range("W2", Range("AC" & rowNumTemp)).Select
    Selection.FillDown
    Range("W2:W" & rowNumTemp).Select
    Selection.AutoFill Destination:=Range("W2:AC" & rowNumTemp), Type:=xlFillDefault
    Range("W2:AC" & rowNumTemp).Select
    Application.CalculateFull

End Sub



Sub I_Compare_Test_run()
' set formula as a variant
' if the formula value returns value
' return row number and column number into variable
' and enter the row number and colum number into index/match to pull forming data
' repeat until the last row
' repeat until the last column

    Call I_Compare_Test("Existing", "temp", W2)

End Sub


Sub I_Compare_Test(compareSheet, fromSheet, poNum)

    ' Define the last row number for "Existing"
    Dim rowNumExisting As Long
    Sheets("Existing").Select
    Worksheets("Existing").AutoFilterMode = False
    Range("A1").Select
    Selection.End(xlDown).Select
    rowNumExisting = ActiveCell.Row

    ' Define the last row number for "temp"
    Dim rowNumTemp As Long
    Sheets("temp").Select
    Worksheets("temp").AutoFilterMode = False
    Range("A1").Select
    Selection.End(xlDown).Select
    rowNumTemp = ActiveCell.Row

    Dim dblTolerance As Double
    Dim tmp As Range

    'Get source range
    Set tmp = Sheets(compareSheet).Range("W2")

    'Get tolerance from sheet or change this to an assignment to hard code it
    dblTolerance = Sheets(compareSheet).Range("AC13")

    'use the temporary variable to cycle through the first array
    Do Until tmp.Value = poNum

        'Use absolute function to determine if you are within tolerance and if so put match in the column
        'NOTE:  Adjust the column offset (set to 4 here) to match whichever column you want result in
        If Abs(tmp.Value - tmp.Offset(0, 2).Value) < dblTolerance Then
            tmp.Offset(0, 4).Value = "Match"
        Else
            tmp.Offset(0, 4).Value = "No Match"
        End If

        'Go to the next row
        Set tmp = tmp.Offset(1, 0)
    Loop

    'Clean up
    Set tmp = Nothing
End Sub





Sub J_Compare_get_all_data_same_to_combined()

    ' Define the last row number for "temp"
    Dim rowNumTemp As Long
    Sheets("temp").Select
    Worksheets("temp").AutoFilterMode = False
    Range("A1").Select
    Selection.End(xlDown).Select
    rowNumTemp = ActiveCell.Row

    ' Get all things that has seven data points the same and transport them into combined
    Range("V2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-7]=RC[1],RC[-6]=RC[2],RC[-5]=RC[3],RC[-4]=RC[4],RC[-3]=RC[5],RC[-2]=RC[6],RC[-1]=RC[7]),""Both the same"", ""Updated"")"
    Selection.AutoFill Destination:=Range("V2:V" & rowNumTemp), Type:=xlFillDefault
    Range("V2:V" & rowNumTemp).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=Falseß
    Application.CutCopyMode = False

    ' Copy all the same lines into combined ->
    ' Select cells minus headder and copy ->
    ' Paste Copied data to "Combined" Sheet (End)

    Dim RngToCopy As Range
    With Sheets("temp")
        If .FilterMode Then .AutoFilterMode = False     'turns it off if it was active
        .Range("A1:V" & rowNumTemp).AutoFilter    'turn it on again for that range
        With .AutoFilter.Range
            .AutoFilter Field:=22, Criteria1:= _
        "Both the same"
            On Error Resume Next
            Set RngToCopy = .Offset(1).Resize(.Rows.count - 1).SpecialCells(xlCellTypeVisible)
            On Error GoTo 0
            If Not RngToCopy Is Nothing Then RngToCopy.Copy Sheets("combined").Range("A1").End(xlDown).Offset(1, 0)
        End With
    End With

    ' Delete No-Update lines from the "temp" (no message box)
    ' Range("A2:A100").SpecialCells(xlCellTypeVisible).EntireRow.Delete
    With Sheets("temp")
        Lr = Cells(Rows.count, 1).End(xlUp).Row
        If Lr > 1 Then
            Range("A2:A" & Lr).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End If
            .AutoFilterMode = False
    End With

End Sub

Sub K_Compare_seven_data()

    ' Compare 15-21 columns in the right side
        ' in Combined, Shipped>=PO PCS : All shipped
        ' in Combined, Shipped<PO PCS & in Existing, Shipped>=PO PCS -> "Error: Old data-all shipped"

    ' For all seven data -> Compare & highlight the first row
    Range("AD2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-7]=RC[-15],"""","" Forming"")"
    Range("AE2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-7]=RC[-15],"""","" Packing"")"
    Range("AF2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-7]=RC[-15],"""","" Shipped"")"
    Range("AG2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-7]=RC[-15],"""","" Sailed Date"")"
    Range("AH2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-7]=RC[-15],"""","" ETA"")"
    Range("AI2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-7]=RC[-15],"""","" Port"")"
    Range("AJ2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-7]=RC[-15],"""","" Vessel"")"

    ' Determin the row number for the temp
    Dim rowNumTemp2 As Long
    Worksheets("temp").AutoFilterMode = False
    Range("A1").Select
    Selection.End(xlDown).Select
    rowNumTemp2 = ActiveCell.Row

    ' Fill data from AD2 - AJ rowNumTemp2
    Range("AD2:AJ" & rowNumTemp2).Select
    Selection.DataSeries Rowcol:=xlColumns, Type:=xlAutoFill, Date:=xlDay, _
        Trend:=True
    Selection.FillDown

    ' combine comparison data into one cell
    Range("AK2").Select
    ActiveCell.FormulaR1C1 = _
        "=RC[-7]&RC[-6]&RC[-5]&RC[-4]&RC[-3]&RC[-2]&RC[-1]&"" Changed"""

    ' Fill data from AD2 - AJ rowNumTemp2
    Range("AK2:AK" & rowNumTemp2).Select
    Selection.DataSeries Rowcol:=xlColumns, Type:=xlAutoFill, Date:=xlDay, _
        Trend:=True
    Selection.FillDown

    ' Eliminate Formula
    Range("AD2:AK" & rowNumTemp2).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=Falseß
    Application.CutCopyMode = False

End Sub

Sub L_Copy_all_compared_to_combined()

    ' Flood formula-eleminated change combi to V
    Range("AK2:AK" & rowNumTemp2).Select
    Selection.Copy
    Range("V2").Select
    ActiveSheet.Paste

    ' Paste Copied data to "Combined" Sheet
    Range("A2:V" & rowNumTemp2).Select
    Selection.Copy
    Sheets("Combined").Select
    Worksheets("Combined").AutoFilterMode = False
    Range("A1").Select
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    ActiveSheet.Paste

    ' Delete Temp Sheet
    Application.DisplayAlerts = False
    Sheets("temp").Select
    ActiveWindow.SelectedSheets.Delete

    ' Clear selection on current sheet
    Application.CutCopyMode = False
    With ActiveSheet
   .EnableSelection = xlNoSelection
    End With

    ' Clear autofilter
    Dim i   As Long
    For i = 1 To Worksheets.count
        With Sheets(i)
            .AutoFilterMode = False
        End With
    Next



End Sub



Sub eliminate_formula()

    ' Define the last row number for "temp"
    Dim rowNumTemp As Long
    Sheets("temp").Select
    Worksheets("temp").AutoFilterMode = False
    Range("A1").Select
    Selection.End(xlDown).Select
    rowNumTemp = ActiveCell.Row

    ' select the entire thing, and copy-paste value only to eliminate formula
    Range("W2:AC" & rowNumTemp).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=Falseß
    Application.CutCopyMode = False

End Sub

'----------------------------------------------



Sub C_port_filtered_data(fromSheet, filterCri, destSheet)

    ' Delete previous filter
    Dim i   As Long
    For i = 1 To Worksheets.count
        With Sheets(i)
            .AutoFilterMode = False
        End With
    Next

    ' Define the last row number of the line with data
    Worksheets(fromSheet).Select
    Range("A1").Select
    Dim fSht As Worksheet
    Dim fLastRow As Long
    Set fSht = ThisWorkbook.Worksheets(fromSheet)
    fLastRow = fSht.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

    ' Apply AutoFilter to sort out "filterCri" datas from "fromSheet"
    Range("V1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A1:V" & fLastRow).AutoFilter Field:=22, Criteria1:= _
        filterCri

    ' generate a message if there is no filtered data// unfinished
    If ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).count - 1 > 0 Then

        ' Select cells minus headder and assign variable
        Range("A2:V" & fLastRow).Select
        Selection.SpecialCells(xlVisible).Select
        Dim filtered As Range
        Set filtered = Selection
        Selection.Copy

        ' Select the last row +1 of the destination, and paste the values
        Worksheets(destSheet).Select
        Range("A1").Select
        Dim sht As Worksheet
        Dim LastRow As Long
        Set sht = ThisWorkbook.Worksheets(destSheet)
        LastRow = sht.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        ' Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
        ' Will return the row number of the last cell even when only a single cell in the last row has data.

        Range("A" & LastRow + 1).Select
        ActiveSheet.Paste

        ' Clear selection on current sheet
        Application.CutCopyMode = False
        With ActiveSheet
        .EnableSelection = xlNoSelection
        End With

        'Clear autofilters on the sheets used
        Worksheets(fromSheet).AutoFilterMode = False
        Worksheets(destSheet).AutoFilterMode = False

    ' if there is no filtered data, produce a message
    Else
        'Clear autofilters on the sheets used
        Worksheets(fromSheet).AutoFilterMode = False
        Worksheets(destSheet).AutoFilterMode = False
        MsgBox "No data with filter criteria of '" + filterCri + "' to copy."
    End If


End Sub

'__________________________________________
Sub H_Bring_updated_lines_from_combined(fromSheet, filterCri, destSheet)

    Call Y_port_filtered_data(fromSheet, filterCri, destSheet)

    ' Delete ported date from the Combined sheet
    Sheets(fromSheet).Select
    Range("A1").Select
    Selection.End(xlDown).Select
    Dim LastRow As Long
    LastRow = ActiveCell.Row
    ActiveSheet.AutoFilterMode = False
    ActiveSheet.Range("$A$1:$W" & LastRow).AutoFilter Field:=22, Criteria1:=filterCri
    ActiveSheet.Range("$A$1:$W$" & LastRow).Offset(1, 0).SpecialCells _
    (xlCellTypeVisible).EntireRow.Delete

    ' Clear autofilter from Combined sheet
    Worksheets(fromSheet).AutoFilterMode = False

End Sub
'___________________________________________

' Excel macro to export all VBA source code in this project to text files for proper source control versioning

' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model

Public Sub Z_ExportVisualBasicCode()

    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24

    Dim VBComponent As Object
    Dim count As Integer
    Dim path As String
    Dim directory As String
    Dim extension As String
    Dim fso As New FileSystemObject

    directory = ActiveWorkbook.path & "\VisualBasic"
    count = 0

    If Not fso.FolderExists(directory) Then
        Call fso.CreateFolder(directory)
    End If

    Set fso = Nothing

    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        Select Case VBComponent.Type
            Case ClassModule, Document
                extension = ".cls"
            Case Form
                extension = ".frm"
            Case Module
                extension = ".bas"
            Case Else
                extension = ".txt"
        End Select
        On Error Resume Next
        Err.Clear
        path = directory & "\" & VBComponent.Name & extension
        Call VBComponent.Export(path)
        If Err.Number <> 0 Then
            Call MsgBox("Failed to export " & VBComponent.Name & " to " & path, vbCritical)
        Else
            count = count + 1
            Debug.Print "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & path
        End If
        On Error GoTo 0
    Next

    Application.StatusBar = "Successfully exported " & CStr(count) & " VBA files to " & directory
    Application.OnTime Now + TimeSerial(0, 0, 10), "ClearStatusBar"

End Sub
