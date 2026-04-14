'----------------------------------------------------------------------------------
' MACRO : FilterAndSave
' PURPOSE: Opens an Excel file, lets user pick a filter column + column range +
'          save folder, then splits data by each unique value and saves each
'          subset as a separate .xlsx file named  <UniqueValue>_<DD-MM-YYYY>.xlsx
'----------------------------------------------------------------------------------

Option Explicit

Sub FilterAndSave()

    '--------------------------------------------------------------------------
    ' 0. DECLARATIONS
    '--------------------------------------------------------------------------
    Dim wbSource        As Workbook
    Dim wsSource        As Worksheet
    Dim wbDest          As Workbook
    Dim wsDest          As Worksheet

    Dim filterColLetter As String
    Dim filterColIdx    As Long
    Dim fromColLetter   As String
    Dim toColLetter     As String
    Dim fromColIdx      As Long
    Dim toColIdx        As Long

    Dim saveFolder      As String
    Dim lastRow         As Long
    Dim lastCol         As Long
    Dim headerRow       As Long

    Dim uniqueVals      As Collection
    Dim cellVal         As String
    Dim v               As Variant

    Dim destRange       As Range
    Dim srcRange        As Range
    Dim copyRange       As Range

    Dim fileName        As String
    Dim todayStr        As String
    Dim i               As Long

    todayStr  = Format(Date, "DD-MM-YYYY")
    headerRow = 1                            ' Change if headers are not on row 1

    '--------------------------------------------------------------------------
    ' 1. OPEN SOURCE FILE
    '--------------------------------------------------------------------------
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title      = "Step 1 of 4 – Select the Excel file to process"
    fd.Filters.Clear
    fd.Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb"
    fd.AllowMultiSelect = False

    If fd.Show <> -1 Then
        MsgBox "No file selected. Macro cancelled.", vbExclamation, "Cancelled"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts  = False

    Set wbSource = Workbooks.Open(fd.SelectedItems(1))
    Set wsSource = wbSource.Sheets(1)           ' Uses first sheet; adjust if needed

    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSource.Cells(headerRow, wsSource.Columns.Count).End(xlToLeft).Column

    If lastRow <= headerRow Then
        MsgBox "The selected file has no data rows.", vbExclamation, "No Data"
        wbSource.Close False
        Application.ScreenUpdating = True
        Application.DisplayAlerts  = True
        Exit Sub
    End If

    '--------------------------------------------------------------------------
    ' 2. PICK FILTER COLUMN
    '--------------------------------------------------------------------------
    filterColLetter = ""
    Do While filterColLetter = ""
        filterColLetter = UCase(Trim(InputBox( _
            "Step 2 of 4 – Enter the COLUMN LETTER to filter by." & vbCrLf & _
            "(e.g.  A, B, C …  Last column with data: " & _
            ColNumToLetter(lastCol) & ")", _
            "Select Filter Column")))

        If filterColLetter = "" Then
            Dim cancelFilter As Integer
            cancelFilter = MsgBox("No column entered. Cancel the macro?", _
                                  vbYesNo + vbQuestion, "Cancel?")
            If cancelFilter = vbYes Then
                wbSource.Close False
                Application.ScreenUpdating = True
                Application.DisplayAlerts  = True
                Exit Sub
            End If
        Else
            ' Validate
            On Error Resume Next
            filterColIdx = wsSource.Columns(filterColLetter).Column
            On Error GoTo 0
            If filterColIdx = 0 Or filterColIdx > lastCol Then
                MsgBox "'" & filterColLetter & "' is not a valid column. Please try again.", _
                       vbExclamation, "Invalid Column"
                filterColLetter = ""
                filterColIdx    = 0
            End If
        End If
    Loop

    '--------------------------------------------------------------------------
    ' 3. PICK COLUMN RANGE TO SAVE
    '--------------------------------------------------------------------------
    Dim colRangeInput As String
    fromColIdx = 0
    toColIdx   = 0

    Do While fromColIdx = 0 Or toColIdx = 0
        colRangeInput = Trim(InputBox( _
            "Step 3 of 4 – Enter the FROM and TO column letters to include in saved files." & vbCrLf & _
            "Format:  FROM , TO    (e.g.   A , F )" & vbCrLf & _
            "Available columns: A  to  " & ColNumToLetter(lastCol), _
            "Select Column Range", "A , " & ColNumToLetter(lastCol)))

        If colRangeInput = "" Then
            Dim cancelRange As Integer
            cancelRange = MsgBox("No range entered. Cancel the macro?", _
                                  vbYesNo + vbQuestion, "Cancel?")
            If cancelRange = vbYes Then
                wbSource.Close False
                Application.ScreenUpdating = True
                Application.DisplayAlerts  = True
                Exit Sub
            End If
        Else
            Dim parts() As String
            parts = Split(colRangeInput, ",")
            If UBound(parts) < 1 Then
                MsgBox "Please enter two column letters separated by a comma.", _
                       vbExclamation, "Invalid Input"
            Else
                fromColLetter = UCase(Trim(parts(0)))
                toColLetter   = UCase(Trim(parts(1)))

                On Error Resume Next
                fromColIdx = wsSource.Columns(fromColLetter).Column
                toColIdx   = wsSource.Columns(toColLetter).Column
                On Error GoTo 0

                If fromColIdx = 0 Or toColIdx = 0 Or _
                   fromColIdx > lastCol Or toColIdx > lastCol Then
                    MsgBox "One or both column letters are invalid. Please try again.", _
                           vbExclamation, "Invalid Column"
                    fromColIdx = 0
                    toColIdx   = 0
                ElseIf fromColIdx > toColIdx Then
                    MsgBox "The FROM column must be to the left of (or equal to) the TO column.", _
                           vbExclamation, "Invalid Range"
                    fromColIdx = 0
                    toColIdx   = 0
                End If
            End If
        End If
    Loop

    '--------------------------------------------------------------------------
    ' 4. PICK SAVE FOLDER
    '--------------------------------------------------------------------------
    Dim folderDlg As FileDialog
    Set folderDlg = Application.FileDialog(msoFileDialogFolderPicker)
    folderDlg.Title = "Step 4 of 4 – Choose the folder to save filtered files in"

    If folderDlg.Show <> -1 Then
        MsgBox "No folder selected. Macro cancelled.", vbExclamation, "Cancelled"
        wbSource.Close False
        Application.ScreenUpdating = True
        Application.DisplayAlerts  = True
        Exit Sub
    End If

    saveFolder = folderDlg.SelectedItems(1)
    If Right(saveFolder, 1) <> "\" Then saveFolder = saveFolder & "\"

    '--------------------------------------------------------------------------
    ' 5. COLLECT UNIQUE VALUES IN FILTER COLUMN (skipping header)
    '--------------------------------------------------------------------------
    Set uniqueVals = New Collection
    On Error Resume Next
    For i = headerRow + 1 To lastRow
        cellVal = Trim(CStr(wsSource.Cells(i, filterColIdx).Value))
        If cellVal <> "" Then
            uniqueVals.Add cellVal, cellVal   ' Key = Value prevents duplicates
        End If
    Next i
    On Error GoTo 0

    If uniqueVals.Count = 0 Then
        MsgBox "No data found in the filter column. Macro cancelled.", _
               vbExclamation, "No Unique Values"
        wbSource.Close False
        Application.ScreenUpdating = True
        Application.DisplayAlerts  = True
        Exit Sub
    End If

    '--------------------------------------------------------------------------
    ' 6. LOOP THROUGH UNIQUE VALUES, FILTER & SAVE
    '--------------------------------------------------------------------------
    Dim savedCount As Long
    savedCount = 0

    For Each v In uniqueVals

        ' --- a) Create new workbook ---
        Set wbDest = Workbooks.Add
        Set wsDest = wbDest.Sheets(1)

        ' --- b) Copy header row (selected column range only) ---
        Set srcRange = wsSource.Range( _
            wsSource.Cells(headerRow, fromColIdx), _
            wsSource.Cells(headerRow, toColIdx))
        srcRange.Copy wsDest.Cells(1, 1)

        ' --- c) Copy matching data rows ---
        Dim destRow As Long
        destRow = 2

        For i = headerRow + 1 To lastRow
            If Trim(CStr(wsSource.Cells(i, filterColIdx).Value)) = CStr(v) Then
                Set copyRange = wsSource.Range( _
                    wsSource.Cells(i, fromColIdx), _
                    wsSource.Cells(i, toColIdx))
                copyRange.Copy wsDest.Cells(destRow, 1)
                destRow = destRow + 1
            End If
        Next i

        ' --- d) Auto-fit columns for readability ---
        wsDest.Cells.EntireColumn.AutoFit

        ' --- e) Build safe file name (strip illegal characters) ---
        Dim safeName As String
        safeName = CStr(v)
        Dim illegalChars As Variant
        illegalChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        Dim c As Variant
        For Each c In illegalChars
            safeName = Join(Split(safeName, c), "_")
        Next c

        fileName = saveFolder & safeName & "_" & todayStr & ".xlsx"

        ' --- f) Save and close destination workbook ---
        wbDest.SaveAs fileName, FileFormat:=xlOpenXMLWorkbook
        wbDest.Close False
        savedCount = savedCount + 1

    Next v

    '--------------------------------------------------------------------------
    ' 7. CLOSE SOURCE & NOTIFY
    '--------------------------------------------------------------------------
    wbSource.Close False
    Application.ScreenUpdating = True
    Application.DisplayAlerts  = True

    MsgBox "Done! " & savedCount & " file(s) saved to:" & vbCrLf & saveFolder, _
           vbInformation, "Macro Complete"

End Sub


'----------------------------------------------------------------------------------
' HELPER: Convert a column number to its letter(s)  e.g. 1→A, 27→AA
'----------------------------------------------------------------------------------
Private Function ColNumToLetter(colNum As Long) As String
    Dim result As String
    Dim n      As Long
    n = colNum
    Do While n > 0
        Dim remainder As Long
        remainder = (n - 1) Mod 26
        result = Chr(65 + remainder) & result
        n = (n - 1 - remainder) \ 26
    Loop
    ColNumToLetter = result
End Function
