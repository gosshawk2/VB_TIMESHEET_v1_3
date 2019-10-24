Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices


'VB EXCEL MODULE v1.0 written by Daniel Goss 2018.
'To link to EXCEL and perform certain tasks and actions remotely.
'thus using VB.NET as the dashboard into EXCEL applications - NEW and Already written.
'FEB 2018

Module DanG_EXCEL_Module
    'The following needs to be in the MAIN control form also:
    Public GlobalWorkBook As Microsoft.Office.Interop.Excel.Workbook = Nothing
    Public SelectedWorkSheet As String = ""
    Public Sub Update_And_Save_Sheet(ByRef Excel_Filename As String, ByVal Sheetname As String)
        Dim xlsWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlsWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim xls As New Microsoft.Office.Interop.Excel.Application

        Dim resourcesFolder = IO.Path.GetFullPath(Application.StartupPath & "\..\..\Resources\")
        Dim fileName = Excel_Filename


        xlsWorkBook = xls.Workbooks.Open(resourcesFolder & fileName)
        xlsWorkSheet = xlsWorkBook.Sheets(Sheetname)

        xlsWorkSheet.Cells(1, 1) = Excel_Filename

        xlsWorkBook.Close()
        xls.Quit()

        MsgBox("file saved to " & resourcesFolder)
    End Sub

    Public Sub releaseObject(ByVal obj As Object)
        If IsNothing(obj) Then
            Exit Sub
        End If
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Public Function BrowseFilename(Optional DefaultFilterIndex As Integer = 1) As String
        Dim ExcelFileDlg As New OpenFileDialog()
        Dim Filename As String = ""
        Dim InitialDir As String = ""
        Dim WorkbookName As String = ""

        BrowseFilename = ""
        InitialDir = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        'frmMainForm.InitialDir = "..\PDF Files"

        If System.IO.Directory.Exists(Filename) Then
            ExcelFileDlg.InitialDirectory = Filename
        Else
            ExcelFileDlg.InitialDirectory = InitialDir
        End If

        ExcelFileDlg.Filter = "EXCEL Files (*.xls)|*.xls?" & "|CSV Files (*.csv)|*.csv" & "|All Files (*.*)|*.*"
        ExcelFileDlg.FilterIndex = DefaultFilterIndex
        ExcelFileDlg.RestoreDirectory = False
        ExcelFileDlg.ShowDialog()

        WorkbookName = ExcelFileDlg.FileName
        BrowseFilename = WorkbookName
    End Function

    Public Sub CreateNewWorkbook(ByVal NewEXCELWorkbookName As String, ByRef SheetNames() As String, Optional ByRef Message As String = "",
                       Optional ByVal UseResourceFolder As Boolean = False, Optional ByVal WorkbookPath As String = "")

        'Class not registered error
        Dim xlsWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlsWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim xlsWorkSheets As Microsoft.Office.Interop.Excel.Worksheets

        Dim xlapp As New Microsoft.Office.Interop.Excel.Application
        Dim resourcesFolder = IO.Path.GetFullPath(Application.StartupPath & "\..\..\Resources\")
        Dim FileName As String = NewEXCELWorkbookName 'PATH NAME is implied 
        Dim SheetIDX As Integer
        Dim SheetName As String = ""
        Dim TotalSheets As Integer

        If Len(NewEXCELWorkbookName) = 0 Then
            MsgBox("Error in GetExcelSheets: NO WORKBOOK NAME was SUPPLIED")
            Exit Sub
        End If

        'DO WE NEED TO TEST IF THE WORKBOOK IS ALREADY OPEN ??????????????????????????????
        xlsWorkBook = xlapp.Workbooks.Add()
        xlapp.DisplayAlerts = False

        'If Len(WorkbookPath) > 0 Then
        ' If Right(WorkbookPath, 1) = "\" Then
        'xlsWorkBook = xlapp.Workbooks.Open(WorkbookPath & FileName)
        'Else
        'xlsWorkBook = xlapp.Workbooks.Open(WorkbookPath & "\" & FileName)
        'End If
        'Else
        'If UseResourceFolder = True Then
        'xlsWorkBook = xlapp.Workbooks.Open(resourcesFolder & FileName)
        'Else
        'xlsWorkBook = xlapp.Workbooks.Open(FileName) 'Implied path
        'End If
        'End If
        TotalSheets = xlsWorkBook.Sheets.Count
        If SheetExists("", "sheet1", xlsWorkBook) Then
            xlsWorkSheet = CType(xlsWorkBook.Sheets.Item("sheet1"), Microsoft.Office.Interop.Excel.Worksheet)
            xlsWorkSheet.Name = "Summary"
        End If
        Message = "New Workbook: saved as: " & xlsWorkBook.Path & "," & NewEXCELWorkbookName
        'MsgBox("New Workbook: saved as: " & xlsWorkBook.Path & "," & NewEXCELWorkbookName)
        xlsWorkBook.SaveAs(NewEXCELWorkbookName)
        xlsWorkBook.Close()
        GlobalWorkBook = xlsWorkBook

        xlapp.Quit()
        releaseObject(xlapp)
        releaseObject(xlsWorkBook)
        releaseObject(xlsWorkSheet)


    End Sub

    Public Function xlsIsOpen(Optional ByRef MEssage As String = "") As Boolean
        xlsIsOpen = False
        Dim isAlreadyOpen As Boolean = False
        Dim xlsWorkbook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlApp As New Microsoft.Office.Interop.Excel.Application
        Dim strWorkBookName As String = ""
        Dim strWB As String = ""


        Try
            If Not IsNothing(GlobalWorkBook) Then
                'strWorkBookName = GlobalWorkBook.Name - causes an unhinged type error.
                'xlsWorkbook = GlobalWorkBook.Application.Workbooks.Item(strWorkBookName)
                'xlsWorkbook = xlApp.Workbooks.Item(GlobalWorkBook)
                xlsWorkbook = GlobalWorkBook
                If Not IsNothing(xlsWorkbook) Then
                    isAlreadyOpen = True
                    xlsWorkbook.Application.Visible = True
                End If
            End If


        Catch theException As Exception
            isAlreadyOpen = False
            MEssage = theException.ToString
        End Try
        releaseObject(xlApp)
        releaseObject(xlsWorkbook)

        Return isAlreadyOpen
    End Function

    Public Sub OpenWorkbookAndSetGlobal(ByVal WorkbookFilename As String, Optional ByVal WorkbookPath As String = "", Optional ByVal CloseBookFirst As Boolean = False,
                                        Optional ByRef ErrMessages As String = "", Optional ByVal StayOpen As Boolean = False,
                                        Optional ByVal UseResourceFolder As Boolean = False, Optional ByVal ResourcesFolder As String = "")
        Dim Filename As String = ""
        Dim xlApp As New Microsoft.Office.Interop.Excel.Application 'Creates a new INSTANCE OF EXCEL IN MEMORY  .
        Dim xlsWorkbook As Microsoft.Office.Interop.Excel.Workbook = Nothing 'DOES this create another new instance in memory ?????

        Filename = WorkbookFilename
        Try

            xlApp.DisplayAlerts = False
            If CloseBookFirst = True Then
                If xlsIsOpen() Then 'THIS function might have opened another instance of EXCEL ???? - without quitting it.
                    GlobalWorkBook.Close()
                End If
            End If
            If Len(WorkbookPath) > 0 Then
                If Right(WorkbookPath, 1) = "\" Then
                    xlsWorkbook = xlApp.Workbooks.Open(WorkbookPath & Filename)
                Else
                    xlsWorkbook = xlApp.Workbooks.Open(WorkbookPath & "\" & Filename)
                End If
            Else
                If UseResourceFolder = True Then
                    xlsWorkbook = xlApp.Workbooks.Open(ResourcesFolder & Filename)
                Else
                    xlsWorkbook = xlApp.Workbooks.Open(Filename) 'Implied path
                End If
            End If
            xlsWorkbook.Application.DisplayAlerts = False
            GlobalWorkBook = xlsWorkbook
            If StayOpen = False Then
                xlsWorkbook.Close()
                xlsWorkbook.Application.Quit()
            Else
                'Keep workbook open - make it visible !
                xlsWorkbook.Application.Visible = True
            End If
        Catch ex As Exception
            releaseObject(xlApp)
        End Try
        releaseObject(xlsWorkbook)
    End Sub

    Public Sub AddNewSheet(ByVal WorkbookName As String, ByVal OpenWorkbookName As String, ByRef NewSheetName As String, Optional ByVal UseResourceFolder As Boolean = False, Optional ByVal WorkbookPath As String = "",
                            Optional ByVal LoadSheet As Boolean = False, Optional ByVal PlaceFirst As Boolean = False, Optional ByRef ErrMessages As String = "")
        Dim xlsWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim targetWorkbook As Microsoft.Office.Interop.Excel.Workbook
        Dim Open_Workbook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlsWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim xlsWorkSheets As Microsoft.Office.Interop.Excel.Worksheets

        Dim xlapp As New Microsoft.Office.Interop.Excel.Application
        Dim resourcesFolder = IO.Path.GetFullPath(Application.StartupPath & "\..\..\Resources\")
        Dim FileName As String = WorkbookName 'PATH NAME is implied 
        Dim SheetIDX As Integer
        Dim SheetName As String = ""
        Dim TotalSheets As Integer
        Dim ret As VariantType = Nothing

        If Len(WorkbookName) = 0 Then
            MsgBox("Error in AddNewSheet: NO WORKBOOK NAME was SUPPLIED")
            Exit Sub
        End If

        If Not xlsIsOpen(ErrMessages) Then
            'xlapp = New Microsoft.Office.Interop.Excel.Application
            If Len(WorkbookPath) > 0 Then
                If Right(WorkbookPath, 1) = "\" Then
                    xlsWorkBook = xlapp.Workbooks.Open(WorkbookPath & FileName)
                Else
                    xlsWorkBook = xlapp.Workbooks.Open(WorkbookPath & "\" & FileName)
                End If
            Else
                If UseResourceFolder = True Then
                    xlsWorkBook = xlapp.Workbooks.Open(resourcesFolder & FileName)
                Else
                    xlsWorkBook = xlapp.Workbooks.Open(FileName) 'Implied path
                End If
            End If
        Else
            xlsWorkBook = GlobalWorkBook
        End If
        'DO WE NEED TO TEST IF THE WORKBOOK IS ALREADY OPEN ??????????????????????????????
        'we need xlApp as this refers to the OTHER workbook that we are copying the sheet from.
        xlsWorkBook.Application.DisplayAlerts = False
        TotalSheets = xlsWorkBook.Sheets.Count
        If SheetExists("", NewSheetName, xlsWorkBook, False, "") Then
            MsgBox("WORKBOOK already contains a SHEET NAMED: " & NewSheetName)
            Exit Sub
        End If
        targetWorkbook = xlapp.ActiveWorkbook '??????????????? will this work ? but this is the same as xlsWorkbook ???
        If Len(OpenWorkbookName) > 0 Then
            'INSERT NEW SHEET - JUST LOADED:
            Open_Workbook = xlapp.Workbooks.Open(OpenWorkbookName)
            'targetWorkbook = Open_Workbook

            'The Following is unreliable: May need to read in sheet first and then COPY to new WORKBOOK.

            If PlaceFirst = False Then
                Open_Workbook.Sheets.Item(NewSheetName).move(after:=xlsWorkBook.Sheets(TotalSheets))
            Else
                Open_Workbook.Sheets.Item(NewSheetName).move(before:=xlsWorkBook.Sheets(1))
            End If
            'Need to select the sheet from the OPEN workbook:
            'BUT NEED TO FIND THE INDEX OF THE SHEET from passing the sheet name first:
        Else
            'We just want to create a new BLANK sheet in the existing (new) workbook:
            xlsWorkSheet = CType(xlsWorkBook.Sheets.Add(), Microsoft.Office.Interop.Excel.Worksheet)
            xlsWorkSheet.Name = NewSheetName
        End If
        xlsWorkBook.Save()
        'NEEDS SAVING FIRST !
        xlsWorkBook.Close()
        xlapp.Quit()
        releaseObject(xlapp)
        releaseObject(xlsWorkBook)
        'releaseObject(xlsWorkSheet)


    End Sub

    Public Function SheetExists(ByVal WorkbookName As String, ByRef SheetName As String, ByVal OpenWorkbook As Microsoft.Office.Interop.Excel.Workbook,
                                Optional ByVal UseResourceFolder As Boolean = False, Optional ByVal WorkbookPath As String = "") As Boolean
        Dim xlsWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlsWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim xlApp As New Microsoft.Office.Interop.Excel.Application
        Dim resourcesFolder = IO.Path.GetFullPath(Application.StartupPath & "\..\..\Resources\")
        Dim FileName = WorkbookName 'PATH NAME is implied 
        Dim SheetNames As String = ""
        Dim SheetIDX As Integer


        SheetExists = False
        If Len(WorkbookName) = 0 Then
            If IsNothing(OpenWorkbook) Then
                MsgBox("Error in SheetExists: NO WORKBOOK was SUPPLIED")
                Exit Function
            Else
                xlsWorkBook = OpenWorkbook
            End If
        Else
            If Len(WorkbookPath) > 0 Then
                xlsWorkBook = xlApp.Workbooks.Open(WorkbookPath & FileName)
            Else
                If UseResourceFolder = True Then
                    xlsWorkBook = xlApp.Workbooks.Open(resourcesFolder & FileName)
                Else
                    xlsWorkBook = xlApp.Workbooks.Open(FileName) 'Implied path
                End If
            End If
        End If
        If Len(SheetName) = 0 Then
            MsgBox("No SHEET NAME SUPPLIED")
            Exit Function
        End If
        'DO WE NEED TO TEST IF THE WORKBOOK IS ALREADY OPEN ??????????????????????????????


        'xlsWorkSheet = CType(xlsWorkBook.Sheets.Item(SheetName), Microsoft.Office.Interop.Excel.Worksheet)

        For Each xlsWorkSheet In xlsWorkBook.Sheets
            If UCase(xlsWorkSheet.Name) = UCase(SheetName) Then
                SheetExists = True
                Exit Function
            End If
        Next

        'xlsWorkBook.Close()
        'xlApp.Quit()

    End Function

    Sub GetExcelSheets(ByVal WorkbookName As String, ByRef SheetNames() As String,
                       Optional ByVal UseResourceFolder As Boolean = False, Optional ByVal WorkbookPath As String = "")
        Dim xlsWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlsWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim xlsWorkSheets As Microsoft.Office.Interop.Excel.Worksheets

        Dim xlapp As New Microsoft.Office.Interop.Excel.Application
        Dim resourcesFolder = IO.Path.GetFullPath(Application.StartupPath & "\..\..\Resources\")
        Dim FileName As String = WorkbookName 'PATH NAME is implied 
        Dim SheetIDX As Integer
        Dim SheetName As String = ""
        Dim TotalSheets As Integer

        If Len(WorkbookName) = 0 Then
            MsgBox("Error in GetExcelSheets: NO WORKBOOK NAME was SUPPLIED")
            Exit Sub
        End If

        'DO WE NEED TO TEST IF THE WORKBOOK IS ALREADY OPEN ??????????????????????????????
        xlapp.DisplayAlerts = False
        If Len(WorkbookPath) > 0 Then
            If Right(WorkbookPath, 1) = "\" Then
                xlsWorkBook = xlapp.Workbooks.Open(WorkbookPath & FileName)
            Else
                xlsWorkBook = xlapp.Workbooks.Open(WorkbookPath & "\" & FileName)
            End If
        Else
            If UseResourceFolder = True Then
                xlsWorkBook = xlapp.Workbooks.Open(resourcesFolder & FileName)
            Else
                xlsWorkBook = xlapp.Workbooks.Open(FileName) 'Implied path
            End If
        End If
        TotalSheets = xlsWorkBook.Sheets.Count
        ReDim SheetNames(TotalSheets + 1)
        For SheetIDX = 1 To TotalSheets
            xlsWorkSheet = xlsWorkBook.Sheets(SheetIDX)

            'xlsWorkSheet = CType(xlsWorkBook.Worksheets(SheetIDX), Microsoft.Office.Interop.Excel.Worksheet)
            SheetNames(SheetIDX) = xlsWorkSheet.Name

        Next

        xlsWorkBook.Close()
        xlapp.Quit()
        releaseObject(xlsWorkSheet)
        xlsWorkSheet = Nothing
        releaseObject(xlsWorkBook)
        xlsWorkBook = Nothing
        releaseObject(xlapp)
        xlapp = Nothing



    End Sub


    Public Sub Create_PivotTable(ByVal WorkbookName As String, ByVal WorkbookOpen As Microsoft.Office.Interop.Excel.Workbook, ByVal SourceSheetName As String,
                                  ByVal Range_StartCol As Long, ByVal Range_StartRow As Long, ByVal Range_EndCol As Long, ByVal Range_EndRow As Long,
                                  ByVal Pos_StartCol As Long, ByVal Pos_StartRow As Long,
        ByVal PivotTableName As String, NewPivotsheetname As String, Optional ByVal OpenTheWorkbook As Boolean = False, Optional StartAdr As String = "A1", Optional ByVal UseResourceFolder As Boolean = False, Optional ByVal WorkbookPath As String = "")

        'NewPivotsheetname = "PIVOTS"
        'SourceSheetName = "RAW DATA SUMMARY"
        'Sheets.Add
        'ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=
        '    "Pick_Performance!R1C1:R1677C17", Version:=6).CreatePivotTable _
        'tabledestination:="Sheet10!R3C1", TableName:="PivotTable2", DefaultVersion _
        '    :=6
        'Sheets("Sheet10").Select
        'Cells(3, 1).Select
        'Range("B9").Select
        'With ActiveSheet.PivotTables("PivotTable2").PivotFields("SHORT AREA")
        '.Orientation = xlPageField
        '.Position = 1
        'End With
        'With ActiveSheet.PivotTables("PivotTable2").PivotFields("PICKERS")
        '.Orientation = xlRowField
        '.Position = 1
        'End With
        'ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables(
        '   "PivotTable2").PivotFields("NUMBER OF PICKS"), "Sum of NUMBER OF PICKS", xlSum
        'ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables(
        '"PivotTable2").PivotFields("PICKED QUANTITY"), "Sum of PICKED QUANTITY", xlSum
        'ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables(
        '"PivotTable2").PivotFields("PICKING SET-UP HOURS"),
        '"Sum of PICKING SET-UP HOURS", xlSum
        'ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables(
        '"PivotTable2").PivotFields("PICKING TRAVEL HOURS"),
        '"Sum of PICKING TRAVEL HOURS", xlSum
        'ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables(
        '"PivotTable2").PivotFields("PICKING WAITING TIME HOURS"),
        '"Sum of PICKING WAITING TIME HOURS", xlSum
        'With ActiveSheet.PivotTables("PivotTable2").PivotFields("NON-PROCESSING HOURS")
        '.Orientation = xlRowField
        '.Position = 2
        'End With
        'ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables(
        '   "PivotTable2").PivotFields("NON-PROCESSING HOURS"),
        '   "Count of NON-PROCESSING HOURS", xlCount
        'With ActiveSheet.PivotTables("PivotTable2").PivotFields(
        '"Count of NON-PROCESSING HOURS")
        '.Caption = "Sum of NON-PROCESSING HOURS"
        '.Function = xlSum
        'End With
        'ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables(
        '   "PivotTable2").PivotFields("PICKING PROCESSING HOURS"),
        '   "Sum of PICKING PROCESSING HOURS", xlSum
        'ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables(
        '"PivotTable2").PivotFields("PICKING TOTAL HOURS"), "Sum of PICKING TOTAL HOURS" _
        ' , xlSum
        'ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables(
        '"PivotTable2").PivotFields("PICKS PER HOUR"), "Sum of PICKS PER HOUR", xlSum
        'With ActiveSheet.PivotTables("PivotTable2").PivotFields("QTY PER HOUR")
        '.Orientation = xlRowField
        '.Position = 2
        'End With
        'ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables(
        '   "PivotTable2").PivotFields("QTY PER HOUR"), "Count of QTY PER HOUR", xlCount
        'With ActiveSheet.PivotTables("PivotTable2").PivotFields("Count of QTY PER HOUR"
        ')
        '.Caption = "Sum of QTY PER HOUR"
        '.Function = xlSum
        'End With
        'Range("L8").Select
        Dim xlsWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlsWorkSheets As Microsoft.Office.Interop.Excel.Worksheets
        Dim xlsPivotSourceWorksheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim xlApp As New Microsoft.Office.Interop.Excel.Application
        Dim resourcesFolder = IO.Path.GetFullPath(Application.StartupPath & "\..\..\Resources\")
        Dim fileName = WorkbookName 'PATH NAME is implied 
        Dim SheetNames As String = ""
        Dim SheetIDX As Integer
        Dim NewSht As Microsoft.Office.Interop.Excel.Worksheet
        Dim pvtCache As Microsoft.Office.Interop.Excel.PivotCache
        Dim pvtTable As Microsoft.Office.Interop.Excel.PivotTable
        Dim Startpvt As String
        Dim srcData As String
        Dim TimeStamp As String
        Dim Response As Integer

        TimeStamp = CStr(Now())
        TimeStamp = Replace(TimeStamp, "/", "_")
        TimeStamp = Replace(TimeStamp, ":", "_")

        xlApp.DisplayAlerts = False
        If Not IsNothing(WorkbookOpen) Then
            xlsWorkBook = WorkbookOpen
            'Grab the variable from the already opened excel - global xlsWorkbook required !
            'Assume USER wants to Open the workbook first before pressing the Insert Pivot Table button ?
        Else
            If OpenTheWorkbook Then
                'Test if workbook open already ? -- does it need to stay open ?
                If Len(WorkbookPath) > 0 Then
                    If Right(WorkbookPath, 1) = "\" Then
                        xlsWorkBook = xlApp.Workbooks.Open(WorkbookPath & fileName, False, False)
                    Else
                        xlsWorkBook = xlApp.Workbooks.Open(WorkbookPath & "\" & fileName, False, False)
                    End If
                Else
                    If UseResourceFolder = True Then
                        xlsWorkBook = xlApp.Workbooks.Open(resourcesFolder & fileName, False, False)
                    Else
                        xlsWorkBook = xlApp.Workbooks.Open(fileName, False, False) 'Implied path
                    End If
                End If
            Else
                'set xlsWorkbook (local variable) = GlobalWorkbookVariable (Globally) - defined in main Control.

            End If

        End If
        'Create the Source Data sheet from the SourceSheetName
        xlsPivotSourceWorksheet = xlsWorkBook.Sheets.Item(SourceSheetName) 'originally uses SET.
        With xlsPivotSourceWorksheet
            'ADDRESS(ROW,COL,ABS/REL,A1/R1C1=TRUE/FALSE,SheetName)
            srcData = SourceSheetName & "!" & .Range(.Cells(Range_StartRow, Range_StartCol), .Cells(Range_EndRow, Range_EndCol)).Address(ReferenceStyle:=Excel.XlReferenceStyle.xlR1C1)

        End With
        If Not SheetExists(WorkbookName, NewPivotsheetname, xlsWorkBook) Then
            NewSht = DirectCast(xlApp.Worksheets.Add(xlsWorkBook.Worksheets(1), Type.Missing, Type.Missing, Type.Missing), Excel.Worksheet)

            'xlsWorkBook.NewSht.add

            xlsWorkBook.Sheets.Item(1).Name = NewPivotsheetname
            Startpvt = NewPivotsheetname & "!" & xlsWorkBook.Sheets.Item(NewSht.Name).Range(StartAdr).Address(ReferenceStyle:=Excel.XlReferenceStyle.xlR1C1) 'EXCEL USES SET HERE.
            pvtCache = xlsWorkBook.PivotCaches.Create(SourceType:=Excel.XlPivotTableSourceType.xlDatabase, SourceData:=srcData) 'EXCEL USES SET HERE.
            pvtTable = pvtCache.CreatePivotTable(TableDestination:=Startpvt, TableName:=PivotTableName)
            Call Add_Pivot_Field(WorkbookName, Nothing, "PIVOTS", "PivotTable1", "Report Type", "Report Type", "PAGE", "A1", 1, "", False, "Report Type", "goodsin", False, "")
        Else
            'Response = MsgBox(NewPivotsheetname & " Already Exists - Create a new sheet ?", vbYesNo, "Pivot Sheet Already Exists")
            'If Response = vbYes Then
            'Set NewSht = Sheets.Add
            'NewSht.Name = NewPivotsheetname & "_" & TimeStamp
            'Else
            'Set NewSht = Sheets(NewPivotsheetname)
            'End If
            'ActiveWorkbook.Worksheets(NewPivotsheetname).Delete
            'NewSht = Sheets.Add
            NewSht = xlsWorkBook.Sheets.Item(NewPivotsheetname)
            'Startpvt = NewSht.Name & "!" & NewSht.Range("A1").Address(ReferenceStyle:=xlR1C1)
            'pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=srcData)
            Startpvt = xlsWorkBook.Sheets.Item(NewSht.Name) & "!" & xlsWorkBook.Sheets.Item(NewSht.Name).Range(StartAdr).Address(ReferenceStyle:=Excel.XlReferenceStyle.xlR1C1) 'EXCEL USES SET HERE.
            pvtCache = xlsWorkBook.PivotCaches.Create(SourceType:=Excel.XlPivotTableSourceType.xlDatabase, SourceData:=srcData)
            pvtTable = pvtCache.CreatePivotTable(TableDestination:=Startpvt, TableName:=PivotTableName)
        End If
        'xlsWorkBook.Close()
        Try
            xlsWorkBook.Save()
        Catch
            xlsWorkBook.SaveAs(xlsWorkBook.Name)
        End Try




    End Sub

    Public Sub Add_Pivot_Field(ByVal WorkbookName As String, ByVal WorkbookOpen As Microsoft.Office.Interop.Excel.Workbook, ByVal PivotSheetName As String,
                               ByVal PivotTableName As String, ByVal FieldName As String, ByVal FieldTitle As String, ByVal FieldType As String, Optional StartAdr As String = "A1",
                Optional ByVal FieldPosition As Integer = 0, Optional ByVal NumFormat As String = "", Optional ByVal manualupdate As Boolean = False,
                               Optional ByVal FuncName As String = "", Optional ByVal FuncValue As String = "", Optional ByVal UseResourceFolder As Boolean = False, Optional ByVal WorkbookPath As String = "")

        Dim xlsWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlsWorkSheets As Microsoft.Office.Interop.Excel.Worksheets
        Dim xlsPivotSourceWorksheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim xlApp As New Microsoft.Office.Interop.Excel.Application
        Dim resourcesFolder = IO.Path.GetFullPath(Application.StartupPath & "\..\..\Resources\")
        Dim fileName = WorkbookName 'PATH NAME is implied 
        Dim SheetNames As String = ""
        Dim SheetIDX As Integer
        Dim NewSht As Microsoft.Office.Interop.Excel.Worksheet
        Dim pvtCache As Microsoft.Office.Interop.Excel.PivotCache
        Dim pvtTable As Microsoft.Office.Interop.Excel.PivotTable
        Dim Startpvt As String
        Dim srcData As String

        If Not IsNothing(WorkbookOpen) Then
            xlsWorkBook = WorkbookOpen
        Else
            xlApp.DisplayAlerts = False
            If Len(WorkbookPath) > 0 Then
                If Right(WorkbookPath, 1) = "\" Then
                    xlsWorkBook = xlApp.Workbooks.Open(WorkbookPath & fileName)
                Else
                    xlsWorkBook = xlApp.Workbooks.Open(WorkbookPath & "\" & fileName)
                End If
            Else
                If UseResourceFolder = True Then
                    xlsWorkBook = xlApp.Workbooks.Open(resourcesFolder & fileName)
                Else
                    xlsWorkBook = xlApp.Workbooks.Open(fileName) 'Implied path
                End If
            End If
        End If

        'PivotSheetName = "PIVOTS"
        'Startpvt = xlsWorkBook.NewSht.Name & "!" & xlsWorkBook.NewSht.Range(StartAdr).Address(ReferenceStyle:=Excel.XlReferenceStyle.xlR1C1)
        'pvtCache = xlsWorkBook.ActiveWorkbook.PivotCaches.Create(SourceType:=xlApp.xlDatabase, SourceData:=srcData)
        'pvtTable = pvtCache.CreatePivotTable(TableDestination:=Startpvt, TableName:=PivotTableName)

        'GETTING INVALID INDEX HERE: - could be that the Pivot Table named in PivotTableName does not exist yet ?
        'or that PIVOT SHEET NAME - PIVOTS - DOES NOT EXIST ? - BECAUSE THE SHEET HAS NOT SAVED THE PIVOT TABLE YET ? - BUT SHOULD BE OPEN AND CURRENT WORKBOOK+WORKSHEET ?
        'pvtTable = xlsWorkBook.Sheets.Item(PivotSheetName).PivotTables(PivotTableName)
        pvtTable = xlsWorkBook.ActiveSheet.pivottables(PivotTableName)
        If UCase(FieldType) = "PAGE" Then
            pvtTable.PivotFields(FieldName).Orientation = xlApp.xlPageField
            pvtTable.PivotFields(FieldName).currentpage = FuncValue
        ElseIf UCase(FieldType) = "COLUMN" Then
            pvtTable.PivotFields(FieldName).Orientation = xlApp.xlColumnField
        ElseIf UCase(FieldType) = "ROW" Then
            pvtTable.PivotFields(FieldName).Orientation = xlApp.xlRowField
        ElseIf UCase(FieldType) = "SUM" Then
            pvtTable.AddDataField(pvtTable.PivotFields(FieldName), FieldTitle, xlApp.xlSum)
            'pvt.PivotFields(FieldName).Caption = FieldTitle
            'pvt.PivotFields(FieldName).Function = xlSum
        ElseIf UCase(FieldType) = "COUNT" Then
            pvtTable.AddDataField(pvtTable.PivotFields(FieldName), FieldTitle, xlApp.xlCount)
            'pvt.PivotFields(FieldName).Function = xlCount
        ElseIf UCase(FieldType) = "DATA" Then
            pvtTable.PivotFields(FieldName).Orientation = xlApp.xlDataField

        Else
            MsgBox("Field Type not recognised")
        End If
        If Len(NumFormat) > 0 Then
            pvtTable.PivotFields(FieldName).NumberFormat = NumFormat
        End If
        If FieldPosition > 0 Then
            pvtTable.PivotFields(FieldName).Position = FieldPosition
        End If
        If Len(FuncName) > 0 And UCase(FuncName) = "XLSUM" Then
            pvtTable.PivotFields(FieldName).function = Excel.XlConsolidationFunction.xlSum
        ElseIf UCase(FuncName) = "XLCOUNT" Then
            pvtTable.PivotFields(FieldName).function = Excel.XlConsolidationFunction.xlCount
        End If


    End Sub

    Public Sub Add_Calculated_Field(ByVal WorkbookName As String, ByVal WorkbookOpen As Microsoft.Office.Interop.Excel.Workbook,
                                    ByVal PivotSheetName As String, ByVal PivotTableName As String, ByVal FieldName As String, ByVal FieldTitle As String,
                                    ByVal Calculation As String, Optional ByVal FieldPosition As Integer = 0, Optional ByVal NumFormat As String = "",
                                    Optional ByVal manualupdate As Boolean = False, Optional ByVal UseResourceFolder As Boolean = False, Optional ByVal WorkbookPath As String = "")


        Dim xlsWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlApp As New Microsoft.Office.Interop.Excel.Application
        Dim resourcesFolder = IO.Path.GetFullPath(Application.StartupPath & "\..\..\Resources\")
        Dim fileName = WorkbookName 'PATH NAME is implied 
        Dim SheetNames As String = ""
        Dim pvtTable As Microsoft.Office.Interop.Excel.PivotTable

        If Not IsNothing(WorkbookOpen) Then
            xlsWorkBook = WorkbookOpen
        Else
            xlApp.DisplayAlerts = False
            If Len(WorkbookPath) > 0 Then
                If Right(WorkbookPath, 1) = "\" Then
                    xlsWorkBook = xlApp.Workbooks.Open(WorkbookPath & fileName)
                Else
                    xlsWorkBook = xlApp.Workbooks.Open(WorkbookPath & "\" & fileName)
                End If
            Else
                If UseResourceFolder = True Then
                    xlsWorkBook = xlApp.Workbooks.Open(resourcesFolder & fileName)
                Else
                    xlsWorkBook = xlApp.Workbooks.Open(fileName) 'Implied path
                End If
            End If
        End If


        pvtTable = xlsWorkBook.Worksheets(PivotSheetName).PivotTables(PivotTableName)

        'Calc = "=IF(OldField" & si & "=0,0,(NewField" & si & "-OldField" & si & ")/OldField" & si & ")"


        pvtTable.CalculatedFields.Add(FieldName, Calculation, True)

        With pvtTable.PivotFields(FieldName)
            .Orientation = xlApp.xlDataField
            .Function = xlApp.xlSum
            '.Position = FieldPosition
            .NumberFormat = NumFormat
            .Caption = FieldTitle
        End With

        'ActiveSheet.PivotTables("PivotTable1").CalculatedFields.Add "UPH", _
        '        "='PICKED QUANTITY' /'TOTAL HOURS'", True
        '    ActiveSheet.PivotTables("PivotTable1").PivotFields("UPH").Orientation = _
        '        xlDataField
        '    ActiveSheet.PivotTables("PivotTable1").CalculatedFields.Add "Downtime", _
        '        "='TOTAL HOURS' -'TOTAL PRODUCTIVE HOURS'", True
        '    ActiveSheet.PivotTables("PivotTable1").PivotFields("Downtime").Orientation = _
        '        xlDataField


    End Sub

    Public Sub CreateChart()
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")

        'add data
        xlWorkSheet.Cells(1, 1) = ""
        xlWorkSheet.Cells(1, 2) = "Student1"
        xlWorkSheet.Cells(1, 3) = "Student2"
        xlWorkSheet.Cells(1, 4) = "Student3"

        xlWorkSheet.Cells(2, 1) = "Term1"
        xlWorkSheet.Cells(2, 2) = "80"
        xlWorkSheet.Cells(2, 3) = "65"
        xlWorkSheet.Cells(2, 4) = "45"

        xlWorkSheet.Cells(3, 1) = "Term2"
        xlWorkSheet.Cells(3, 2) = "78"
        xlWorkSheet.Cells(3, 3) = "72"
        xlWorkSheet.Cells(3, 4) = "60"

        xlWorkSheet.Cells(4, 1) = "Term3"
        xlWorkSheet.Cells(4, 2) = "82"
        xlWorkSheet.Cells(4, 3) = "80"
        xlWorkSheet.Cells(4, 4) = "65"

        xlWorkSheet.Cells(5, 1) = "Term4"
        xlWorkSheet.Cells(5, 2) = "75"
        xlWorkSheet.Cells(5, 3) = "82"
        xlWorkSheet.Cells(5, 4) = "68"

        'create chart
        Dim chartPage As Excel.Chart
        Dim xlCharts As Excel.ChartObjects
        Dim myChart As Excel.ChartObject
        Dim chartRange As Excel.Range

        xlCharts = xlWorkSheet.ChartObjects
        myChart = xlCharts.Add(10, 80, 300, 250)
        chartPage = myChart.Chart
        chartRange = xlWorkSheet.Range("A1", "d5")
        chartPage.SetSourceData(Source:=chartRange)
        chartPage.ChartType = Excel.XlChartType.xlColumnClustered

        xlWorkSheet.SaveAs("C:\vbexcel.xlsx")
        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

        MsgBox("Excel file created , you can find the file c:\")
    End Sub

    Public Sub CloseWorkbook(xlsWorkbookName As Excel.Workbook)
        xlsWorkbookName.Close()

    End Sub


    Public Sub QuitExcel(xlsWorkBook As Microsoft.Office.Interop.Excel.Workbook)
        Dim xlsWorkSheets As Microsoft.Office.Interop.Excel.Worksheets
        Dim xlApp As New Microsoft.Office.Interop.Excel.Application


        xlsWorkBook.Close()
        xlApp.Quit()
        releaseObject(xlApp)
        releaseObject(xlsWorkBook)
    End Sub

    Function Get_Total_Rows_From_CSVFile(CSVFilename As String, Optional ByRef TotalFields As Long = 0) As Long
        Dim rownum As Long
        Dim TotalRows As Long
        Dim wholeRows As String()

        Using csvReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(CSVFilename)
            csvReader.TextFieldType = FileIO.FieldType.Delimited
            csvReader.SetDelimiters(",")
            TotalRows = 0
            rownum = 0
            While Not csvReader.EndOfData
                Try
                    wholeRows = csvReader.ReadFields()

                    TotalFields = wholeRows.Length
                Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                    MsgBox("Line " & ex.Message & "is NOT valid and will be skipped.")
                End Try
                rownum = rownum + 1
            End While

        End Using
        TotalRows = rownum
        Get_Total_Rows_From_CSVFile = TotalRows
    End Function

    Function CSVFileToArray(ByRef NewstrArray As String(,), ByVal CSVFilename As String, Optional ByRef TotalFields As Long = 0,
                                Optional ByRef Message As String = "") As Long
        Dim wholeRows As String()
        Dim numRowsRead As Long
        Dim colnum As Long
        Dim RowNum As Long
        Dim currentFields As String
        Dim Percentage As Long
        Dim TotalRows As Long
        Dim RowMessage As String
        Dim messages As String

        ReDim NewstrArray(1, 1)
        CSVFileToArray = 0
        numRowsRead = 0
        colnum = 0
        RowNum = 0
        messages = ""
        ReDim wholeRows(1)
        wholeRows(0) = ""
        Percentage = 0.0F
        frmMainGIForm.pbarMain.Style = ProgressBarStyle.Blocks
        frmMainGIForm.pbarMain.Value = 0
        'ProgressBarControl.Maximum = 100
        'ProgressBarControl.Value = 0
        'ProgressLabel.Text = "0%"
        'Import from CSV File TEXT field and check validity:
        Try
            Using csvReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(CSVFilename)
                csvReader.TextFieldType = FileIO.FieldType.Delimited
                csvReader.SetDelimiters(",")
                TotalRows = Get_Total_Rows_From_CSVFile(CSVFilename, TotalFields)
                ReDim NewstrArray(TotalFields, TotalRows)
                'frmProgressBar.Show()
                While Not csvReader.EndOfData
                    Try
                        wholeRows = csvReader.ReadFields()
                        colnum = 0
                        For Each currentFields In wholeRows
                            'Me.rtbOutput.Text = currentFields
                            'MsgBox(CStr(RowNum) & ": Field " & CStr(colnum) & " = " & currentFields)
                            '
                            '****************************************************
                            '
                            'AUTHOR : DANIEL GOSS . COPYRIGHT 24-FEB-2017. V1.2
                            'MODIFIED: DANIEL GOSS . COPYRIGHT 01-SEP-2017. v1.3
                            ' - added extra parameter - Messages
                            'MODIFIED: DANIEL GOSS . COPYRIGHT 03-SEP-2017. v1.4 now.
                            ' - removed extra parameter - Messages
                            ' - removed reference to GetFieldValue_From_Index below.
                            ' - added error logging and message logging 12-09-2017 01:50
                            '*****************************************************

                            If TotalRows > 0 Then
                                Percentage = CLng((RowNum / TotalRows) * 100)

                            End If
                            frmMainGIForm.pbarMain.Value = CInt(Percentage)
                            RowMessage = "Read " & CStr(RowNum) & " ROW / " & CStr(TotalRows) & " ROWS"
                            'RichTextBoxMessageControl.Text = "Reading from CSV file ..."
                            'RichTextBoxMessageControl.Text = vbCrLf & RowMessage
                            'ProgressBarControl.Value = Percentage
                            'ProgressLabel.Text = CStr(ProgressBarControl.Value) & "%"
                            NewstrArray(colnum, RowNum) = currentFields

                            Application.DoEvents()

                            colnum = colnum + 1
                        Next
                    Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                        If Len(messages) > 0 Then
                            messages = messages & ","
                        End If
                        messages = messages & " Line " & ex.Message & "is NOT valid and will be skipped."
                        Message = messages
                        'frmMainGIForm.logger.LogMessage("AR_Messages_v2_7.log", Application.StartupPath, Message, frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                    End Try
                    RowNum = RowNum + 1
                End While
            End Using
        Catch ex As Exception
            Message = "Error In CSVFileToArray: " & ex.Message.ToString
            'frmMainGIForm.logger.LogError("AR_Error_v2_7.log", Application.StartupPath, Message, "GetFieldValue_From_Index()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End Try
        numRowsRead = RowNum
        'currentRow is an array containing all the data now.
        'Me.rtbOutput.Text = ""
        'frmProgressBar.Hide()
        CSVFileToArray = numRowsRead
    End Function


    Function GetFieldValue_From_Fieldname(ByVal AllRows As Object, ByVal RowIDX As Long, ByVal FieldName As String, Optional ByRef FieldPos As Long = 0) As String
        Dim ReturnValue As String
        Dim colIDX As Long
        Dim ArrayValues As String

        'AllRows(Fields,0) - row 0 contains all the fieldnames
        GetFieldValue_From_Fieldname = ""
        ReturnValue = ""
        For colIDX = 0 To UBound(AllRows, 1)
            ArrayValues = AllRows(colIDX, RowIDX)
            If UCase(FieldName) = UCase(AllRows(colIDX, 0)) Then
                ReturnValue = ArrayValues
                FieldPos = colIDX
                Exit For
            End If
            'colIDX = colIDX + 1
        Next
        GetFieldValue_From_Fieldname = ReturnValue
    End Function

    Public Function ExtractExcelRows(
        ByVal FileName As String,
        ByVal SheetName As String,
        ByVal StartRow As Long,
        ByVal StartCol As Long,
        Optional ByVal EndRow As Long = 0,
        Optional ByVal EndCol As Long = 0,
        Optional ByRef dt As DataTable = Nothing) As Object(,)

        Dim Proceed As Boolean = False
        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkBooks As Excel.Workbooks = Nothing
        Dim xlWorkBook As Excel.Workbook = Nothing
        Dim xlWorkSheet As Excel.Worksheet = Nothing
        Dim xlWorkSheets As Excel.Sheets = Nothing
        Dim xlCells As Excel.Range = Nothing
        Dim ExcelArray(,) As Object
        Dim dtDeliveryDate As Date
        'Dim dt As New DataTable
        Dim dr As DataRow
        Dim tempTable As New DataTable
        Dim datacol As New DataColumn
        Dim LastRow As Long
        Dim LastCol As Long
        Dim xlUsedRange As Excel.Range
        Dim RowIDX As Integer
        Dim ColIDX As Integer
        Dim TotalFields As Integer
        Dim Percentage As Single = 0.0F

        Percentage = 0.0F
        xlUsedRange = Nothing
        frmMainGIForm.pbarMain.Style = ProgressBarStyle.Blocks
        frmMainGIForm.pbarMain.Value = 0
        ExcelArray = Nothing
        If IO.File.Exists(FileName) Then

            xlApp = New Excel.Application
            xlApp.DisplayAlerts = False
            xlWorkBooks = xlApp.Workbooks
            If GlobalWorkBook Is Nothing Then
                xlWorkBook = xlWorkBooks.Open(FileName)
            Else
                'BUT if we do get a COM error here - may need to add xlAPP as a public variable too ??? = GlobalxlApp
                'OR keep saving to a Temporary EXCEL file for EACH procedure ???? to preserve.
                xlWorkBook = GlobalWorkBook
            End If
            xlApp.Visible = False

                xlWorkSheets = xlWorkBook.Sheets

                '
                ' For/Next finds our sheet
                '
                For x As Integer = 1 To xlWorkSheets.Count
                    xlWorkSheet = CType(xlWorkSheets(x), Excel.Worksheet)

                    If xlWorkSheet.Name = SheetName Then
                        Proceed = True
                        Exit For
                    End If
                    'xltoleft not part of xlworksheet.



                Next

                If Proceed Then

                    'dt.Columns.AddRange(
                    '    New DataColumn() _
                    '    {
                    '        New DataColumn With {.ColumnName = "Identifier", .DataType = GetType(Int32), .AutoIncrement = True, .AutoIncrementSeed = 1},
                    '        New DataColumn With {.ColumnName = "SomeDate", .DataType = GetType(Date)}
                    '    }
                    ')

                    If EndCol = 0 Then
                        LastCol = xlWorkSheet.Cells(2, xlWorkSheet.Columns.Count).end(Excel.XlDirection.xlToLeft).column
                    Else
                        LastCol = EndCol
                    End If
                    'LastCol = ThisWB.Worksheets(SearchWorksheetName).Cells(2, Columns.Count).End(xlToLeft).Column
                    'LastRow = xlWorkSheet.Cells.Find(What:="*", After:= ["A1"], SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlPrevious)
                    If EndRow = 0 Then
                        LastRow = xlWorkSheet.Cells(xlWorkSheet.Rows.Count, 1).End(Excel.XlDirection.xlUp).row
                    Else
                        LastRow = EndRow
                    End If

                    xlUsedRange = xlWorkSheet.Range(xlWorkSheet.Cells(StartRow, StartCol), xlWorkSheet.Cells(LastRow, LastCol))
                    'sort sheet before extraction:

                    Try

                        ExcelArray = CType(xlUsedRange.Value(Excel.XlRangeValueDataType.xlRangeValueDefault), Object(,))

                        If ExcelArray IsNot Nothing Then
                            Dim bounds As Integer = ExcelArray.GetUpperBound(0)
                            'dtDeliveryDate = ExcelArray
                            TotalFields = UBound(ExcelArray, 2)
                            'Clear the grid:

                            For ColIDX = 1 To TotalFields
                                tempTable.Columns.Add("Column" & CStr(ColIDX))
                            Next

                            For RowIDX = 1 To bounds
                                'dr = dt.NewRow
                                'For ColIDX = 1 To 10
                                'dt.Columns.Add(ExcelArray(RowIDX, ColIDX))
                                If (ExcelArray(RowIDX, 1) IsNot Nothing) Then
                                    dr = tempTable.NewRow
                                    For ColIDX = 1 To TotalFields
                                        dr.Item("Column" & CStr(ColIDX)) = ExcelArray(RowIDX, ColIDX)

                                        ''Need to test for each field to make sure its correct type:
                                        'If VarType(ExcelArray(row, 1)) = VariantType.Date Then
                                        'dt.Columns.Add(ExcelArray(RowIDX, ColIDX).ToString, GetType(String))
                                        'dr.Item(ColIDX) = ExcelArray(RowIDX, ColIDX).ToString
                                        'Else
                                        'dt.Rows.Add(New Object() {Nothing, Nothing})
                                    Next
                                    tempTable.Rows.Add(dr)
                                End If
                                'dt.Rows.Add(New Object() {Nothing, ExcelArray(RowIDX, ColIDX).ToString})
                                'dt.Rows.Add(dr)
                                'End If
                                'Next
                                Percentage = (RowIDX / bounds) * 100
                                frmMainGIForm.pbarMain.Value = CInt(Percentage)
                                Application.DoEvents()
                            Next
                            'Dim dr As DataRow

                            'With xlWorkSheet
                            'For RowIDX = 2 To bounds - 1
                            'dr = dt.NewRow

                            'dr.Item("Column1") = .Cells(RowIDX, 1).Text
                            'dr.Item("Column2") = .Cells(RowIDX, 2).Text


                            'dt.Rows.Add(dr)
                            'Next
                            'End With
                        End If
                    Finally
                        dt = tempTable
                        releaseObject(xlUsedRange)
                    End Try
                Else
                    MessageBox.Show(SheetName & " not found.")
                End If

                xlWorkBook.Close()
                xlApp.UserControl = True
                xlApp.Quit()
                releaseObject(xlCells)
                releaseObject(xlUsedRange)
                releaseObject(xlWorkSheets)
                releaseObject(xlWorkSheet)
                releaseObject(xlWorkBook)
                releaseObject(xlWorkBooks)
                releaseObject(xlApp)

                Marshal.FinalReleaseComObject(xlWorkSheet)
                xlWorkSheet = Nothing
            Else
                MessageBox.Show("'" & FileName & "' not located. ")
        End If

        Return ExcelArray

    End Function

    Public Function ExtractExcelRange(
        ByVal FileName As String,
        ByVal SheetName As String,
        ByVal SearchText As String,
        ByVal SearchCol As Long,
        Optional SearchFieldType As String = "STRING",
        Optional tempFilename As String = "",
        Optional InclTitlesForDataTable As Boolean = False,
        Optional StartSearchCol As Long = 1,
        Optional StartSearchRow As Long = 1,
        Optional ByVal EndRow As Long = 0,
        Optional ByVal EndCol As Long = 0,
        Optional ByRef dt As DataTable = Nothing,
        Optional ByVal LastSearchDate As String = "") As Object(,)

        Dim Proceed As Boolean = False
        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkBooks As Excel.Workbooks = Nothing
        Dim xlWorkBook As Excel.Workbook = Nothing
        Dim xlWorkSheet As Excel.Worksheet = Nothing
        Dim xlWorkSheets As Excel.Sheets = Nothing
        Dim xlCells As Excel.Range = Nothing
        Dim ExcelArray(,) As Object
        Dim dtFoundDate As Date
        'Dim dt As New DataTable
        Dim dr As DataRow
        Dim tempTable As New DataTable
        Dim datacol As New DataColumn
        Dim LastRow As Long
        Dim LastCol As Long
        Dim FoundFirstRow As Long = 0
        Dim FoundLastRow As Long = 0
        Dim LastSearchDate_FirstRow As Long = 0
        Dim LastSearchDate_LastRow As Long = 0
        Dim xlUsedRange As Excel.Range
        Dim RowIDX As Integer
        Dim ColIDX As Integer
        Dim TotalFields As Integer
        Dim Percentage As Single = 0.0F
        Dim FoundText As String = ""
        Dim Titles() As String = Nothing
        Dim TheFilename As String = ""

        Percentage = 0.0F
        xlUsedRange = Nothing
        frmMainGIForm.pbarMain.Style = ProgressBarStyle.Blocks
        frmMainGIForm.pbarMain.Value = 0
        frmMainGIForm.txtMessages.Text = "Please Wait ... Get data from Spreadsheet"
        ExcelArray = Nothing
        'If Len(tempFilename) > 0 Then
        'TheFilename = tempFilename
        'Else
        TheFilename = FileName
        'End If
        If IO.File.Exists(TheFilename) Then

            xlApp = New Excel.Application
            xlApp.DisplayAlerts = False
            xlWorkBooks = xlApp.Workbooks
            'If frmMainGIForm.MainWorkBook Is Nothing Then
            xlWorkBook = xlWorkBooks.Open(TheFilename)
            xlWorkSheets = xlWorkBook.Sheets
            'Else
            'xlWorkBook = xlWorkBooks.Add(frmMainGIForm.MainWorkBook)
            'xlWorkBook = frmMainGIForm.MainWorkBook
            'xlWorkSheets = xlWorkBook.Sheets
            'End If
            xlApp.Visible = False

            ' For/Next finds our sheet
            '
            For x As Integer = 1 To xlWorkSheets.Count
                xlWorkSheet = CType(xlWorkSheets(x), Excel.Worksheet)

                If xlWorkSheet.Name = SheetName Then
                    Proceed = True
                    Exit For
                End If
                'xltoleft not part of xlworksheet.
            Next

            If Proceed Then
                If EndCol = 0 Then
                    LastCol = xlWorkSheet.Cells(2, xlWorkSheet.Columns.Count).end(Excel.XlDirection.xlToLeft).column
                Else
                    LastCol = EndCol
                End If
                If EndRow = 0 Then
                    LastRow = xlWorkSheet.Cells(xlWorkSheet.Rows.Count, 1).End(Excel.XlDirection.xlUp).row
                Else
                    LastRow = EndRow
                End If

                'NEed to go through each row in spreadsheet and try and find the searchtext to return the row it starts on.

                For RowIDX = StartSearchRow To LastRow
                    xlUsedRange = xlWorkSheet.Range(xlWorkSheet.Cells(RowIDX, SearchCol), xlWorkSheet.Cells(RowIDX, SearchCol))
                    FoundText = xlUsedRange.Text
                    If UCase(SearchFieldType) = "DATE" Then

                        If Len(LastSearchDate) > 0 Then
                            If IsDate(FoundText) And IsDate(LastSearchDate) Then

                                If CDate(FoundText) = CDate(LastSearchDate) Then
                                    If LastSearchDate_FirstRow = 0 Then
                                        LastSearchDate_FirstRow = RowIDX
                                    Else
                                        LastSearchDate_LastRow = RowIDX
                                    End If
                                End If
                            Else

                            End If 'ISDATE()
                        End If

                        If IsDate(FoundText) And IsDate(SearchText) Then

                            If CDate(FoundText) = CDate(SearchText) Then
                                If FoundFirstRow = 0 Then
                                    FoundFirstRow = RowIDX
                                Else
                                    FoundLastRow = RowIDX

                                End If
                            End If
                        Else

                        End If 'ISDATE()


                    ElseIf UCase(SearchFieldType) = "INTEGER" Then
                        If IsNumeric(FoundText) And IsNumeric(SearchText) Then

                            If CInt(FoundText) = CInt(SearchText) Then
                                If FoundFirstRow = 0 Then
                                    FoundFirstRow = RowIDX
                                Else
                                    FoundLastRow = RowIDX

                                End If
                            End If
                        End If 'IsINTEGER()
                    ElseIf UCase(SearchFieldType) = "LONG" Then
                        If IsNumeric(FoundText) And IsNumeric(SearchText) Then

                            If CLng(FoundText) = CLng(SearchText) Then
                                If FoundFirstRow = 0 Then
                                    FoundFirstRow = RowIDX
                                Else
                                    FoundLastRow = RowIDX

                                End If
                            End If
                        End If 'IsLONG()
                    Else 'Field Type = STRING
                        If UCase(FoundText) = UCase(SearchText) Then
                            If FoundFirstRow = 0 Then
                                FoundFirstRow = RowIDX
                            Else
                                FoundLastRow = RowIDX

                            End If
                        End If
                    End If
                    'If FoundLastRow > 1 Then Exit For
                    Percentage = (RowIDX / LastRow) * 100
                    frmMainGIForm.pbarMain.Value = CInt(Percentage)
                    Application.DoEvents()
                Next
                If Len(LastSearchDate) > 0 Then
                    FoundLastRow = LastSearchDate_LastRow
                End If
                If FoundFirstRow > 0 And FoundLastRow > 0 Then
                    xlUsedRange = xlWorkSheet.Range(xlWorkSheet.Cells(FoundFirstRow, StartSearchCol), xlWorkSheet.Cells(FoundLastRow, LastCol))
                    'sort sheet before extraction:

                    Try

                        ExcelArray = CType(xlUsedRange.Value(Excel.XlRangeValueDataType.xlRangeValueDefault), Object(,))
                        frmMainGIForm.txtMessages.Text = "Please Wait ... Processing Spreadsheet Data"

                    Finally
                        dt = tempTable
                        releaseObject(xlUsedRange)
                    End Try
                    'MessageBox.Show("Cannot Find Date In Sheet: " & SearchText)
                End If 'Found First Row
            Else
                MessageBox.Show(SheetName & " not found.")
            End If 'Proceed

            'Throwing up exception error here on other laptops:
            'xlWorkBook.Close()
            xlApp.UserControl = True
            xlApp.Quit()
            releaseObject(xlCells)
            releaseObject(xlUsedRange)
            releaseObject(xlWorkSheets)
            releaseObject(xlWorkSheet)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkBooks)
            releaseObject(xlApp)

            Marshal.FinalReleaseComObject(xlWorkSheet)
            xlWorkSheet = Nothing
        Else
            MessageBox.Show("'" & FileName & "' not located. ")
        End If 'If file found

        Return ExcelArray

    End Function

    Sub SortSheet(WBFileName As String, Worksheetname As String, ByVal StartRow As Long, ByVal StartCol As Long, SortColumn1 As Long, ByRef ExcelArray(,) As Object,
                    Optional ByVal CheckRow As Long = 1, Optional CheckCol As Long = 1, Optional SortColumn2 As Long = 0,
                    Optional REverseSort As Boolean = False, Optional ByVal PreserveWorkbook As Boolean = False, Optional ByVal tempFilename As String = "")
        Dim LastRow As Long
        Dim LastCol As Long
        Dim EndCol As Long
        Dim EndRow As Long
        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkBooks As Excel.Workbooks = Nothing
        Dim xlWorkBook As Excel.Workbook = Nothing
        Dim xlWorkSheet As Excel.Worksheet = Nothing
        Dim xlWorkSheets As Excel.Sheets = Nothing
        Dim xlCells As Excel.Range = Nothing
        Dim xlUsedRange As Excel.Range
        Dim xlWholeRange As Excel.Range

        If IO.File.Exists(WBFileName) Then

            xlApp = New Excel.Application
            xlApp.DisplayAlerts = False
            xlWorkBooks = xlApp.Workbooks
            xlWorkBook = xlWorkBooks.Open(WBFileName)

            xlApp.Visible = False

            xlWorkSheets = xlWorkBook.Sheets
            xlWorkSheet = xlWorkBook.Worksheets(Worksheetname)

            If EndCol = 0 Then
                LastCol = xlWorkSheet.Cells(CheckRow, xlWorkSheet.Columns.Count).end(Excel.XlDirection.xlToLeft).column
            Else
                LastCol = EndCol
            End If
            'LastCol = ThisWB.Worksheets(SearchWorksheetName).Cells(2, Columns.Count).End(xlToLeft).Column
            'LastRow = xlWorkSheet.Cells.Find(What:="*", After:= ["A1"], SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlPrevious)
            If EndRow = 0 Then
                LastRow = xlWorkSheet.Cells(xlWorkSheet.Rows.Count, CheckCol).End(Excel.XlDirection.xlUp).row
            Else
                LastRow = EndRow
            End If
            If LastRow = 0 Or LastCol = 0 Then
                MsgBox("CAnnot Find Last Row")
                xlWorkBook.Close()
                xlApp.UserControl = True
                xlApp.Quit()
                releaseObject(xlCells)
                releaseObject(xlWholeRange)
                releaseObject(xlWorkSheets)
                releaseObject(xlWorkSheet)
                releaseObject(xlWorkBook)
                releaseObject(xlWorkBooks)
                releaseObject(xlApp)
                Exit Sub
            End If
            xlWorkSheet.Sort.SortFields.Clear()

            If SortColumn1 > 0 Then
                StartCol = 1
                StartRow = 2
                'for DELIVERY DATE:
                xlUsedRange = xlWorkSheet.Range(xlWorkSheet.Cells(StartRow, SortColumn1), xlWorkSheet.Cells(LastRow, SortColumn1))
                If REverseSort = True Then
                    xlWorkSheet.Sort.SortFields.Add(Key:=xlUsedRange, SortOn:=0, Order:=2, DataOption:=0)
                Else
                    xlWorkSheet.Sort.SortFields.Add(Key:=xlUsedRange, SortOn:=0, Order:=1, DataOption:=0)
                End If
            End If

            If SortColumn2 > 0 Then
                StartCol = 1
                StartRow = 2
                'for DELIVERY REF:
                xlUsedRange = xlWorkSheet.Range(xlWorkSheet.Cells(StartRow, SortColumn2), xlWorkSheet.Cells(LastRow, SortColumn2))
                xlWorkSheet.Sort.SortFields.Add(Key:=xlUsedRange, SortOn:=0, Order:=1, DataOption:=0)
            End If
            xlWholeRange = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(LastRow, LastCol))

            With xlWorkSheet.Sort
                .SetRange(xlWholeRange)
                .Header = 1
                .MatchCase = False
                .Orientation = 1
                .SortMethod = 1
                .Apply()
            End With
        Else
            MsgBox("Filename is empty")
            'xlWorkBook.Close()
            'xlApp.UserControl = True
            'xlApp.Quit()
            releaseObject(xlCells)
            releaseObject(xlWholeRange)
            releaseObject(xlWorkSheets)
            releaseObject(xlWorkSheet)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkBooks)
            releaseObject(xlApp)
            If xlWorkSheet IsNot Nothing Then
                Marshal.FinalReleaseComObject(xlWorkSheet)
            End If
            xlWorkSheet = Nothing
                Exit Sub
            End If
            If LastCol > 0 And LastRow > 0 Then
            xlWholeRange = xlWorkSheet.Range(xlWorkSheet.Cells(StartRow, StartCol), xlWorkSheet.Cells(LastRow, LastCol))

            If PreserveWorkbook Then
                'OR IS IT THE worksheet that we need to declare public ???????????????
                frmMainGIForm.MainWorkBook = xlWorkBook
            End If

            ExcelArray = CType(xlWholeRange.Value(Excel.XlRangeValueDataType.xlRangeValueDefault), Object(,))

            If PreserveWorkbook And Len(tempFilename) > 0 Then
                xlWorkBook.SaveCopyAs(tempFilename)
            End If
        End If
        xlWorkBook.Close()
        xlApp.UserControl = True
        xlApp.Quit()
        releaseObject(xlCells)
        releaseObject(xlWholeRange)
        releaseObject(xlWorkSheets)
        releaseObject(xlWorkSheet)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkBooks)
        releaseObject(xlApp)

        Marshal.FinalReleaseComObject(xlWorkSheet)
        xlWorkSheet = Nothing

    End Sub

End Module
