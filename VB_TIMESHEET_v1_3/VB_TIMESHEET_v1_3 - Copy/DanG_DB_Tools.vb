Imports System.Data.OleDb
Imports System.Configuration

'****************************************************************************************************************************************************************
'*
'*  Program Title:                                            VB GOODS INWARDS TIMESHEET v 1.0
'*
'*  Language:                                                 VB.NET - Visual Studio 2015 - © 2017 Microsoft
'*
'*  AUTHOR:                                                   Daniel Goss
'*
'*  Copyright:                                                Copyright 2016-2018
'*
'*  Module Code:                                              Module: DanG_DB_Tools
'*
'*  External Name:                                            DanG_DB_Tools.vb
'*
'*  Module preious amended:                                   17/06/2018 17:35
'*                                                              lINKS TO frmMainGIForm
'*
'*  Module Last Amendment:                                    25/07/2018 16:25
'*      ADD new procedure: Sub Convert_OldEmpNo_NewEmpNo(ByVal CodeType As String, ByVal TheCodeNumber As Long, ByRef OldCode As String, ByRef NewCode As String, _
'*                          Optional LimitNumberIndex As Integer = 1)
'*      Converts old 6-digit employee numbers to the new 7-digit version.
'****************************************************************************************************************************************************************


Module DanG_DB_Tools

    Sub SaveMyConfigKey(ByVal myKey As String, ByVal myValue As String)
        'Dim cAppConfig As Configuration
        'Dim asSettings As AppSettingsSection

        'cAppConfig = ConfigurationManager.OpenExeConfiguration(Application.StartupPath & "\Config_File_Ex.exe")
        'asSettings = cAppConfig.AppSettings

        'asSettings.Settings.Item(myKey).Value = myValue
        'cAppConfig.Save(ConfigurationSaveMode.Modified)

    End Sub

    Sub LoadMyConfigKey(myKey As String, ByRef myValue As String)
        'Dim appSettings = ConfigurationManager.AppSettings
        Dim Result As String = ""

        'Try
        'Result = appSettings(myKey)
        If IsNothing(Result) Then
            Result = ""
        End If

        'Catch ec As ConfigurationErrorsException
        'MsgBox("config Exception Error: " & ex.Message, vbOK, "config Error - Asset Register " & frmMainGIForm.myVersion)

        'End Try

        myValue = Result
    End Sub

    Function FindFocussedControl(ByVal ctr As Control) As Control
        'ADDED 27/05/2018
        Dim container As ContainerControl = TryCast(ctr, ContainerControl)
        Do While (container IsNot Nothing)
            ctr = container.ActiveControl
            container = TryCast(ctr, ContainerControl)
        Loop
        Return ctr
    End Function

    Function RemoveExtractedFields(ByVal OldArray As String(), ByVal ExtractedArray As String(), ByVal substring As String, ByRef NewArray As String(), Optional ByVal ElementsOldStartIDX As Integer = 0,
                                   Optional ByVal ElementsNewStartIDX As Integer = 0) As String
        Dim TotalExtracted As Integer
        Dim FieldIDX As Integer
        Dim ElementsOld As Integer
        Dim ElementsNew As Integer
        Dim commaPos As Integer
        Dim tempArr As String()
        Dim NewString As String
        Dim FinalString As String
        Dim EX As Integer

        TotalExtracted = 0
        FinalString = ""
        commaPos = 1
        ElementsOld = ElementsOldStartIDX
        ElementsNew = ElementsNewStartIDX
        Try
            ReDim tempArr(1)
            For EX = 0 To UBound(ExtractedArray)
                NewString = ""
                For FieldIDX = 0 To UBound(OldArray)
                    If UCase(ExtractedArray(EX)) = UCase(OldArray(FieldIDX)) Then
                        'no operation needed
                    Else
                        If Len(NewString) = 0 Then
                            NewString = OldArray(FieldIDX)
                        Else
                            NewString = NewString & substring & OldArray(FieldIDX)
                        End If
                    End If
                Next
                OldArray = strToStringArray(NewString, ",")

            Next
            NewArray = OldArray
            FinalString = StringArrayToString(NewArray, ",")
        Catch exc As Exception
            MsgBox("Exception Error in RemoveExtractedFields: " & exc.ToString, vbOK, "Exception Error in Asset Register")
        End Try

        Return FinalString
    End Function

    Function GetFields(DBName As String, DBTable As String) As String
        Dim con As OleDbConnection
        Dim cmd As OleDbCommand
        Dim dt As DataTable
        Dim DataSource As String
        Dim connString As String
        Dim strSQL As String
        Dim NumRows As Integer
        Dim NumFields As Integer
        Dim NumBlankFields As Integer
        Dim Provider As String = ""
        Dim colIDX As Integer = 0
        Dim Fieldnames As String = ""
        Dim dr1 As OleDbDataReader

        GetFields = ""
        NumBlankFields = 0
        NumFields = 0
        NumRows = 0
        strSQL = ""
        DataSource = ""
        connString = ""
        If Len(DBTable) = 0 Then
            MsgBox("Error in Insert Record: Database Table not provided", vbOK, "Error in Asset Register " & frmMainGIForm.myVersion)
            Exit Function
        End If
        If Len(DBName) = 0 Then
            MsgBox("Error in Insert Record: No Database Name specified", vbOK, "Error in Asset Register " & frmMainGIForm.myVersion)
            Exit Function
        Else
            If Len(Provider) = 0 Then
                If InStr(DBName, ".MDB") > 0 Then
                    Provider = "Provider=Microsoft.JET.4.0;"
                ElseIf InStr(DBName, ".accdb") Then
                    Provider = "Provider=Microsoft.ACE.OLEDB.12.0;"
                Else
                    MsgBox("Error in Insert Record: Unrecognised Database type: " & DBName, vbOK, "Error in Asset Register " & frmMainGIForm.myVersion)
                    Exit Function
                End If
            End If
        End If
        If Not Right(DBName, 1) = ";" Then
            DBName = DBName & ";"
        End If
        DataSource = frmMainGIForm.dbsource & DBName
        If Not Right(Provider, 1) = ";" Then
            Provider = Provider & ";"
        End If
        connString = Provider & DataSource
        con = New OleDbConnection(connString)

        con.Open()
        strSQL = "SELECT * FROM " & DBTable
        cmd = New OleDbCommand(strSQL, con)
        dr1 = cmd.ExecuteReader()

        NumFields = dr1.FieldCount - 1

        While colIDX <= NumFields
            If colIDX = 0 Then
                Fieldnames = dr1.GetName(colIDX)
            Else
                Fieldnames = Fieldnames & "," & dr1.GetName(colIDX)
            End If
            colIDX = colIDX + 1
        End While
        con.Close()
        dr1.Close()
        GetFields = Fieldnames

    End Function

    Function CheckBlankValues(ByVal FieldNames As String, ByRef FieldValues As String, ByRef ExcludeFields As String) As Integer
        Dim fldIDX As Integer
        Dim ExcludeMe As String
        Dim ValueArray As String()
        Dim FieldNameArray As String()
        Dim TotalFields As Integer
        Dim TotalExcluded As Integer
        Dim NewFieldValues As String
        Dim ValueIDX As Integer

        ExcludeMe = ""
        TotalExcluded = 0
        ValueIDX = 0
        NewFieldValues = ""
        ReDim ValueArray(1)
        ReDim FieldNameArray(1)
        Try
            FieldNameArray = strToStringArray(FieldNames, ",")
            ValueArray = strToStringArray(FieldValues, ",", 1) ' START at ELEMENT 1 to force a NOTHING into ELEMENT 0 to match the FieldNAme array.
            TotalFields = UBound(FieldNameArray) 'FieldNameArray could have an extra element - ID as element 0  - this will cause an INDEX overflow in the VALUES array.
            For fldIDX = 0 To TotalFields - 1
                If Len(ValueArray(fldIDX)) = 0 Or IsNothing(ValueArray(fldIDX)) Then
                    If Not IsNothing(FieldNameArray(fldIDX)) Then
                        If fldIDX = 0 Then
                            ExcludeMe = FieldNameArray(fldIDX)
                        Else
                            ExcludeMe = ExcludeMe & "," & FieldNameArray(fldIDX)
                        End If
                        'Try and exclude the corresponding FieldName but only works if either "ID" is excluded from the Fieldnames OR

                        TotalExcluded = TotalExcluded + 1 'ValueArray includes the extra field NULL / Nothing as the first element 0.
                    End If
                Else
                    If ValueIDX = 0 Then
                        NewFieldValues = FieldValues
                    Else
                        NewFieldValues = NewFieldValues & "," & FieldValues
                    End If
                    ValueIDX = ValueIDX + 1
                End If

            Next
            ExcludeFields = ExcludeMe
            FieldValues = NewFieldValues
        Catch ex As Exception
            MsgBox("Exception Error in CheckBlankValues: " & ex.ToString, vbOK, "Exception Error in Asset Register " & frmMainGIForm.myVersion)
        End Try

        Return TotalExcluded

    End Function

    Function InsertUpdateRecord(UPDATE As Boolean, DBName As String, ByRef Provider As String, DBTable As String, FieldNames As String, FieldValues As String,
                                    Optional ByVal Criteria As String = "", Optional ByVal ExcludeFields As String = "") As Boolean

        Dim con As OleDbConnection
        Dim cmd As OleDbCommand
        Dim da As OleDbDataAdapter
        Dim ds As DataSet
        Dim DataSource As String
        Dim connString As String
        Dim strSQL As String
        Dim ExcludedFields As String
        Dim NumRows As Integer
        Dim NumFields As Integer
        Dim NumBlankFields As Integer

        'Dim sql = "INSERT INTO Contacts (FName, LName, Age, " &
        '  "[Address Line 1], [Address Line 2], City, State, Zip, " &
        ' "[Home Phone], [Work Phone], Email, Sex) " &
        '"VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

        NumBlankFields = 0
        NumFields = 0
        NumRows = 0
        strSQL = ""
        DataSource = ""
        connString = ""
        ExcludedFields = ""
        InsertUpdateRecord = False
        If Len(DBTable) = 0 Then
            MsgBox("Error in Insert Record: Database Table not provided", vbOK, "Error in Asset Register " & frmMainGIForm.myVersion)
            InsertUpdateRecord = False
            Exit Function
        End If
        If Len(DBName) = 0 Then
            MsgBox("Error in Insert Record: No Database Name specified", vbOK, "Error in Asset Register " & frmMainGIForm.myVersion)
            InsertUpdateRecord = False
            Exit Function
        Else
            If Len(Provider) = 0 Then
                If InStr(DBName, ".MDB") > 0 Then
                    Provider = "Provider=Microsoft.JET.4.0;"
                ElseIf InStr(DBName, ".accdb") Then
                    Provider = "Provider=Microsoft.ACE.OLEDB.12.0;"
                Else
                    MsgBox("Error in Insert Record: Unrecognised Database type: " & DBName, vbOK, "Error in Asset Register " & frmMainGIForm.myVersion)
                    InsertUpdateRecord = False
                    Exit Function
                End If
            End If
        End If
        If Not Right(DBName, 1) = ";" Then
            DBName = DBName & ";"
        End If
        DataSource = frmMainGIForm.dbsource & DBName
        If Not Right(Provider, 1) = ";" Then
            Provider = Provider & ";"
        End If
        connString = Provider & DataSource

        If Len(FieldNames) = 0 Then
            FieldNames = GetFields(DBName, DBTable)
        End If
        'NumBlankFields = CheckBlankValues(FieldNames, FieldValues, ExcludedFields) 'any blank values passed in fieldvalues will put corresponding fieldnames into ExcludedFields output.
        If Len(ExcludeFields) > 0 Then
            If Len(ExcludedFields) = 0 Then
                ExcludedFields = ExcludeFields
            Else
                ExcludedFields = ExcludedFields & "," & ExcludeFields
            End If
        End If
        If UPDATE = True Then
            strSQL = PrepareUpdate(DBName, DBTable, FieldNames, FieldValues, Criteria, ExcludedFields, connString)
        Else
            strSQL = PrepareInsert(DBName, DBTable, FieldNames, FieldValues, ExcludedFields, connString)
        End If
        If Len(strSQL) = 0 Then
            MsgBox("Error in InsertUpdateRecord : query is blank.", vbOK, "Error in Asset Register")
            InsertUpdateRecord = False
            Exit Function
        End If
        Try
            con = New OleDbConnection(connString)
            cmd = New OleDbCommand(strSQL, con)

            'Using con
            'Using cmd
            con.Open()
            'cmd.Parameters.AddWithValue("@p1", "Value For FName")
            'cmd.Parameters.AddWithValue("@p2", "Value For LName")
            'cmd.Parameters.AddWithValue("@p3", Convert.ToInt32("Value For Age"))
            '....and so On For the other parameters .....
            cmd.ExecuteNonQuery()
            'End Using
            'End Using
            con.Close()


        Catch ex As Exception
            MsgBox("Exception Error in InsertUpdateRecord Record: " & ex.ToString, vbOK, "Exception Error in AssetRegister v" & frmMainGIForm.myVersion)
        End Try
        InsertUpdateRecord = True 'needs a test to make sure the record HAS been inserted. Use the OleDbDataReader.RecordsAffected property

    End Function

    Function DeleteRecord(DBName As String, ByRef Provider As String, DBTable As String, ByVal Criteria As String) As Boolean
        Dim con As OleDbConnection
        Dim cmd As OleDbCommand
        Dim da As OleDbDataAdapter
        Dim ds As DataSet
        Dim DataSource As String
        Dim connString As String
        Dim strSQL As String
        Dim ExcludedFields As String
        Dim NumRows As Integer
        Dim NumFields As Integer
        Dim NumBlankFields As Integer
        Dim Value As String
        Dim EqualPos As Integer

        NumBlankFields = 0
        NumFields = 0
        NumRows = 0
        strSQL = ""
        DataSource = ""
        connString = ""
        ExcludedFields = ""
        DeleteRecord = False
        If Len(DBTable) = 0 Then
            MsgBox("Error in Delete Record: Database Table not provided", vbOK, "Error in Asset Register " & frmMainGIForm.myVersion)
            DeleteRecord = False
            Exit Function
        End If
        If Len(DBName) = 0 Then
            MsgBox("Error in Delete Record: No Database Name specified", vbOK, "Error in Asset Register " & frmMainGIForm.myVersion)
            DeleteRecord = False
            Exit Function
        Else
            If Len(Provider) = 0 Then
                If InStr(DBName, ".MDB") > 0 Then
                    Provider = "Provider=Microsoft.JET.4.0;"
                ElseIf InStr(DBName, ".accdb") Then
                    Provider = "Provider=Microsoft.ACE.OLEDB.12.0;"
                Else
                    MsgBox("Error in Delete Record: Unrecognised Database type: " & DBName, vbOK, "Error in Asset Register " & frmMainGIForm.myVersion)
                    DeleteRecord = False
                    Exit Function
                End If
            End If
        End If
        If Not Right(DBName, 1) = ";" Then
            DBName = DBName & ";"
        End If
        DataSource = frmMainGIForm.dbsource & DBName
        If Not Right(Provider, 1) = ";" Then
            Provider = Provider & ";"
        End If
        connString = Provider & DataSource
        strSQL = "DELETE FROM " & DBTable
        EqualPos = InStr(Criteria, "=")
        If Len(Criteria) > 0 And EqualPos > 0 Then
            If EqualPos < Len(Criteria) Then
                strSQL = strSQL & " WHERE " & Criteria
            End If
        End If
        Try
            con = New OleDbConnection(connString)
            cmd = New OleDbCommand(strSQL, con)

            'Using con
            'Using cmd
            con.Open()
            'cmd.Parameters.AddWithValue("@p1", "Value For FName")
            'cmd.Parameters.AddWithValue("@p2", "Value For LName")
            'cmd.Parameters.AddWithValue("@p3", Convert.ToInt32("Value For Age"))
            '....and so On For the other parameters .....
            cmd.ExecuteNonQuery()
            'End Using
            'End Using
            con.Close()


        Catch ex As Exception
            MsgBox("Exception Error in InsertUpdateRecord Record: " & ex.ToString, vbOK, "Exception Error in AssetRegister v" & frmMainGIForm.myVersion)
        End Try
        DeleteRecord = True 'needs a test to make sure the record HAS been Deleted. Use the OleDbDataReader.RecordsAffected property

    End Function

    Function PrepareInsert(ByVal DBName As String, ByVal DBTable As String, ByRef Fieldnames As String, ByVal Fieldvalues As String, Optional ByRef ExcludeFields As String = "",
                           Optional connString As String = "") As String
        Dim FieldNameArray As String()
        Dim IgnoreFieldsArray As String()
        Dim ValueArray As String()
        Dim FinalCMD As String
        Dim IDX As Integer
        Dim fldName As String
        Dim NumFields As Integer
        Dim NewFieldsArray As String()

        FinalCMD = ""
        ReDim FieldNameArray(1)
        ReDim NewFieldsArray(1)
        ReDim IgnoreFieldsArray(1)
        Try
            If Len(ExcludeFields) > 0 Then
                IgnoreFieldsArray = strToStringArray(ExcludeFields, ",")
            End If
            If Len(DBTable) = 0 Then
                MsgBox("Error in PrepareInsert: No Database Table specified", vbOK, "Error in Asset Register")
                PrepareInsert = ""
                Exit Function
            End If
            If Len(Fieldnames) = 0 Then
                NumFields = GetNumFields(connString, "SELECT * FROM " & DBTable, DBName, Fieldnames)
            End If
            FieldNameArray = strToStringArray(Fieldnames, ",")
            If Len(Fieldvalues) > 0 Then
                ValueArray = strToStringArray(Fieldvalues, ",")
            Else
                MsgBox("Error in PrepareInsert: No values specified", vbOK, "Error in Asset Register")
                PrepareInsert = ""
                Exit Function
            End If
            Fieldnames = RemoveExtractedFields(FieldNameArray, IgnoreFieldsArray, ",", FieldNameArray)
            If UBound(FieldNameArray) > 0 And UBound(ValueArray) > 0 Then

                If UBound(FieldNameArray) < UBound(ValueArray) Then
                    'MsgBox("Error in PrepareInsert: Number of Fields passed are LESS THAN Number of Values Passed.", vbOK, "MISS-MATCH Error in Asset Register")
                    'PrepareInsert = ""
                    'Exit Function
                End If
                If UBound(FieldNameArray) > UBound(ValueArray) Then
                    'MsgBox("Error in PrepareInsert: Number of Fields passed are GREATER THAN Number of Values Passed.", vbOK, "MISS-MATCH Error in Asset Register")
                    'PrepareInsert = ""
                    'Exit Function
                End If
            End If
            IDX = 0
            FinalCMD = "INSERT INTO " & DBTable & " (" & Fieldnames & ") VALUES (" & Fieldvalues & ")"
            'MsgBox("FinalCMD = " & FinalCMD)
        Catch ex As Exception
            MsgBox("Error Exception in PrepareInsert: " & ex.ToString, vbOK, "Exception Error in Asset Register")
        End Try
        PrepareInsert = FinalCMD


    End Function

    Function PrepareUpdate(ByVal DBName As String, ByVal TableName As String,
                           ByRef Fieldnames As String, ByVal Fieldvalues As String, Optional ByVal criteria As String = "",
                           Optional ByRef ExcludeFields As String = "", Optional connString As String = "") As String
        Dim FieldNameArray As String()
        Dim IgnoreFieldsArray As String()
        Dim ValueArray As String()
        Dim FinalCMD As String
        Dim IDX As Integer
        Dim fldName As String
        Dim fldValue As String
        Dim NumFields As Integer
        Dim UpdateCmd As String = "UPDATE"

        FinalCMD = ""
        ReDim FieldNameArray(1)
        ReDim IgnoreFieldsArray(1)
        ReDim ValueArray(1)
        Try
            'If Len(IgnoreFields) > 0 Then
            'IgnoreFieldsArray = strToStringArray(IgnoreFields, ",")
            'End If
            If Len(TableName) = 0 Then
                MsgBox("Error in PrepareUpdate: No Database Table specified", vbOK, "Error in Asset Register")
                PrepareUpdate = ""
                Exit Function
            End If
            If Len(Fieldnames) = 0 Then
                'NumFields = GetNumFields(connString, "SELECT * FROM " & TableName, DBName, Fieldnames)
                Fieldnames = GetFields(DBName, TableName)
            End If
            FieldNameArray = strToStringArray(Fieldnames, ",")
            If Len(Fieldvalues) > 0 Then
                ValueArray = strToStringArray(Fieldvalues, ",")
            Else
                MsgBox("Error in PrepareUpdate: No values specified", vbOK, "Error in Asset Register " & frmMainGIForm.myVersion)
                PrepareUpdate = ""
                Exit Function
            End If
            Fieldnames = RemoveExtractedFields(FieldNameArray, IgnoreFieldsArray, ",", FieldNameArray) 'rebuilds whole list without the extracted fields
            If UBound(FieldNameArray) > 0 And UBound(ValueArray) > 0 Then
                If UBound(FieldNameArray) < UBound(ValueArray) Then
                    MsgBox("Error in PrepareUpdate: Number of Fields Passed are LESS than Number of VALUES passed.", vbOK, "Miss-Match Error in Asset Register " & frmMainGIForm.myVersion)
                    PrepareUpdate = ""
                    Exit Function
                End If
                If UBound(FieldNameArray) > UBound(ValueArray) Then
                    MsgBox("Error in PrepareUpdate: Number of Fields Passed are GREATER than Number of VALUES passed.", vbOK, "Miss-Match Error in Asset Register " & frmMainGIForm.myVersion)
                    PrepareUpdate = ""
                    Exit Function
                End If
            End If
            IDX = 0
            UpdateCmd = "UPDATE " & TableName & " SET "
            For IDX = 0 To UBound(FieldNameArray)
                fldName = FieldNameArray(IDX)
                fldValue = ValueArray(IDX)
                If Len(fldValue) > 0 Then
                    If IDX = 0 Then
                        FinalCMD = fldName & " = " & Chr(34) & fldValue & Chr(34)
                    Else
                        FinalCMD = FinalCMD & "," & fldName & " = " & Chr(34) & fldValue & Chr(34)
                    End If
                Else
                    'No value
                End If
            Next
            FinalCMD = UpdateCmd & FinalCMD
            If Len(criteria) > 0 Then
                FinalCMD = FinalCMD & " WHERE " & criteria
            End If

        Catch ex As Exception
            MsgBox("Error Exception in PrepareUpdate: " & ex.ToString, vbOK, "Exception Error in Asset Register")
        End Try
        PrepareUpdate = FinalCMD


    End Function

    'Function strToStringArray(theString As String, SubString As String, Optional ByVal ElementStartIDX As Integer = 0) As String()
    'Dim tempArr As String()
    'Dim IDX As Integer
    'Dim Elements As Integer
    'Dim commaPos As Integer
    'Dim Extract As String
    '
    '   commaPos = 1
    '  IDX = 0
    ' Elements = ElementStartIDX
    'ReDim tempArr(1)
    'Try
    '
    'Do Until commaPos = 0
    '           commaPos = InStr(IDX + 1, theString, SubString)
    'If commaPos > 0 Then
    '               Extract = Mid(theString, IDX + 1, (commaPos - (IDX + 1)))
    'Else
    '               Extract = Mid(theString, IDX + 1, Len(theString))
    'End If
    'ReDim Preserve tempArr(UBound(tempArr) + 1)
    '           tempArr(Elements) = Extract
    '          IDX = commaPos
    '         Elements = Elements + 1
    'Loop
    ''Array.Copy(tempArr, strToStringArray, UBound(tempArr))
    'Catch ex As Exception
    '       MsgBox("Error Exception in strToStringArray: " & ex.ToString, vbOK, "Exception Error in Asset Register")
    'End Try

    'Return tempArr

    'End Function

    'Function StringArrayToString(ByVal theArray As String(), SubString As String, Optional ByVal ElementStartIDX As Integer = 0) As String
    'Dim tempArr As String()
    'Dim IDX As Integer
    'Dim Elements As Integer
    'Dim FinalString As String

    '   IDX = 0
    '  Elements = 0
    ' FinalString = ""
    'ReDim tempArr(1)
    'Try
    'For Elements = 0 To UBound(theArray)
    'If Len(theArray(Elements)) = 0 Or theArray(Elements) = Nothing Then
    'dont copy into string
    'Else
    'If Elements = 0 Then
    '                   FinalString = theArray(Elements)
    'Else
    '                   FinalString = FinalString & SubString & theArray(Elements)
    'End If
    'End If

    'Next
    'Array.Copy(tempArr, strToStringArray, UBound(tempArr))
    'Catch ex As Exception
    '       MsgBox("Error Exception in strToStringArray: " & ex.ToString, vbOK, "Exception Error in Asset Register " & frmMainGIForm.myVersion)
    'End Try

    'Return FinalString

    'End Function

    Function ConvertBadChars(Entry As String, REplaceWith As String, Optional IncludeComma As Boolean = False, Optional IncludeSpeechMarks As Boolean = False,
                                Optional IncludeSPACE As Boolean = False, Optional IncludeSingleQuote As Boolean = False) As String
        Dim GoodEntry As String = ""
        Dim UglyEntry As Object
        Dim BadCharArr As Object
        Dim IDX As Integer
        Dim BadCharList As Object

        ConvertBadChars = ""
        UglyEntry = Entry
        BadCharList = "["
        BadCharList = BadCharList & "," & "]"
        BadCharList = BadCharList & "," & "/"
        BadCharList = BadCharList & "," & "*"
        BadCharList = BadCharList & "," & "\"
        BadCharList = BadCharList & "," & "?"

        If IncludeComma Then
            BadCharList = BadCharList & "," & ","
        End If
        If IncludeSpeechMarks Then
            BadCharList = BadCharList & "," & Chr(34)
        End If
        If IncludeSPACE Then
            BadCharList = BadCharList & "," & Chr(32)
        End If
        If IncludeSingleQuote Then
            BadCharList = BadCharList & "," & Chr(39)
        End If
        UglyEntry = Entry
        'BadChars = Array(BadCharList)
        BadCharArr = Split(BadCharList, ",")
        If Len(UglyEntry) > 0 Then
            For IDX = LBound(BadCharArr) To UBound(BadCharArr)
                GoodEntry = Replace(UglyEntry, BadCharArr(IDX), REplaceWith, 1)
            Next
        End If
        ConvertBadChars = GoodEntry

    End Function

    Function strToStringArray(TheString As String, substring As String, Optional ByVal ElementStartIDX As Integer = 0,
    Optional Encase_Fields As Boolean = False, Optional RemoveBadChars As Boolean = False, Optional IncludeCommaInBadChars As Boolean = False,
            Optional REplaceWith As String = "_", Optional IncludeSpeechMarksInBadChars As Boolean = False,
            Optional NormalDelimQuoteCode As Integer = 34,
            Optional DateDelimCode As Integer = 39) As String()
        Dim tempArr() As String
        Dim IDX As Integer
        Dim Elements As Integer
        Dim commaPos As Integer
        Dim Extract As String
        Dim dtDate As DateTime

        commaPos = 1
        IDX = 0
        Elements = ElementStartIDX
        ReDim tempArr(1)

        Do Until commaPos = 0
            commaPos = InStr(IDX + 1, TheString, substring)
            If commaPos > 0 Then
                Extract = Mid(TheString, IDX + 1, (commaPos - (IDX + 1))) 'from delim to next delim
            Else
                Extract = Mid(TheString, IDX + 1, Len(TheString)) 'From 1 to end of string
            End If
            ReDim Preserve tempArr(UBound(tempArr) + 1)
            If RemoveBadChars Then
                Extract = ConvertBadChars(Extract, REplaceWith, IncludeCommaInBadChars, IncludeSpeechMarksInBadChars)
            End If
            If Encase_Fields Then
                If InStr(Extract, "-") > 0 And IsDate(Extract) Then
                    'If IsDate(Extract) Then
                    'Its a Date - dont encase ? 
                    'dtDate = CDate(Extract)
                    'Extract = dtDate.ToString("yyyy-MM-dd HH:mm:ss")
                    Extract = Chr(DateDelimCode) & Extract & Chr(DateDelimCode)
                Else
                    Extract = Chr(NormalDelimQuoteCode) & Extract & Chr(NormalDelimQuoteCode)
                End If
            End If
            tempArr(Elements) = Extract
            IDX = commaPos
            Elements = Elements + 1
        Loop
        'Array.Copy(tempArr, strToStringArray, UBound(tempArr))
        strToStringArray = tempArr

    End Function


    Function StringArrayToString(ByRef theArray() As String, substring As String, Optional ByVal ElementStartIDX As Integer = 0,
    Optional RemoveBadChars As Boolean = False, Optional IncludeCommaInBadChars As Boolean = False, Optional REplaceWith As String = "_") As String
        Dim tempArr() As String
        Dim IDX As Integer
        Dim Elements As Integer
        Dim FinalString As String
        Dim RawEntry As String

        IDX = 0
        Elements = ElementStartIDX
        FinalString = ""
        StringArrayToString = ""
        ReDim tempArr(1)
        For Elements = ElementStartIDX To UBound(theArray)
            If Len(theArray(Elements)) = 0 Then
                'dont copy into string
            Else
                RawEntry = theArray(Elements)
                If RemoveBadChars Then
                    RawEntry = ConvertBadChars(RawEntry, REplaceWith, IncludeCommaInBadChars)
                End If
                If Elements = ElementStartIDX Then

                    FinalString = RawEntry
                Else
                    FinalString = FinalString & substring & RawEntry
                End If
            End If
        Next

        'Array.Copy(tempArr, strToStringArray, UBound(tempArr))
        StringArrayToString = FinalString

    End Function

    Function GetNumFields(connStr As String, Sqlstr As String, ByVal DBase As String,
                          Optional ByRef FieldList As String = "", Optional ByRef FieldTypes As String = "",
                          Optional ByRef ColumnWidths As String = "") As Integer
        Dim Numfields As Integer
        Dim con As OleDbConnection
        Dim cmd As OleDbCommand
        Dim myReader As OleDbDataReader
        Dim Provider As String
        Dim DataSource As String
        Dim ConnString As String
        Dim dt As DataTable
        Dim fldIDX As Integer
        Dim myRowData As DataRow
        Dim myColData As DataColumn

        GetNumFields = 0
        FieldList = ""
        FieldTypes = ""
        Numfields = 0
        If Len(connStr) = 0 Then
            If Len(DBase) = 0 Then
                MsgBox("Error in GetNumFields: PLEASE SUPPLY Database Name")
                GetNumFields = 0
                Exit Function
            Else
                If InStr(DBase, ".mdb") > 0 Then
                    Provider = "Provider=Microsoft.Jet.OLEDB.4.0;"
                ElseIf InStr(DBase, ".accdb") > 0 Then
                    Provider = "Provider=Microsoft.ACE.OLEDB.12.0;"
                Else
                    MsgBox("Error in GetNumFields: Unrecognised Database type: " & DBase, vbOK, "Error in Asset Register")
                    GetNumFields = 0
                    Exit Function
                End If
            End If
            If Not Right(DBase, 1) = ";" Then
                DBase = DBase & ";"
            End If
            DataSource = frmMainGIForm.dbsource & DBase
            ConnString = Provider & DataSource
        Else
            ConnString = connStr
        End If
        If Len(Sqlstr) = 0 Then
            MsgBox("Error in GetNumFields: query string passed is empty.", vbOK, "Error in Asset Register")
            GetNumFields = 0
            Exit Function
        End If
        Try
            'cmd = New OleDbCommand(strSQL, con)
            con = New OleDbConnection(ConnString)
            'cmd = con.CreateCommand()
            'cmd.CommandText = Sqlstr
            con.Open() 'ERROR PRONE HERE !!!
            cmd = New OleDbCommand(Sqlstr, con)
            myReader = cmd.ExecuteReader
            If myReader.HasRows Then
                Numfields = myReader.FieldCount
                'DO WE NEED an IsDBNull() here ?
                For fldIDX = 0 To Numfields - 1
                    If fldIDX = 0 Then
                        If Not IsDBNull(myReader.GetName(fldIDX)) Then
                            FieldList = myReader.GetName(fldIDX)
                            FieldTypes = myReader.GetName(fldIDX).GetType.ToString
                        Else
                            MsgBox("empty value, index= : " & CStr(fldIDX))
                        End If
                    Else
                        If Not IsDBNull(myReader.GetName(fldIDX)) Then
                            FieldList = FieldList & "," & myReader.GetName(fldIDX)
                            FieldTypes = FieldTypes & "," & myReader.GetName(fldIDX).GetType.ToString
                            ColumnWidths = ColumnWidths & "," & CStr(myReader.GetSchemaTable().Columns(fldIDX).MaxLength)
                        Else
                            MsgBox("empty value, index= :" & CStr(fldIDX))
                        End If
                    End If
                Next
            End If
            'dt = myReader.GetSchemaTable()
            myReader.Close()
            con.Close()

        Catch ex As Exception
            MsgBox("Error In GetnumFields: " & ex.Message.ToString)
        End Try
        GetNumFields = Numfields
    End Function

    Function SQLToArray(ByVal DBASE As String, ByVal Provider As String, ByVal SQLQuery As String, ByRef Messages As String, Optional ReturnLastID As Boolean = False, Optional DBName As String = "") As Object
        Dim NumRows As Integer
        Dim NumFields As Integer
        Dim RowNumber As Long
        Dim FieldNum As Integer
        Dim Arr(,) As String = {{0}, {"NONE"}}
        Dim con As OleDbConnection
        Dim NewValue As Long
        Dim ConnString As String
        Dim da As OleDbDataAdapter
        Dim ds As DataSet
        Dim cmd As OleDbCommand
        Dim DataSource As String
        Dim strSQL As String

        If Len(DBASE) = 0 Then
            'MsgBox("Error in SQLToArray: Database name not provided", vbOK, "Error in Asset Register")
            Messages = "Error in SQLToArray: Database name not provided"
            SQLToArray = Nothing
            Exit Function
        Else
            If InStr(DBASE, ".mdb") > 0 Then
                Provider = "Provider=Microsoft.Jet.OLEDB.4.0;"
            ElseIf InStr(DBASE, ".accdb") > 0 Then
                Provider = "Provider=Microsoft.ACE.OLEDB.12.0;"
            Else
                'MsgBox("Error in SQLToArray: Unrecognised Database type: " & DBASE, vbOK, "Error in Asset Register v" & frmMainGIForm.myVersion)
                Messages = "Error in SQLToArray: Unrecognised Database type: " & DBASE
                Exit Function
            End If

        End If
        If Not Right(DBASE, 1) = ";" Then
            DBASE = DBASE & ";"
        End If
        DataSource = frmMainGIForm.dbsource & DBASE
        If Len(Provider) = 0 Then
            'MsgBox("Error in SQLToArray: Provider is blank", vbOK, "Error in Asset Register v" & frmMainGIForm.myVersion)
            Messages = "Error in SQLToArray: Provider is blank"
            SQLToArray = Nothing
            Exit Function
        End If
        If Not Right(Provider, 1) = ";" Then
            Provider = Provider & ";"
        End If
        ConnString = Provider & DataSource
        strSQL = SQLQuery
        If Len(SQLQuery) = 0 Then

            If ReturnLastID Then

                If Len(DBName) = 0 Then
                    Messages = "Error in SQLToArray(): DBName is blank."
                    Exit Function
                End If
                strSQL = "SELECT @@IDENTITY FROM " & DBName
            Else
                'MsgBox("Error in SQLToArray: QUERY (strSQL) is blank.", vbOK, "Error in Asset Register v" & frmMainGIForm.myVersion)
                Messages = "Error in SQLToArray: QUERY (strSQL) is blank."
                SQLToArray = Nothing
                Exit Function
            End If
        End If
        Try
            con = New OleDbConnection(ConnString)
            cmd = New OleDbCommand(strSQL, con)
            con.Open()
            da = New OleDbDataAdapter(cmd)
            ds = New DataSet()
            da.Fill(ds, "srcTable")
            If IsDBNull(ds.Tables("srcTable")) Then
                'MsgBox("Problem: Empty Fields", vbOK, "Problem in Asset Register v" & frmMainGIForm.myVersion)
                Messages = "Problem: Empty Fields"
            Else
                NumRows = ds.Tables("srcTable").Rows.Count
            End If

            '******************************* untested for extra IDENTITY code inserted to return last Id

            NumFields = GetNumFields(ConnString, strSQL, "AssetRegister.accdb")
            ReDim Arr(NumFields, NumRows)
            For RowNumber = 0 To NumRows - 1
                NewValue = UBound(Arr, 2)
                For FieldNum = 0 To NumFields - 1
                    If Not IsDBNull(ds.Tables("srcTable").Rows(RowNumber).Item(FieldNum)) Then
                        Arr(FieldNum, RowNumber) = ds.Tables("srcTable").Rows(RowNumber).Item(FieldNum)
                    Else
                        Arr(FieldNum, RowNumber) = ""
                    End If
                Next
            Next
            con.Close()


        Catch ex As Exception
            'MsgBox("Exception Error in SQLToArray: " & ex.ToString, vbOK, "Exception Error in Asset Register v" & frmMainGIForm.myVersion)
            If InStr(ex.ToString, "already opened") > 0 Then
                Messages = "EXCEPTION ERROR - database already opened in another application - please close."
            Else
                Messages = "EXCEPTION ERROR - in SQLToArray() " & ex.ToString
            End If
        End Try

        SQLToArray = Arr

    End Function

    Sub PopulateDataSource(ByRef DSource As Object, ByRef DBase As String, ByRef Provider As String, ByRef SQLStr As String, ByRef Messages As String, Optional ByRef NumRows As Integer = 0)
        Dim con As OleDbConnection
        Dim cmd As OleDbCommand
        Dim da As OleDbDataAdapter
        Dim ds As DataSet
        Dim connStr As String
        Dim DataSource As String
        Dim RowIDX As Integer

        'WHERE IS COLUMN1 COMMING FROM IN TBLREGISTER ?????????????????????????
        If Len(DBase) = 0 Then
            DBase = frmMainGIForm.DBName
        Else
            If InStr(DBase, ".mdb") > 0 Then
                Provider = "Provider=Microsoft.Jet.OLEDB.4.0;"
            ElseIf InStr(DBase, ".accdb") > 0 Then
                Provider = "Provider=Microsoft.ACE.OLEDB.12.0;"
            Else
                'MsgBox("Error in PopulateDataGrid: Unrecognised Database type: " & DBase, vbOK, "Error in Asset Register")
                Messages = "Error in PopulateDataGrid: Unrecognised Database type: " & DBase
                Exit Sub
            End If
        End If
        If Not Right(DBase, 1) = ";" Then
            DBase = DBase & ";"
        End If
        If Len(Provider) = 0 Then
            'MsgBox("Error in PopulateDataGrid: Provider is empty: ", vbOK, "Error in Asset Register")
            Messages = "Error in PopulateDataGrid: Provider is empty."
            Exit Sub
        End If
        If Not Right(Provider, 1) = ";" Then
            Provider = Provider & ";"
        End If
        Try
            DataSource = frmMainGIForm.dbsource & DBase
            connStr = Provider & DataSource
            con = New OleDbConnection(connStr)
            con.Open()
            cmd = New OleDbCommand(SQLStr, con)

            da = New OleDbDataAdapter(cmd)
            ds = New DataSet()
            da.Fill(ds, "srcTable")
            If IsDBNull(ds.Tables("srcTable")) Then
                'MsgBox("Problem: Empty Fields", vbOK, "Problem in Asset Register v" & frmMainGIForm.myVersion)
                Messages = "Problem: Empty Fields"
                NumRows = 0
            Else
                DSource = ds.Tables("srcTable") 'Take all fields and data
                NumRows = ds.Tables("srcTable").Rows.Count
                'ds.Tables("srcTable").Columns.Add.Caption = "Row Num"
                For RowIDX = 1 To NumRows
                    'ds.Tables("srcTable").item(rowidx).value = CStr(RowIDX)
                Next
            End If

            con.Close()
        Catch ex As Exception
            'MsgBox("Exception Error in PopulateDataSource: " & ex.ToString, vbOK, "Exception Error in Asset Register v" & frmMainGIForm.myVersion)
            Messages = "Exception Error in PopulateDataSource: " & ex.ToString
        End Try

    End Sub

    Function GetFieldNames_MDB(MDBName As String, DBTable As String, Optional ByVal Provider As String = "Provider=Microsoft.Jet.OLEDB.4.0;") As String

    End Function

    Function GetFieldNames_accDB(accDBName As String, DBTable As String, Optional ByRef Provider As String = "Provider=Microsoft.ACE.OLEDB.12.0;") As String

    End Function

    Public Function Find_me(ByRef DBName As String, ByRef Provider As String, ByVal DBTable As String,
                            ByVal SearchField As String, ByVal SearchText As String, ByVal SearchType As String,
                            ByVal ReturnField As String, ByRef ReturnValue As String,
                            Optional ByRef AllFieldValues As Object() = Nothing,
                            Optional ByRef AllFieldNames As String() = Nothing, Optional ByVal Criteria As String = "",
                            Optional ByRef FieldnameLoop As String = "", Optional SortField As String = "ID",
                            Optional ByVal Reversed As Boolean = False, Optional ByRef myMessages As String = "",
                            Optional ByVal Operation As String = " = ") As Boolean
        Dim QryStr As String
        Dim cmd As OleDbCommand
        Dim DataSource As String
        Dim ConnStr As String
        Dim con As OleDbConnection
        Dim WholeRow As Object()
        Dim dr1 As OleDbDataReader
        Dim dr2 As OleDbDataReader
        Dim NumFields As Integer
        Dim dtFieldInfo As DataTable
        Dim Fieldnames As String()
        Dim IDX As Integer
        Dim strDateOut As String = ""
        Dim strDateIn As String = ""

        'The parameters have changed in this function as of 07-Mar-2017 at 07:35 -
        ' There is now an additional parameter that HAS to be passed and that is the SEARCH TYPE.
        ' This is simply a STRING or INTEGER or FLOAT . This is used in the first WHERE clause.
        ' When searching for an ID with an extra CRITERIA string passed - mainly for the HISTORY search
        ' when searching for the right ID in the registry table to transfer it to the history table
        ' - the ID value will need to be passed as integer - to stop quotes being put around the value variable.

        Find_me = False
        NumFields = 0
        ReDim Fieldnames(1)
        If Len(DBName) = 0 Then
            'MsgBox("Error in Find_me: NO Database Name specified", vbOK, "Error in Asset Register")
            myMessages = "Error in Find_me: NO Database Name specified"
            Find_me = False
            Exit Function
        Else
            If InStr(DBName, ".MDB") > 0 Then
                If Len(Provider) = 0 Then
                    Provider = "Provider=Microsoft.Jet.OLEDB.4.0;"
                End If
            ElseIf InStr(DBName, ".accdb") Then
                If Len(Provider) = 0 Then
                    Provider = "Provider=Microsoft.ACE.OLEDB.12.0;"
                End If
            Else
                'MsgBox("Error in Find_me: Unrecognised Database type: " & DBName, vbOK, "Error in Asset Register " & frmMainGIForm.myVersion)
                myMessages = "Error in Find_me: Unrecognised Database type:" & DBName
                Find_me = False
                Exit Function
            End If

        End If
        If Len(SearchField) = 0 Then
            'MsgBox("Error in Find_Me: Search Field is blank", vbOK, "Error in Asset Register " & frmMainGIForm.myVersion)
            myMessages = "Error in Find_Me: Search Field is blank"
            Exit Function
        End If
        If Len(SearchText) = 0 Then
            'MsgBox("Error in Find_Me: Search Text is blank", vbOK, "Error in Asset Register " & frmMainGIForm.myVersion)
            myMessages = "Error in Find_Me: Search Text is blank"
            Exit Function
        End If
        If Not Right(DBName, 1) = ";" Then
            DBName = DBName & ";"
        End If
        If Not Right(Provider, 1) = ";" Then
            Provider = Provider & ";"
        End If
        If Len(DBTable) = 0 Then
            'MsgBox("Error in Find_me: No DB Table specified")
            myMessages = "Error in Find_me: No DB Table specified"
            Find_me = False
            Exit Function

        End If
        ReturnValue = ""
        Find_me = False
        DataSource = frmMainGIForm.dbsource & DBName
        ConnStr = Provider & DataSource
        con = New OleDbConnection(ConnStr)
        Try
            con.Open()
            If UCase(SearchType) = "STRING" Then
                'QryStr = "SELECT * FROM " & DBTable & " WHERE (" & SearchField & " " & Operation & " '" & SearchText & "')"
                QryStr = "SELECT * FROM " & DBTable & " WHERE (" & SearchField & Operation & "'" & SearchText & "')"
            ElseIf UCase(SearchType) = "INTEGER" Then
                QryStr = "SELECT * FROM " & DBTable & " WHERE (" & SearchField & Operation & SearchText & ")"
            Else
                QryStr = "SELECT * FROM " & DBTable & " WHERE (" & SearchField & Operation & SearchText & "f" & ")"
            End If
            If Len(Criteria) > 0 Then
                'QryStr = "SELECT * FROM " & DBTable & " WHERE (" & SearchField & " = '" & SearchText & "')"
                QryStr = QryStr & " AND " & Criteria
            End If
            If Len(SortField) > 0 Then
                QryStr = QryStr + " ORDER BY " + SortField
                If Reversed = False Then
                    QryStr = QryStr + " ASC"
                Else
                    QryStr = QryStr + " DESC"
                End If
            End If
            '
            cmd = New OleDbCommand(QryStr, con)
            dr1 = cmd.ExecuteReader()

            NumFields = dr1.FieldCount - 1
            ReDim Fieldnames(NumFields)
            IDX = 0
            While IDX <= NumFields
                Fieldnames(IDX) = dr1.GetName(IDX)

                IDX = IDX + 1
            End While
            AllFieldNames = Fieldnames
            dr1.Close()
            dr2 = cmd.ExecuteReader()
            ReDim WholeRow(NumFields)
            If dr2.HasRows Then
                While dr2.Read() 'cycles through each record found:
                    'DescriptionText.Text = dr("Description").ToString
                    'CostText.Text = dr("Cost").ToString
                    'PriceText.Text = dr("Price").ToString
                    If Not IsDBNull(dr2(SearchField).ToString) Then
                        If UCase(dr2(SearchField).ToString) = UCase(SearchText) Then
                            If Len(ReturnField) > 0 Then
                                ReturnValue = dr2(ReturnField).ToString
                                Find_me = True
                            End If
                            'Got to extract the whole row:
                            dr2.GetValues(WholeRow)
                            AllFieldValues = WholeRow

                            If UCase(DBTable) = "TBLREGISTER" Then
                                strDateOut = dr2("DateOut").ToString
                                strDateIn = dr2("DateIn").ToString
                            End If

                            If Len(FieldnameLoop) > 0 Then
                                'When returning Asset - ignore items in registry that have already been returned - ie have both DateIN and DateOUT filled in.
                                If Len(strDateOut) > 0 Then
                                    If Len(strDateIn) = 0 Then 'YES asset has been taken out but not returned yet - now exit.
                                        Find_me = True
                                        Exit While
                                    End If
                                End If
                            Else
                                Find_me = True
                                Exit While
                            End If
                        End If
                    End If
                End While
            Else
                'No Rows
                Find_me = False
            End If
            dr2.Close()
        Catch ex As Exception
            'MsgBox("Exception Error in Find_me: " & ex.ToString, vbOK, "Exception Error in Asset Register v" & frmMainGIForm.myVersion)
            myMessages = "Exception Error in Find_me: " & ex.ToString
        End Try

        con.Close()

    End Function

    Function GetValuebyFieldname(ByVal SearchObjects As Object(), ByVal Fieldnames As String(), ByVal ReturnField As String) As String
        Dim fieldidx As Int32
        Dim ReturnValue As String
        Dim NumFields As Integer

        'Need to cycle through the Fieldnames and find the correct index.
        'THEN apply it to the SearchObjects to get the value.
        'NumFields = GetNumFields("", "SELECT * FROM " & dbTable, "AssetRegister.accdb", Fieldnames)
        ReturnValue = ""
        GetValuebyFieldname = ""

        If IsNothing(SearchObjects) Then
            Exit Function
        End If
        NumFields = UBound(Fieldnames)
        Try
            For fieldidx = 0 To NumFields 'does this pick up the last field ???
                If UCase(ReturnField) = UCase(Fieldnames(fieldidx)) Then
                    ReturnValue = SearchObjects(fieldidx).ToString
                    Return ReturnValue
                    Exit For
                End If
            Next
        Catch ex As Exception
            MsgBox("Exception Error in GetValuebyFieldname: " & ex.ToString, vbOK, "Exception Error in Asset Register v" & frmMainGIForm.myVersion)
        End Try
        GetValuebyFieldname = ReturnValue

    End Function

    Function ValidateEntries(ByRef FieldValues As String) As String
        Dim FieldIDX As Integer
        Dim FieldArr As String()

        'Validate each entry before it gets passed to Insert or Update procedures:
        'Also any Dates that need a current datestamp.

        ValidateEntries = ""
        ReDim FieldArr(1)
        Try
            FieldArr = strToStringArray(FieldValues, ",")
            For FieldIDX = 0 To UBound(FieldArr)
                If Len(FieldArr(FieldIDX)) = 0 Then
                    'NO VALUE EXISTS:
                    'GET THE FIELD TYPE:
                    'hOW DO WE KNOW WHICH FIELDS NEED SPECIAL ATTENTION ? SUCH AS DATESTAMP FIELDS ?
                    'DateOfHistory - leave as NULL upon INSERT - as later this will be updated:
                    'DateOfHistory = NULL
                    'DateOfEntry = Timestamp

                End If
            Next
        Catch ex As Exception
            MsgBox("Exception Error in ValidateEntries: " & ex.ToString, vbOK, "Exception Error in Asset Register v" & frmMainGIForm.myVersion)
        End Try

        ValidateEntries = FieldValues

    End Function

    Function Get_Filename(Optional ByVal InitialDir As String = "",
                          Optional ByVal FileFilter As String = "CSV Files (*.csv)|*.csv" & "|EXCEL Files (*.xls?)|*.xls?" & "|txt Files (*.txt)|*.txt" & "|All Files (*.*)|*.*",
                          Optional ByVal FilterIndexSelected As Integer = 1) As String
        Dim CSVFileDlg As New OpenFileDialog()
        Dim CSVFilename As String = ""

        'frmMainGIForm.InitialDir = "..\PDF Files"
        If Len(InitialDir) > 0 Then
            CSVFileDlg.InitialDirectory = InitialDir
        Else
            CSVFileDlg.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        End If
        Get_Filename = ""

        CSVFileDlg.Filter = FileFilter
        CSVFileDlg.FilterIndex = FilterIndexSelected
        CSVFileDlg.RestoreDirectory = False
        CSVFileDlg.ShowDialog()

        CSVFilename = CSVFileDlg.FileName
        Get_Filename = CSVFilename

    End Function

    Sub Convert_OldEmpNo_NewEmpNo(ByVal TheCodeType As String, ByVal TheCodeNumber As String, ByRef OldCode As String, ByRef NewCode As String, Optional LimitNumberIndex As Integer = 1)
        Dim CodeNumber As Long = 0
        Dim StartNumber As Integer = 1
        Dim Percentage As Single = 0.0F
        Dim Extract As String
        Dim CodeType As String

        OldCode = ""
        NewCode = ""
        If Len(TheCodeType) > 0 Then
            CodeType = TheCodeType
        End If
        If IsNumeric(TheCodeNumber) Then
            CodeNumber = CLng(TheCodeNumber)
        Else
            If Len(TheCodeNumber) = 0 Then
                NewCode = TheCodeNumber
                Exit Sub
            End If
            If UCase(Left(TheCodeNumber, 3)) = "AGY" Then
                Extract = Mid(TheCodeNumber, 4, Len(TheCodeNumber))
                If IsNumeric(Extract) Then
                    CodeNumber = CLng(Extract)
                    CodeType = "AGENCY"
                Else
                    NewCode = TheCodeNumber
                    Exit Sub
                End If
            End If
        End If

        If UCase(CodeType) = "AGENCY" Then
            If LimitNumberIndex + (CodeNumber - 1) < 10 Then
                OldCode = "AGY" & "00" & CStr((CodeNumber - 1) + LimitNumberIndex)
            ElseIf LimitNumberIndex + (CodeNumber - 1) < 100 Then
                OldCode = "AGY" & "0" & CStr((CodeNumber - 1) + LimitNumberIndex)
            Else
                OldCode = "AGY" & CStr((CodeNumber - 1) + LimitNumberIndex)
            End If

            If LimitNumberIndex + (CodeNumber - 1) < 10 Then
                NewCode = "AGY" & "000" & CStr((CodeNumber - 1) + LimitNumberIndex)
            ElseIf LimitNumberIndex + (CodeNumber - 1) < 100 Then
                NewCode = "AGY" & "00" & CStr((CodeNumber - 1) + LimitNumberIndex)
            ElseIf LimitNumberIndex + (CodeNumber - 1) < 1000 Then
                NewCode = "AGY" & "0" & CStr((CodeNumber - 1) + LimitNumberIndex)
            Else
                NewCode = "AGY" & CStr((CodeNumber - 1) + LimitNumberIndex)
            End If
        Else
            If LimitNumberIndex + (CodeNumber - 1) < 10 Then
                OldCode = "00000" & CStr((CodeNumber - 1) + LimitNumberIndex)
            ElseIf LimitNumberIndex + (CodeNumber - 1) < 100 Then
                OldCode = "0000" & CStr((CodeNumber - 1) + LimitNumberIndex)
            ElseIf LimitNumberIndex + (CodeNumber - 1) < 1000 Then
                OldCode = "000" & CStr((CodeNumber - 1) + LimitNumberIndex)
            ElseIf LimitNumberIndex + (CodeNumber - 1) < 10000 Then
                OldCode = "00" & CStr((CodeNumber - 1) + LimitNumberIndex)
            ElseIf LimitNumberIndex + (CodeNumber - 1) < 100000 Then
                OldCode = "0" & CStr((CodeNumber - 1) + LimitNumberIndex)
            Else
                OldCode = CStr((CodeNumber - 1) + LimitNumberIndex)
            End If

            If LimitNumberIndex + (CodeNumber - 1) < 10 Then
                NewCode = "000000" & CStr((CodeNumber - 1) + LimitNumberIndex)
            ElseIf LimitNumberIndex + (CodeNumber - 1) < 100 Then
                NewCode = "00000" & CStr((CodeNumber - 1) + LimitNumberIndex)
            ElseIf LimitNumberIndex + (CodeNumber - 1) < 1000 Then
                NewCode = "0000" & CStr((CodeNumber - 1) + LimitNumberIndex)
            ElseIf LimitNumberIndex + (CodeNumber - 1) < 10000 Then
                NewCode = "000" & CStr((CodeNumber - 1) + LimitNumberIndex)
            ElseIf LimitNumberIndex + (CodeNumber - 1) < 100000 Then
                NewCode = "00" & CStr((CodeNumber - 1) + LimitNumberIndex)
            ElseIf LimitNumberIndex + (CodeNumber - 1) < 1000000 Then
                NewCode = "0" & CStr((CodeNumber - 1) + LimitNumberIndex)
            Else
                NewCode = CStr((CodeNumber - 1) + LimitNumberIndex)
            End If
        End If


    End Sub

End Module

