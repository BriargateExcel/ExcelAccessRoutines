Attribute VB_Name = "ExcelAccessRoutines"
Option Explicit
'@Folder ExcelAccess

Const TARGET_DB As String = "Headcount with Queries.accdb"
Const AccountingCalendarTable As String = "AccountingCalendarTable"
Const ControlAccountTable As String = "ControlAccountTable"
Const PoPTable As String = "PoPTable"

' References:
' Windows Script Host Object Model library
'   Used to get the network to create the database
' Microsoft ActiveX Data Objects 6.1 Library
'   Brings in ADODB (connections, recordsets, and fields)
' Microsoft ADO Ext. 6.0 for DDL and Security
'   Brings in ADOX (catalog, table, index)
' Microsoft Scripting Runtime
'   Brings in FileSystemObject

Private Type NetworkDataType
    Drv As String
    NtWk As WshNetwork
    LocalFolder As Boolean
    DataBasePath As String
    Cnn As ADODB.Connection
    Cat As ADOX.Catalog
    Initialized As Boolean
End Type

Private pNetworkData As NetworkDataType

Public Sub test()
    Initialize
    MsgBox GetLastDayOfMonthFromCompanyAndAccountingMonth("LM", #11/1/2018#)
    
    '    MsgBox GetFirstDayOfWeekFromLMWeek("2014-50")
    '    MsgBox GetWeeklyHoursFromControlAccountAndDate("8J2PS01132-01U", #12/1/2016#)
    '    MsgBox GetMonthlyHoursFromControlAccountAndDate("8J2PS01132-01U", #12/1/2016#)
    '    MsgBox GetHoursByCompanyAndDateRange("LM", #11/1/2018#, #12/31/2018#)
    '    MsgBox GetHoursByCompanyAndDateRangeAndPeriod("LM", #11/1/2018#, #12/31/2018#, "OY3")
    '    MsgBox GetHoursByCompanyAndDateRangeAndPeriod("LM", #11/1/2018#, #12/31/2018#, "OY4")
    '
    '    MsgBox GetCompanyFromControlAccount("8G3SN04311-03")
    Wrapup
End Sub                                          ' test

Public Function GetLastDayOfMonthFromCompanyAndAccountingMonth( _
       ByVal Company As String, _
       ByVal AccountingMonth As Date) _
        As Date

    If Day(AccountingMonth) <> 1 Then
        MsgBox "All accounting months have a day = 1", _
               vbOKOnly Or vbCritical, _
               "Not An Accounting Month"
        Stop
    End If

    Dim SQLQuery As String
    SQLQuery = "SELECT Date as Dt FROM " & AccountingCalendarTable & " " & _
               "WHERE " & AccountingCalendarTable & "." & Company & "Month = #" & _
               AccountingMonth & "# ORDER BY Date DESC;"

    GetLastDayOfMonthFromCompanyAndAccountingMonth = GetDateFromDatabase(SQLQuery)

End Function                                     ' GetLastDayOfMonthFromCompanyAndAccountingMonth

Public Function GetFirstDayOfMonthFromCompanyAndAccountingMonth( _
       ByVal Company As String, _
       ByVal AccountingMonth As Date) _
        As Date

    If Day(AccountingMonth) <> 1 Then
        MsgBox "All accounting months have a day = 1", _
               vbOKOnly Or vbCritical, _
               "Not An Accounting Month"
        Stop
    End If

    Dim SQLQuery As String
    SQLQuery = "SELECT Date as Dt FROM " & AccountingCalendarTable & " " & _
               "WHERE " & AccountingCalendarTable & "." & Company & "Month = #" & _
               AccountingMonth & "# ORDER BY Date;"

    GetFirstDayOfMonthFromCompanyAndAccountingMonth = GetDateFromDatabase(SQLQuery)

End Function                                     ' GetFirstDayOfMonthFromCompanyAndAccountingMonth

Public Function GetNumberOfWeeksFromControlAccountAndAccountingMonth( _
       ByVal ControlAccount As String, _
       ByVal AccountingMonth As Date) _
        As Long

    If Day(AccountingMonth) <> 1 Then
        MsgBox "All accounting months have a day = 1", _
               vbOKOnly Or vbCritical, _
               "Not An Accounting Month"
        Stop
    End If

    If DateInControlAccount(AccountingMonth, ControlAccount) Then
        Dim StartOfMonth As Date
        StartOfMonth = GetFirstDayOfMonthFromControlAccountAndAccountingMonth(ControlAccount, AccountingMonth)

        Dim EndOfMonth As Date
        EndOfMonth = GetLastDayOfMonthFromControlAccountAndAccountingMonth(ControlAccount, AccountingMonth)
    
        GetNumberOfWeeksFromControlAccountAndAccountingMonth = (EndOfMonth - StartOfMonth + 1) / 7

    Else
        GetNumberOfWeeksFromControlAccountAndAccountingMonth = 0
    End If

End Function                                     ' GetNumberOfWeeksFromControlAccountAndAccountingMonth

Public Function DateInControlAccount( _
        ByVal Dt As Date, _
        ByVal ControlAccount As String) _
        As Boolean

    Dim PoP As String
    PoP = GetPeriodFromControlAccount(ControlAccount)

    Dim StartOfPeriod As Date
    StartOfPeriod = GetStartDateOfPeriod(PoP)

    Dim EndOfPeriod As Date
    EndOfPeriod = GetEndDateOfPeriod(PoP)

    DateInControlAccount = (Dt >= StartOfPeriod And Dt <= EndOfPeriod)

End Function                                     ' DateInControlAccount

Public Function GetStartDateOfPeriod(ByVal PoP As String) As Date
    Dim SQLQuery As String
    SQLQuery = "SELECT StartDate AS Dt FROM " & _
               PoPTable & " WHERE PeriodName = """ & _
               PoP & """;"

    GetStartDateOfPeriod = GetDateFromDatabase(SQLQuery)

End Function                                     ' GetStartDateOfPeriod

Public Function GetEndDateOfPeriod(ByVal PoP As String) As Date
    Dim SQLQuery As String
    SQLQuery = "SELECT EndDate AS Dt FROM " & _
               PoPTable & " WHERE PeriodName = """ & _
               PoP & """;"

    GetEndDateOfPeriod = GetDateFromDatabase(SQLQuery)

End Function                                     ' GetEndDateOfPeriod

Public Function GetLastDayOfFirstWeekFromControlAccountAndAccountingMonth( _
       ByVal ControlAccount As String, _
       ByVal AccountingMonth As Date) _
        As Date
    ' Return the last day of the first week of the month for a given control account

    Dim PoP As String
    PoP = GetPeriodFromControlAccount(ControlAccount)
    
    Dim AccountingCalendar As String
    AccountingCalendar = GetAccountingCalendarFromControlAccount(ControlAccount)
    
    Dim Dt As Date
    Dt = GetFirstDayOfMonthFromControlAccountAndAccountingMonth(ControlAccount, AccountingMonth)
    
    GetLastDayOfFirstWeekFromControlAccountAndAccountingMonth = GetLastDayOfWeekFromControlAccountAndDate(ControlAccount, Dt)

End Function                                     ' GetLastDayOfFirstWeekFromControlAccountAndAccountingMonth

Public Function GetLastDayOfWeekFromControlAccountAndDate( _
       ByVal ControlAccount As String, _
       ByVal Dt As Date) _
        As Date

    Dim PoP As String
    PoP = GetPeriodFromControlAccount(ControlAccount)

    Dim AccountingCalendar As String
    AccountingCalendar = GetAccountingCalendarFromControlAccount(ControlAccount)
    
    Dim LMAccountingWeek As String
    LMAccountingWeek = GetLMWeekFromDate(Dt)

    Dim SQLQuery As String
    SQLQuery = "SELECT Date AS Dt " & _
               "FROM " & AccountingCalendarTable & " " & _
               "WHERE " & AccountingCalendarTable & ".LMWeekNumber = """ & _
               LMAccountingWeek & """ AND " & _
               AccountingCalendarTable & "." & PoP & " = 1 " & _
               "ORDER BY Date DESC;"
               
    GetLastDayOfWeekFromControlAccountAndDate = GetDateFromDatabase(SQLQuery)

End Function                                     ' GetLastDayOfWeekFromControlAccountAndDate

Public Function GetLastDayOfMonthFromControlAccountAndAccountingMonth( _
       ByVal ControlAccount As String, _
       ByVal AccountingMonth As Date) _
        As Date

    If Day(AccountingMonth) <> 1 Then
        MsgBox "All accounting months have a day = 1", _
               vbOKOnly Or vbCritical, _
               "Not An Accounting Month"
        Stop
    End If

    Dim PoP As String
    PoP = GetPeriodFromControlAccount(ControlAccount)

    Dim AccountingCalendar As String
    AccountingCalendar = GetAccountingCalendarFromControlAccount(ControlAccount)

    Dim SQLQuery As String
    SQLQuery = "SELECT Date AS Dt " & _
               "FROM " & AccountingCalendarTable & " " & _
               "WHERE " & AccountingCalendarTable & "." & AccountingCalendar & "Month = " & _
               "#" & AccountingMonth & "# AND " & _
               AccountingCalendarTable & "." & PoP & " = 1 " & _
               "Order By Date DESC;"
               
    GetLastDayOfMonthFromControlAccountAndAccountingMonth = GetDateFromDatabase(SQLQuery)

End Function                                     ' GetLastDayOfMonthFromControlAccountAndAccountingMonth

Public Function GetFirstDayOfMonthFromControlAccountAndAccountingMonth( _
       ByVal ControlAccount As String, _
       ByVal AccountingMonth As Date) _
        As Date
        
    If Day(AccountingMonth) <> 1 Then
        MsgBox "All accounting months have a day = 1", _
               vbOKOnly Or vbCritical, _
               "Not An Accounting Month"
        Stop
    End If

    If Day(AccountingMonth) <> 1 Then
        MsgBox "All accounting months have a day = 1", _
               vbOKOnly Or vbCritical, _
               "Not An Accounting Month"
        Stop
    End If

    Dim PoP As String
    PoP = GetPeriodFromControlAccount(ControlAccount)

    Dim AccountingCalendar As String
    AccountingCalendar = GetAccountingCalendarFromControlAccount(ControlAccount)

    Dim SQLQuery As String
    SQLQuery = "SELECT Date AS Dt " & _
               "FROM " & AccountingCalendarTable & " " & _
               "WHERE " & AccountingCalendarTable & "." & AccountingCalendar & "Month = " & _
               "#" & AccountingMonth & "# AND " & _
               AccountingCalendarTable & "." & PoP & " = 1 " & _
               "ORDER BY Date;"
               
    GetFirstDayOfMonthFromControlAccountAndAccountingMonth = GetDateFromDatabase(SQLQuery)

End Function                                     ' GetFirstDayOfMonthFromControlAccountAndAccountingMonth

Public Function GetLastLMWeek() As String

    Dim SQLQuery As String
    SQLQuery = "SELECT MAX (LMWeekNumber)  AS LMWeek FROM " & AccountingCalendarTable & ";"

    GetLastLMWeek = GetStringFromDatabase(SQLQuery)

End Function                                     ' GetLastLMWeek

Public Function GetLastAccountingMonth() As Date

    Dim SQLQuery As String
    SQLQuery = "SELECT MAX (LMMonth)  AS LMWeek FROM " & AccountingCalendarTable & ";"

    GetLastAccountingMonth = GetDateFromDatabase(SQLQuery)

End Function                                     ' GetLastAccountingMonth

Public Function GetFirstLMWeek() As String

    Dim SQLQuery As String
    SQLQuery = "SELECT MIN (LMWeekNumber)  AS LMWeek FROM " & AccountingCalendarTable & ";"

    GetFirstLMWeek = GetStringFromDatabase(SQLQuery)

End Function                                     ' GetFirstLMWeek

Public Function GetFirstAccountingMonth() As Date

    Dim SQLQuery As String
    SQLQuery = "SELECT MIN (LMMonth)  AS LMWeek FROM " & AccountingCalendarTable & ";"

    GetFirstAccountingMonth = GetDateFromDatabase(SQLQuery)

End Function                                     ' GetFirstAccountingMonth

Public Function GetFirstDayOfWeekFromLMWeek(ByVal LMWeek As String) As Date

    Dim SQLQuery As String
    SQLQuery = "SELECT Date AS Dt " & _
               "FROM " & AccountingCalendarTable & " " & _
               "WHERE " & AccountingCalendarTable & ".LMWeekNumber = """ & _
               LMWeek & """ ORDER BY Date;"
               
    GetFirstDayOfWeekFromLMWeek = GetDateFromDatabase(SQLQuery)

End Function                                     ' GetFirstDayOfWeekFromLMWeek

Public Function GetMonthlyHoursFromControlAccountAndDate( _
       ByVal ControlAccount As String, _
       ByVal Dt As Date) _
        As Single
        
    ' Get accounting calendar from Control Account
    Dim AccountingCalendar As String
    AccountingCalendar = GetAccountingCalendarFromControlAccount(ControlAccount)

    ' Get PoP from Control Account
    Dim PoP As String
    PoP = GetPeriodFromControlAccount(ControlAccount)
    
    ' Get accounting month from accounting calendar and date
    Dim AccountingMonth As Date
    AccountingMonth = GetLMMonthFromCalendarAndDate(AccountingCalendar, Dt)
    
    Dim AccountingHoursColumn As String
    AccountingHoursColumn = AccountingCalendar & "Hours"

    Dim SQLQuery As String
    SQLQuery = "SELECT Sum(" & AccountingHoursColumn & ") AS Hours " & _
               "FROM " & AccountingCalendarTable & " " & _
               "WHERE " & AccountingCalendarTable & "." & AccountingCalendar & "Month = " & _
               "#" & AccountingMonth & "#) AND " & _
               AccountingCalendarTable & "." & PoP & " = 1;"
    
    GetMonthlyHoursFromControlAccountAndDate = GetSingleFromDatabase(SQLQuery)
    
End Function                                     ' GetMonthlyHoursFromControlAccountAndDate

Public Function GetWeeklyHoursFromControlAccountAndDate( _
       ByVal ControlAccount As String, _
       ByVal Dt As Date) _
        As Single
        
    ' Get accounting calendar from Control Account
    Dim AccountingCalendar As String
    AccountingCalendar = GetAccountingCalendarFromControlAccount(ControlAccount)

    ' Get PoP from Control Account
    Dim PoP As String
    PoP = GetPeriodFromControlAccount(ControlAccount)
    
    ' Get accounting week from accounting calendar and date
    Dim AccountingWeek As String
    AccountingWeek = GetLMWeekFromDate(Dt)
    
    Dim AccountingHoursColumn As String
    AccountingHoursColumn = AccountingCalendar & "Hours"

    Dim SQLQuery As String
    SQLQuery = "SELECT Sum(" & AccountingHoursColumn & ") AS Hours " & _
               "FROM " & AccountingCalendarTable & " " & _
               "WHERE " & AccountingCalendarTable & ".LMWeekNumber = """ & _
               AccountingWeek & """ AND " & _
               AccountingCalendarTable & "." & PoP & " = 1;"
    
    GetWeeklyHoursFromControlAccountAndDate = GetSingleFromDatabase(SQLQuery)
    
End Function                                     ' GetWeeklyHoursFromControlAccountAndDate

Public Function GetLMMonthFromCalendarAndDate( _
       ByVal Calendar As String, _
       ByVal Dt As Date) _
        As Date

    Dim MonthColumn As String
    MonthColumn = Calendar & "Month"
    
    Dim SQLQuery As String
    SQLQuery = "SELECT " & MonthColumn & " AS AccountingMonth " & _
               "FROM " & AccountingCalendarTable & " " & _
               "WHERE " & AccountingCalendarTable & ".Date = " & _
               "#" & Dt & "#);"
    
    GetLMMonthFromCalendarAndDate = GetDateFromDatabase(SQLQuery)
    
End Function                                     ' GetLMMonthFromCalendarAndDate

Public Function GetLMWeekFromDate( _
       ByVal Dt As Date) _
        As String

    Dim SQLQuery As String
    SQLQuery = "SELECT LMWeekNumber AS LMWeek " & _
               "FROM " & AccountingCalendarTable & " " & _
               "WHERE " & AccountingCalendarTable & ".Date = " & _
               "#" & Dt & "#;"

    GetLMWeekFromDate = GetStringFromDatabase(SQLQuery)

End Function                                     ' GetLMWeekFromDate

Public Function GetPeriodFromControlAccount(ByVal CtlAcct As String) As String

    Dim SQLQuery As String
    SQLQuery = "SELECT PeriodofPerformance AS Period " & _
               "FROM " & ControlAccountTable & " " & _
               "WHERE " & ControlAccountTable & ".ControlAccount = """ & _
               CtlAcct & """;"
    
    GetPeriodFromControlAccount = GetStringFromDatabase(SQLQuery)
    
End Function                                     ' GetPeriodFromControlAccount

Public Function GetAccountingCalendarFromControlAccount(ByVal CtlAcct As String) As String

    GetAccountingCalendarFromControlAccount = GetCompanyFromControlAccount(CtlAcct)
    
End Function                                     ' GetAccountingCalendarFromControlAccount

Public Function GetCompanyFromControlAccount(ByVal CtlAcct As String) As String

    Dim SQLQuery As String
    SQLQuery = "SELECT AccountingCalendar AS CompanyName " & _
               "FROM " & ControlAccountTable & " " & _
               "WHERE " & ControlAccountTable & ".ControlAccount = """ & _
               CtlAcct & """;"
    
    GetCompanyFromControlAccount = GetStringFromDatabase(SQLQuery)
    
End Function                                     ' GetCompanyFromControlAccount

Public Function GetStringFromDatabase(ByVal SQLQuery As String) As String

    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.Open Source:=SQLQuery, _
             ActiveConnection:=pNetworkData.Cnn, _
             CursorType:=adOpenDynamic, _
             LockType:=adLockOptimistic
    
    GetStringFromDatabase = rst.Fields(0)
    
End Function                                     ' GetStringFromDatabase

Public Function GetVariantFromDatabase(ByVal SQLQuery As String) As Variant

    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.Open Source:=SQLQuery, _
             ActiveConnection:=pNetworkData.Cnn, _
             CursorType:=adOpenForwardOnly, _
             LockType:=adLockOptimistic
    
    Set GetVariantFromDatabase = rst.Fields
    
End Function                                     ' GetVariantFromDatabase

Public Function GetDateFromDatabase(ByVal SQLQuery As String) As Date

    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.Open Source:=SQLQuery, _
             ActiveConnection:=pNetworkData.Cnn, _
             CursorType:=adOpenDynamic, _
             LockType:=adLockOptimistic
    
    GetDateFromDatabase = rst.Fields(0)
    
End Function                                     ' GetDateFromDatabase

Public Function GetLongFromDatabase(ByVal SQLQuery As String) As Long

    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.Open Source:=SQLQuery, _
             ActiveConnection:=pNetworkData.Cnn, _
             CursorType:=adOpenDynamic, _
             LockType:=adLockOptimistic
    
    GetLongFromDatabase = rst.Fields(0)
    
End Function                                     ' GetLongFromDatabase

Public Function GetSingleFromDatabase(ByVal SQLQuery As String) As Single

    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.Open Source:=SQLQuery, _
             ActiveConnection:=pNetworkData.Cnn, _
             CursorType:=adOpenDynamic, _
             LockType:=adLockOptimistic
    
    If IsNull(rst.Fields(0)) Then
        GetSingleFromDatabase = 0
    Else
        GetSingleFromDatabase = rst.Fields(0)
    End If
    
End Function                                     ' GetSingleFromDatabase

Public Function GetHoursByCompanyAndDateRange( _
       ByVal Company As String, _
       ByVal StartDate As Variant, _
       ByVal EndDate As Variant) _
        As Single

    Dim SD As String
    SD = VariantToDateString(StartDate)
    
    Dim ED As String
    ED = VariantToDateString(EndDate)
    
    Dim SQLQuery As String
    SQLQuery = "SELECT Sum(" & AccountingCalendarTable & "." & _
               Company & "Hours) as Hours " & _
               "FROM " & AccountingCalendarTable & " " & _
               "WHERE " & AccountingCalendarTable & ".Date " & _
               "Between #" & SD & "#) And #" & ED & "#);"
    
    GetHoursByCompanyAndDateRange = GetLongFromDatabase(SQLQuery)
    
End Function                                     ' GetHoursByCompanyAndDateRange

Public Function GetHoursByCompanyAndDateRangeAndPeriod( _
       ByVal Company As String, _
       ByVal StartDate As Variant, _
       ByVal EndDate As Variant, _
       ByVal Period As String) _
        As Single

    Dim SD As String
    SD = VariantToDateString(StartDate)
    
    Dim ED As String
    ED = VariantToDateString(EndDate)
    
    Dim SQLQuery As String
    SQLQuery = "SELECT Sum(" & AccountingCalendarTable & "." & _
               Company & "Hours) as Hours " & _
               "FROM " & AccountingCalendarTable & " " & _
               "WHERE " & AccountingCalendarTable & ".Date " & _
               "Between #" & SD & "#) And #" & ED & "#) " & _
               "AND " & AccountingCalendarTable & "." & Period & "=1;"
    
    GetHoursByCompanyAndDateRangeAndPeriod = GetSingleFromDatabase(SQLQuery)
    
End Function                                     ' GetHoursByCompanyAndDateRangeAndPeriod

Public Function VariantToDateString(ByVal InputDate As Variant) As String
    Select Case VarType(InputDate)
    Case vbString
        VariantToDateString = InputDate
    Case vbLong, vbDate, vbInteger
        ' Convert to string
        VariantToDateString = Format$(InputDate, "mm/dd/yyyy")
    Case Else
        MsgBox "Unknown data type passed to VariantToDateString." & _
               vbCrLf & _
               "Update the Select Case statement", _
               vbOKOnly Or vbCritical, _
               "Unknown Data Type"
    End Select
End Function                                     ' VariantToDateString

Public Sub PushTablesToAccess()

    Dim Response As String
    Response = MsgBox("Are you sure you want to overwrite the entire database?" & _
                      vbCrLf & _
                      "(" & TARGET_DB & ")", _
                      vbYesNo Or vbExclamation, _
                      "Overwrite Database?")
    Select Case Response
    Case vbYes
        ' Press on
    Case vbNo
        Exit Sub
    End Select

    Initialize

    Dim Wksht As Worksheet
    
    For Each Wksht In ActiveWorkbook.Worksheets
        Dim Tbl As ListObject
        For Each Tbl In Wksht.ListObjects
            Dim rst As ADODB.Recordset
            Set rst = New ADODB.Recordset
            rst.Open Source:=Tbl.Name, _
                     ActiveConnection:=pNetworkData.Cnn, _
                     CursorType:=adOpenDynamic, _
                     LockType:=adLockOptimistic, _
                     Options:=adCmdTable
            
            ' Clear the Access table before loading records
            pNetworkData.Cnn.Execute "DELETE * FROM " & Tbl.Name & ";"
            
            'Load all records from Excel to Access
            Dim I As Long
            Dim J As Long
            For I = 1 To Tbl.ListRows.Count
                rst.AddNew
                For J = 1 To Tbl.ListColumns.Count
                    If Tbl.DataBodyRange(I, J) = vbEmpty Then
                        Dim FieldName As String
                        FieldName = Tbl.HeaderRowRange(1, J)
                        Select Case rst.Fields(FieldName).Type
                        Case adDate
                            rst(Tbl.HeaderRowRange(1, J).Value) = 0
                        Case adLongVarWChar, adVarWChar
                            rst(Tbl.HeaderRowRange(1, J).Value) = vbNullString
                        Case adDouble
                            rst(Tbl.HeaderRowRange(1, J).Value) = 0
                        Case Else
                            Stop
                        End Select
                    Else
                        On Error Resume Next
                        rst(Tbl.HeaderRowRange(1, J).Value) = Tbl.DataBodyRange(I, J)
                        On Error GoTo 0
                        If Err.Number <> 0 Then
                            MsgBox "The " & _
                                   Tbl.HeaderRowRange(1, J) & " field " & _
                                   "in the " & Tbl.Name & _
                                   " table is too small. " & _
                                   "Needs to be at least " & _
                                   Len(Tbl.DataBodyRange(I, J)) & ".", _
                                   vbOKOnly Or vbCritical, _
                                   "Database Field Too Small"
                            Stop
                        End If
                    End If
                Next J
                rst.Update
            Next I
    
            rst.Close
            Set rst = Nothing
        Next Tbl
    Next Wksht
    
    Wrapup

End Sub                                          ' PushTablesToAccess

Public Sub CreateDB_And_Tables()

    Initialize
    
    DeleteOldAndCreateNewDatabase
    
    'create the new database
    Set pNetworkData.Cat = New ADOX.Catalog
    
    Dim CatString As String
    CatString = "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & pNetworkData.DataBasePath & ";"
    
    pNetworkData.Cat.Create CatString
        
    Dim Wksht As Worksheet
    For Each Wksht In ThisWorkbook.Worksheets
        Dim Tbl As ListObject
        For Each Tbl In Wksht.ListObjects
            CreateOneTable Tbl
        Next Tbl
        
    Next Wksht
    
    Wrapup
    
End Sub                                          ' CreateDB_And_Tables

Public Function MapDrive( _
       ByVal Folder As String, _
       ByVal NtWk As Object) As String
       
    Dim I As Long
    Dim Drive As String

    For I = Asc("A") To Asc("Z")
        Drive = Chr$(I) & ":"
        If TestDrive(Drive, Folder, NtWk) Then
            MapDrive = Drive
            Exit Function
        End If
    Next I
    
    MapDrive = "Failed"
     
End Function                                     ' MapDrive

Private Function TestDrive(ByVal Drive As String, _
                           ByVal Folder As String, ByVal NtWk As Object) As Boolean
    On Error GoTo FailedToMap
    NtWk.MapNetworkDrive Drive, Folder
    TestDrive = True
    Exit Function
FailedToMap:
    TestDrive = False
End Function                                     ' TestDrive

Private Sub UnMapDrive(ByRef NetworkData As NetworkDataType)
    
    If NetworkData.LocalFolder Then Exit Sub
    
    On Error Resume Next
    If NetworkData.Drv <> "Failed" Then
        NetworkData.NtWk.RemoveNetworkDrive NetworkData.Drv
    End If
    On Error GoTo 0
    Set NetworkData.NtWk = Nothing
End Sub                                          ' UnMapDrive

Private Sub InitializeNetwork()

    Set pNetworkData.NtWk = New WshNetwork
    
    pNetworkData.DataBasePath = ThisWorkbook.Path & Application.PathSeparator & TARGET_DB
    
    If Mid$(pNetworkData.DataBasePath, 2, 1) = ":" Then
        pNetworkData.Drv = Left$(pNetworkData.DataBasePath, 2)
        pNetworkData.LocalFolder = True
    Else
        pNetworkData.Drv = MapDrive(pNetworkData.DataBasePath, pNetworkData.NtWk)
        pNetworkData.LocalFolder = False
    End If

End Sub                                          ' InitializeNetwork

Private Sub OpenDataBase()

    Set pNetworkData.Cnn = New ADODB.Connection
    
    With pNetworkData.Cnn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open pNetworkData.DataBasePath
    End With

End Sub                                          ' OpenDataBase

Private Sub DeleteOldAndCreateNewDatabase()

    Dim FS As FileSystemObject
    Set FS = New FileSystemObject
    
    Dim Response As String
    Response = MsgBox("Are you sure you want to delete the database?" & _
                      vbCrLf & _
                      "(" & TARGET_DB & ")", _
                      vbYesNo Or vbExclamation, _
                      "Delete Database?")
    Select Case Response
    Case vbYes
        ' Press on
    Case vbNo
        Exit Sub
    End Select
    
    'delete the DB if it already exists
    Dim ErrorNumber As Long
    On Error Resume Next
    Kill pNetworkData.DataBasePath
    ErrorNumber = Err.Number
    On Error GoTo 0
    
    Select Case ErrorNumber
    Case 0
        ' Indicates database was successfully deleted
    Case 70
        MsgBox "The database is open. It must be closed to delete it.", _
               vbOKOnly Or vbCritical, _
               "Database Open"
        Stop
    Case 53
        ' Indicates database didn't exist when trying to delete it
    Case Else
        MsgBox "There was an error trying to delete the database", _
               vbOKOnly Or vbCritical, _
               "Database Deletion Error"
        Stop
    End Select
    
End Sub                                          ' DeleteOldAndCreateNewDatabase

Private Sub CreateOneTable(ByVal Tbl As ListObject)

    Dim AccessTable As ADOX.Table
    Set AccessTable = New ADOX.Table
        
    AccessTable.Name = Tbl.Name
        
    Dim I As Long
    For I = 1 To Tbl.ListColumns.Count
        Dim FieldName As String
        FieldName = Tbl.HeaderRowRange(I)
            
        Select Case VarType(Tbl.DataBodyRange(1, I))
        Case vbString, vbEmpty
            Dim MaxLength As Long
            MaxLength = FindMaxFieldLength(Tbl, I)
            If MaxLength <= 255 Then
                AccessTable.Columns.Append FieldName, adVarWChar, MaxLength
            Else
                AccessTable.Columns.Append FieldName, adLongVarWChar, MaxLength
            End If
        Case vbInteger, vbLong
            AccessTable.Columns.Append FieldName, adInteger
        Case vbSingle, vbDouble
            AccessTable.Columns.Append FieldName, adDouble
        Case vbDate
            AccessTable.Columns.Append FieldName, adDate
        Case Else
            MsgBox "Extend the Select Case Statement with more types", _
                   vbOKOnly Or vbCritical, _
                   "Unknown Data Type"
            Stop
        End Select
    Next I
        
    pNetworkData.Cat.Tables.Append AccessTable
        
End Sub                                          ' CreateOneTable

Private Function FindMaxFieldLength( _
        ByVal Tbl As ListObject, _
        ByVal ColNum As Long) _
        As Long
    
    If Tbl.ListRows.Count = 1 Then
        FindMaxFieldLength = Len(Tbl.DataBodyRange(ColNum, 1))
        Exit Function
    End If
    
    Dim Ary() As Variant
    ReDim Ary(Tbl.ListRows.Count)
    
    Ary = Tbl.ListColumns(ColNum).DataBodyRange
    
    Dim MaxFieldLength As Long
    Dim I As Long
    
    For I = 1 To Tbl.ListRows.Count
        If Len(Ary(I, 1)) > MaxFieldLength Then
            MaxFieldLength = Len(Ary(I, 1))
        End If
    Next I
    
    FindMaxFieldLength = MaxFieldLength
    
End Function                                     ' FindMaxFieldLength

Private Sub Initialize()

    If pNetworkData.Initialized Then Exit Sub

    InitializeNetwork
    
    OpenDataBase
    
    pNetworkData.Initialized = True
    
End Sub                                          ' Initialize

Private Sub Wrapup()

    If Not pNetworkData.Initialized Then Exit Sub

    pNetworkData.Cnn.Close
    Set pNetworkData.Cnn = Nothing
    
    UnMapDrive pNetworkData
    
    Set pNetworkData.Cat = Nothing
    
    pNetworkData.Initialized = False

End Sub                                          ' Wrapup


