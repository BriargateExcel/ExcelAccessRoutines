Attribute VB_Name = "ExcelAccessRoutines"
Option Explicit

Const TARGET_DB = "Headcount Test.accdb"

Private pNtwk As WshNetwork                      ' Windows Script Host Object Model library
Private pDrv As String
Private pDataBasePath As String
Private pCnn As ADODB.Connection                 ' Microsoft ActiveX Data Objects 6.1 Library
Private pLocalFolder As Boolean

Public Sub TransferTableFromAccess()
    
    Dim ShDest As Worksheet
    Set ShDest = Sheets("Table download")

    Set pCnn = New ADODB.Connection
    
    Dim MyConn As String
    MyConn = ThisWorkbook.Path & Application.PathSeparator & TARGET_DB
    
    With pCnn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MyConn
    End With

    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseServer
    rst.Open Source:="tblPopulation", _
             ActiveConnection:=pCnn, _
             CursorType:=adOpenDynamic, _
             LockType:=adLockOptimistic, _
             Options:=adCmdTable
    
    'clear existing data on the sheet
    ShDest.Activate
    Range("A1").CurrentRegion.Offset(1, 0).Clear
    
    'create field headers
    Dim I As Long
    I = 0
    Dim fld As ADODB.Field
    With Range("A1")
        For Each fld In rst.Fields
            .Offset(0, I).Value = fld.Name
            I = I + 1
        Next fld
    End With
     
    'transfer data to Excel
    Range("A2").CopyFromRecordset rst
    
    ' Close the connection
    rst.Close
    pCnn.Close
    Set rst = Nothing
    Set pCnn = Nothing

End Sub

Public Sub PushTablesToAccess()

    Dim Wksht As Worksheet
    Set pCnn = New ADODB.Connection
    
    Dim MyConn As String
    MyConn = ThisWorkbook.Path & Application.PathSeparator & TARGET_DB
    
    With pCnn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MyConn
    End With

    For Each Wksht In ActiveWorkbook.Worksheets
        Dim Tbl As ListObject
        For Each Tbl In Wksht.ListObjects
            Dim rst As ADODB.Recordset
            Set rst = New ADODB.Recordset
            rst.CursorLocation = adUseServer
            rst.Open Source:=Tbl.Name, _
                     ActiveConnection:=pCnn, _
                     CursorType:=adOpenDynamic, _
                     LockType:=adLockOptimistic, _
                     Options:=adCmdTable
            
            'Load all records from Excel to Access
            Dim I As Long
            Dim J As Long
            For I = 1 To Tbl.ListRows.Count
                rst.AddNew
                For J = 1 To Tbl.ListColumns.Count
                    If Tbl.DataBodyRange(I, J) = vbEmpty Then
                        rst(Tbl.HeaderRowRange(1, J).Value) = vbNullString
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
    
    ' Close the connection
    pCnn.Close
    Set pCnn = Nothing

End Sub

Public Sub CreateDB_And_Tables()
    
    DeleteOldAndCreateNewDatabase
    
    'create the new database
    Dim Cat As ADOX.Catalog                      ' Microsoft ADO Ext. 6.0 for DDL and Security
    Set Cat = New ADOX.Catalog
    
    Dim CatString As String
    CatString = "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & pDataBasePath & ";"
    
    Cat.Create CatString
        
    Dim Wksht As Worksheet
    For Each Wksht In ThisWorkbook.Worksheets
        ' todo: deal with multiple tables on one worksheet
        Dim Tbl As ListObject
        For Each Tbl In Wksht.ListObjects
            CreateOneTable Tbl, Cat
        Next Tbl
        
    Next Wksht
    
    UnMapDrive pDrv, pNtwk
    
End Sub

Public Sub AlterOneRecord()
    
    Dim RowNum As Long
    RowNum = ActiveCell.Row
    
    Dim ColNum As Long
    ColNum = Cells(RowNum, 1).Value
    
    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM tblPopulation WHERE PopID = " & ColNum
    
    Set pCnn = New ADODB.Connection
    
    Dim MyConn As String
    MyConn = ThisWorkbook.Path & Application.PathSeparator & TARGET_DB
    
    With pCnn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MyConn
    End With

    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseServer
    rst.Open Source:=SQLQuery, _
             ActiveConnection:=pCnn, _
             CursorType:=adOpenKeyset, _
             LockType:=adLockOptimistic
    
    'Load all records from Excel to Access.
    Dim I As Long
    For I = 2 To 7
        rst(Cells(1, I).Value) = Cells(RowNum, I).Value
    Next I
    rst.Update
    
    ' Close the connection
    rst.Close
    pCnn.Close
    Set rst = Nothing
    Set pCnn = Nothing
    
End Sub

Public Sub DownloadTop20()
    
    Dim ShDest As Worksheet
    Set ShDest = Sheets("Top 20")

    Dim SQLQuery As String
    SQLQuery = "SELECT TOP 20 * FROM tblPopulation ORDER BY Yr_2050 DESC"

    Set pCnn = New ADODB.Connection
    
    Dim MyConn As String
    MyConn = ThisWorkbook.Path & Application.PathSeparator & TARGET_DB
    
    With pCnn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MyConn
    End With

    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseServer
    rst.Open Source:=SQLQuery, _
             ActiveConnection:=pCnn, _
             CursorType:=adOpenForwardOnly, _
             LockType:=adLockOptimistic, _
             Options:=adCmdText
    
    'clear existing data on the sheet
    ShDest.Activate
    Range("A1").CurrentRegion.Offset(1, 0).Clear
    
    'create field headers
    Dim I As Long
    I = 0
    Dim fld As ADODB.Field
    With Range("A1")
        For Each fld In rst.Fields
            .Offset(0, I).Value = fld.Name
            I = I + 1
        Next fld
    End With
     
    'transfer data to Excel
    Range("A2").CopyFromRecordset rst
    
    ' Close the connection
    rst.Close
    pCnn.Close
    Set rst = Nothing
    Set pCnn = Nothing

End Sub

Public Sub DownloadRegion()
    
    Dim ShDest As Worksheet
    Set ShDest = Sheets("Region")

    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM tblPopulation WHERE Region ='" & Range("PickCountry").Value & "'"

    Set pCnn = New ADODB.Connection
    
    Dim MyConn As String
    MyConn = ThisWorkbook.Path & Application.PathSeparator & TARGET_DB
    
    With pCnn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MyConn
    End With

    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseServer
    rst.Open Source:=SQLQuery, _
             ActiveConnection:=pCnn, _
             CursorType:=adOpenForwardOnly, _
             LockType:=adLockOptimistic, _
             Options:=adCmdText
    
    'clear existing data on the sheet
    ShDest.Activate
    Range("A1").CurrentRegion.Clear
    
    'create field headers
    Dim I As Long
    I = 0
    Dim fld As ADODB.Field
    With Range("A1")
        For Each fld In rst.Fields
            .Offset(0, I).Value = fld.Name
            I = I + 1
        Next fld
    End With
     
    'transfer data to Excel
    Range("A2").CopyFromRecordset rst
    
    ' Close the connection
    rst.Close
    pCnn.Close
    Set rst = Nothing
    Set pCnn = Nothing

End Sub

Public Sub DownloadMultiChoice()
    
    Dim ShDest As Worksheet
    Set ShDest = Sheets("Region")
    
    On Error GoTo Err_Handle
        
    'If you got this far the user has made a selection. Proceed with filtering
    Dim ListOfSelections As String
    Dim arChoice() As Variant
    If Range("MultiPick").Cells.Count = 1 Then
        ListOfSelections = "='" & Range("MultiPick").Value & "'"
    Else
        arChoice = WorksheetFunction.Transpose(Range("MultiPick"))
        ListOfSelections = "IN('" & Join(arChoice, "','") & "')"
    End If
    
    'if the items are numbers instead of text, omit the single quotes --
    'ListOfSelections = "IN(" & Join(arChoice, ",") & ")"
        
    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM tblPopulation WHERE Region " & ListOfSelections
    Debug.Print SQLQuery
    
    Set pCnn = New ADODB.Connection
    
    Dim MyConn As String
    MyConn = ThisWorkbook.Path & Application.PathSeparator & TARGET_DB
    
    With pCnn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MyConn
    End With

    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseServer
    rst.Open Source:=SQLQuery, _
             ActiveConnection:=pCnn, _
             CursorType:=adOpenForwardOnly, _
             LockType:=adLockOptimistic, _
             Options:=adCmdText
    
    'clear existing data on the sheet
    Range("A1").CurrentRegion.Clear
    
    'create field headers
    Dim I As Long
    I = 0
    Dim fld As ADODB.Field
    With Range("A1")
        For Each fld In rst.Fields
            .Offset(0, I).Value = fld.Name
            I = I + 1
        Next fld
    End With
    
    'transfer data to Excel
    Range("A2").CopyFromRecordset rst
    
    ' Close the connection
    rst.Close
    pCnn.Close
    Set rst = Nothing
    Set pCnn = Nothing
    
Err_Exit:
    Exit Sub
    
Err_Handle:
    Select Case Err.Number
    Case 1004                                    'range not found; no items in list
        MsgBox "Please select at least one choice"
        Range("M2").Select
        Resume Err_Exit
    Case Else
        MsgBox "Error " & Err.Number & ": " & Err.Description
    End Select
End Sub

Public Sub AddNewField()
  
    Set pCnn = New ADODB.Connection
    
    Dim MyConn As String
    MyConn = ThisWorkbook.Path & Application.PathSeparator & TARGET_DB
  
    'open the connection
    Set pCnn = New ADODB.Connection
    With pCnn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MyConn
    End With
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = pCnn
    'create the field
    cmd.CommandText = "ALTER TABLE tblPopulation ADD Column Region Char(30)"
    cmd.Execute
    Set cmd = Nothing
    pCnn.Close
    Set pCnn = Nothing
End Sub

Public Sub PopulateOneField()

    Sheets("New Field").Activate
    Dim Rw As Long
    Rw = Range("A65536").End(xlUp).Row

    Set pCnn = New ADODB.Connection
    
    Dim MyConn As String
    MyConn = ThisWorkbook.Path & Application.PathSeparator & TARGET_DB
    
    With pCnn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MyConn
    End With

    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseServer
    'Update one field in all records of the table
    Dim I As Long
    Dim SQLQuery As String
    For I = 2 To Rw
        SQLQuery = "SELECT * FROM tblPopulation WHERE PopID = " & Cells(I, 1).Value
        rst.Open Source:=SQLQuery, _
                 ActiveConnection:=pCnn, _
                 CursorType:=adOpenKeyset, _
                 LockType:=adLockOptimistic
        rst(Cells(1, 3).Value) = Cells(I, 3).Value
        rst.Update
        rst.Close
    Next I
    
    ' Close the connection
    pCnn.Close
    Set rst = Nothing
    Set pCnn = Nothing

End Sub

Private Sub DeleteAField()
    
    Set pCnn = New ADODB.Connection
    
    Dim MyConn As String
    MyConn = ThisWorkbook.Path & Application.PathSeparator & TARGET_DB
  
    'open the connection
    Set pCnn = New ADODB.Connection
    With pCnn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open MyConn
    End With
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = pCnn
    'create the field
    cmd.CommandText = "ALTER TABLE tblPopulation DROP Column Region"
    cmd.Execute
    Set cmd = Nothing
    pCnn.Close
    Set pCnn = Nothing
End Sub

Public Function MapDrive( _
       ByVal Folder As String, _
       ByVal Ntwk As Object) As String
       
    Dim I As Long
    Dim Drive As String

    For I = Asc("A") To Asc("Z")
        Drive = Chr$(I) & ":"
        If TestDrive(Drive, Folder, Ntwk) Then
            MapDrive = Drive
            Exit Function
        End If
    Next I
    
    MapDrive = "Failed"
     
End Function                                     ' MapDrive

Private Function TestDrive(ByVal Drive As String, _
                           ByVal Folder As String, ByVal Ntwk As Object) As Boolean
    On Error GoTo FailedToMap
    Ntwk.MapNetworkDrive Drive, Folder
    TestDrive = True
    Exit Function
FailedToMap:
    TestDrive = False
End Function                                     ' TestDrive

Private Sub UnMapDrive( _
        ByVal Drive As String, _
        ByRef Ntwk As Object)
    
    If pLocalFolder Then Exit Sub
    
    On Error Resume Next
    If Drive <> "Failed" Then
        Ntwk.RemoveNetworkDrive Drive
    End If
    On Error GoTo 0
    Set Ntwk = Nothing
End Sub                                          ' UnMapDrive

Private Sub CreatePrimaryKey( _
        ByVal TableName As String, _
        ByVal KeyColumn As Variant)
    
    Set pCnn = New ADODB.Connection
    
    With pCnn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open pDataBasePath
    End With
    
    'create the catalog
    Dim Cat As ADOX.Catalog
    Set Cat = New ADOX.Catalog
    Cat.ActiveConnection = pCnn
    
    Dim Tbl As ADOX.Table
    Set Tbl = Cat.Tables(TableName)
    
    'delete any existing primary keys
    Dim idx As ADOX.Index
    For Each idx In Tbl.Indexes
        If idx.PrimaryKey Then
            Tbl.Indexes.Delete idx.Name
        End If
    Next idx
    
    'create a new primary key
    Set idx = New ADOX.Index
    With idx
        .PrimaryKey = True
        .Name = "PrimaryKey"
        .Unique = True
    End With
    
    'append the column
    idx.Columns.Append KeyColumn
    
    'append the index to the collection
    Tbl.Indexes.Append idx
    Tbl.Indexes.Refresh
    
    'clean up references
    Set pCnn = Nothing
    Set Cat = Nothing
    Set Tbl = Nothing
    Set idx = Nothing
    
End Sub

Private Sub DeleteOldAndCreateNewDatabase()

    Set pNtwk = New WshNetwork
    
    Dim FS As FileSystemObject
    Set FS = New FileSystemObject                ' Microsoft Scripting Runtime
    
    If Mid$(ThisWorkbook.Path, 2, 1) = ":" Then
        pDrv = ThisWorkbook.Path
        pLocalFolder = True
    Else
        pDrv = MapDrive(ThisWorkbook.Path, pNtwk)
        pLocalFolder = False
    End If
    
    pDataBasePath = pDrv & Application.PathSeparator & TARGET_DB
    
    'delete the DB if it already exists
    Dim ErrorNumber As Long
    On Error Resume Next
    Kill pDataBasePath
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

Private Sub CreateOneTable( _
        ByVal Tbl As ListObject, _
        ByVal Cat As ADOX.Catalog)

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
            AccessTable.Columns.Append FieldName, adVarWChar, MaxLength
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
        
    Cat.Tables.Append AccessTable
        
End Sub

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
    
End Function


