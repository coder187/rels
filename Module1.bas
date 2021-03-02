Attribute VB_Name = "Module1"
'Jonathan Kelly March 2021
'jkelly@ksd.ie
'
'herein routines to :   read relationships
'                       save to table
'                       create from table
'                       used to restore deleted relationships to stock.mdb 25/3/21
'                       use with caution
'
'REF: https://stackoverflow.com/questions/354651/importing-exporting-relationships		
'Patrick Cuff
'
Option Compare Database
Option Explicit

Type table_details
     ReCount As Long
     FieldCount As Long
     tblsSize As Long
End Type

Function SaveRelationshipsToTable(src_database_file As String)
'SaveRelationshipsToTable "c:\users\joan\last_known_good_copy.mdb"
If ZapTargetTable = False Then
    Debug.Print "zap failed -- table locked? see logs"
    Exit Function
End If
    
Dim Source_DB_File_Path As String
Source_DB_File_Path = src_database_file

Dim Source_DB As Database
Set Source_DB = Workspaces(0).OpenDatabase(Source_DB_File_Path, True, True)

Dim rel As Relation
Dim fld As Field
Dim cnt As Integer
Dim tot As Integer
Dim r As Long
Dim rr As Boolean
Dim fld_order As Integer
Dim tbl_details As table_details

cnt = 1
tot = Source_DB.Relations.Count
r = 0
fld_order = 1

    For Each rel In Source_DB.Relations
    
        Debug.Print "rel " & cnt & " of " & tot & " : " & Round(cnt / tot * 100) & "%"
        
        With rel
            tbl_details = Get_TableDetails(Source_DB, .Table)
            
            r = insert_reltable("stock", .Name, .Table, .ForeignTable, .Attributes, tbl_details.ReCount, tbl_details.FieldCount, tbl_details.tblsSize)
            
            If r > 0 Then 'successful insert above
                fld_order = 1 'field relationship order - maybe not needed
                For Each fld In .Fields
                    rr = insert_reltable_fields(r, fld.Name, fld.ForeignName, fld_order)
                    fld_order = fld_order + 1
                Next
            End If
                            
            If rr = False Then 'field insert failed
                Debug.Print "FAILED"
                Debug.Print .Table & " " & .ForeignTable & " " & .Name
            Else    'all good
                Debug.Print "inserted"
            End If
            
            DoEvents
            cnt = cnt + 1
        End With
    Next


End Function

Function insert_reltable(sys As String, rel_name As String, tbl_name As String, f_tbl_name As String, attrib As Long, RecordCnt As Long, FieldCnt As Long, TblSize As Long) As Long
On Error GoTo insert_reltable_err
    Dim mydb As DAO.Database
    Dim q As DAO.QueryDef
    
    Set mydb = CurrentDb
    Set q = mydb.QueryDefs("Rels_Append")
    
    q.Parameters("system").Value = sys
    q.Parameters("relationship name").Value = rel_name
    q.Parameters("table name").Value = tbl_name
    q.Parameters("foreign table name").Value = f_tbl_name
    q.Parameters("attrib").Value = attrib
    q.Parameters("record count").Value = RecordCnt
    q.Parameters("field count").Value = FieldCnt
    q.Parameters("table size").Value = TblSize
    

    q.Execute dbFailOnError
    q.Close
    
    Set q = Nothing
    
    Dim r As DAO.Recordset
    Set r = mydb.OpenRecordset("select @@identity")
    r.MoveFirst
    
    insert_reltable = r(0)
    
    Set r = Nothing
    Set mydb = Nothing

insert_reltable_exit:
    Exit Function
insert_reltable_err:
    DoEvents
    Debug.Print "insert_reltable:" & err.Description & ":" & Now()
    insert_reltable = -1
    Resume insert_reltable_exit

End Function


Function insert_reltable_fields(rel As Long, primaryField As String, ForeignField As String, FieldOrder As Integer) As Boolean
On Error GoTo insert_reltable_fields_err
    Dim mydb As DAO.Database
    Dim q As DAO.QueryDef
    
    Set mydb = CurrentDb
    Set q = mydb.QueryDefs("Fields_Append")
    
'INSERT INTO rel_fields ( id, field_name, field_foreign_name, rel_order )
'SELECT [rel] AS Expr2, [primary field] AS Expr1, [foreign field] AS Expr3, [rel order] AS Expr4;
   
    q.Parameters("rel").Value = rel
    q.Parameters("primary field").Value = primaryField
    q.Parameters("foreign field").Value = ForeignField
    q.Parameters("rel order").Value = FieldOrder

    q.Execute dbFailOnError
    q.Close
    
    Set q = Nothing
    Set mydb = Nothing
    
    insert_reltable_fields = True
    
insert_reltable_fields_exit:
    Exit Function
insert_reltable_fields_err:
    DoEvents
    Debug.Print "insert_reltable_fields:" & err.Description & ":" & Now()
    insert_reltable_fields = False
    Resume insert_reltable_fields_exit

End Function

Function Get_TableDetails(db As Database, tbl As String) As table_details

    Dim r As DAO.Recordset
    Dim rec_count As Long
    Dim field_count As Long
    Dim my_table_details As table_details
    
    Set r = db.OpenRecordset(tbl, dbOpenSnapshot)
    
    my_table_details.FieldCount = r.Fields.Count
    
    If Not r.EOF Then
        r.MoveLast
        r.MoveFirst
        my_table_details.ReCount = r.RecordCount
    Else
        my_table_details.ReCount = 0
    End If
    my_table_details.tblsSize = 0
    Get_TableDetails = my_table_details
Get_RecCount_exit:
    Exit Function
    
Get_RecCount_err:
    Debug.Print "Get_RecCount:" & err.Description & ":" & Now()
    my_table_details.tblsSize = -1
    my_table_details.ReCount = -1
    my_table_details.FieldCount = -1
    Get_TableDetails = my_table_details
    Resume Get_RecCount_exit
End Function

Function ZapTargetTable() As Boolean
On Error GoTo ZapTargetTable_Err

ZapTargetTable = False
CurrentDb.Execute ("delete * from rels;")
ZapTargetTable = True

ZapTargetTable_Exit:
    Exit Function
ZapTargetTable_Err:
    ZapTargetTable = False
    Resume ZapTargetTable_Exit
End Function
Function ZapLogTable() As Boolean
On Error GoTo ZapLogTable_Err

ZapLogTable = False
CurrentDb.Execute ("delete * from error_log;")
ZapLogTable = True

ZapLogTable_Exit:
    Exit Function
ZapLogTable_Err:
    ZapLogTable = False
    Resume ZapLogTable_Exit
End Function


Function CreateRelationships(target_database_file As String)
On Error GoTo CreateRelationships_Err
'CreateRelationships "c:\users\joan\last_backup.mdb"

If Not ZapLogTable Then
    Debug.Print "zap logs failed -- table locked?"
    Exit Function
End If

If CreateFileCopy(target_database_file, target_database_file & "." & CDbl(Now)) = False Then
    Exit Function
End If

Dim Target_DB_File_Path As String
Target_DB_File_Path = target_database_file

Dim Target_DB As Database
Set Target_DB = Workspaces(0).OpenDatabase(Target_DB_File_Path, True, False)

Dim mydb As Database
Dim relationship_list As Recordset
Set mydb = CurrentDb()
Set relationship_list = mydb.OpenRecordset("listrels", dbOpenDynaset)

Dim tot As Integer
Dim cnt As Integer
cnt = 1

Dim rel As Relation

With relationship_list
    If Not .EOF Then
        .MoveLast
        .MoveFirst
        tot = .RecordCount
        
        Do While Not .EOF
            DoEvents
            Debug.Print "creating rel " & cnt & " of " & tot & " : " & Round(cnt / tot * 100) & "%"
            Set rel = Target_DB.CreateRelation(Name:=.Fields("rel_name"), Table:=.Fields("table_name"), ForeignTable:=.Fields("foreign_table"), Attributes:=.Fields("table_attributes"))
            rel.Fields.Append rel.CreateField(.Fields("field_name"))
            rel.Fields(.Fields("field_name")).ForeignName = .Fields("field_foreign_name")
            
            
            Target_DB.Relations.Append rel
        
            Update_Rel_Applied .Fields("rol")
            
            cnt = cnt + 1
        .MoveNext
        Loop
    Else
        Debug.Print "no records to read"
    End If
End With

CreateRelationships_Exit:
    Exit Function
    
CreateRelationships_Err:
    DoEvents
    Debug.Print "CreateRelationships_Err:" & err.Number & err.Description
    
    Select Case err.Number
    Case 3012
        Debug.Print "..skipping..."
        Resume Next
    Case 3201
        If Not (Log_Append(err.Number, err.Description)) Then
            Resume CreateRelationships_Exit
        Else
            Resume Next
        End If
    Case Else
        Resume CreateRelationships_Exit
    End Select

End Function

Function Update_Rel_Applied(id As Long) As Boolean
On Error GoTo Update_Rel_Applied_Err
    CurrentDb.Execute "update rels set applied_to_target = " & CDbl(Now()) & ";"
    Update_Rel_Applied = True
Update_Rel_Applied_Exit:
    Exit Function
Update_Rel_Applied_Err:
    Debug.Print "Update_Rel_Applied:" & err.Description & ":" & Now
    Update_Rel_Applied = False
    Resume Update_Rel_Applied_Exit
End Function

Function CreateFileCopy(src As String, tar As String) As Boolean
On Error GoTo CreateFileCopy_Err
    If Len(Dir(tar)) <> 0 Then
        DoEvents
        Debug.Print "File Exists! Choose New File Name and try again."
        CreateFileCopy = False
        Exit Function
    End If
        
    FileCopy src, tar
    DoEvents
    Do While Len(Dir(tar)) = 0
        'try and wait for the filesystem to compelte the copy.
    Loop
    CreateFileCopy = True
CreateFileCopy_Exit:
    Exit Function
CreateFileCopy_Err:
    Debug.Print "CreateFileCopy:" & err.Description & ":" & Now()
    CreateFileCopy = False
    Resume CreateFileCopy_Exit
End Function

Function Log_Append(id As Integer, errs As String)
'Function Log_Append(ByRef thisError As ErrObject) 'cant get this to work. err object being reset i think
On Error GoTo Log_Append_Err
    Dim mydb As DAO.Database
    Dim q As DAO.QueryDef
    
    Set mydb = CurrentDb
    'Set q = mydb.QueryDefs("Log_Append")
    
    'q.Parameters("id").Value = id
    'q.Parameters("err").Value = Left(errs, 255) 'cant get this to work... possibly needs esc char
    
    mydb.Execute "INSERT INTO error_log ( num, [desc] ) values(" & id & "," & Chr(34) & errs & Chr(34) & ");"
    
    Log_Append = True
Log_Append_Exit:
    Exit Function
Log_Append_Err:
    Debug.Print "Log_Append_Err:" & err.Description & ":" & Now() & "* ATTENTION : UNLOGGED ERROR *****************"
    Log_Append = False
    Resume Log_Append_Exit
    
End Function