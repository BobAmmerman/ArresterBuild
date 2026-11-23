Attribute VB_Name = "modTest4FACTSDisk"
Option Explicit

Public Function NotFACTS_Disk(DiskBatch As String, DiskNum As Integer) As Boolean

Dim rstDiskDataOrg As New ADODB.Recordset
Dim SQL As String
    If rstDiskDataOrg.State = 1 Then
        rstDiskDataOrg.Close
    End If
    
    SQL = "SELECT * FROM ConsumedDisks"
    SQL = SQL & " WHERE Batch_Number = " & DiskBatch
    SQL = SQL & " AND Disk_Number = " & DiskNum
    
    rstDiskDataOrg.Open SQL, dbconResultsData, adOpenDynamic, adLockOptimistic
    
    If rstDiskDataOrg.EOF = False Then
    ' disk found
        NotFACTS_Disk = False
        MsgBox "Disk " & DiskBatch & "-" & DiskNum & " was previously used in a FACTS tube and can NOT be used in standard production." & _
                "Disk should be returned to MOV for retesting."
    Else
    ' disk not found
        NotFACTS_Disk = True
    
    End If
    
    If rstDiskDataOrg.State = 1 Then
        rstDiskDataOrg.Close
    End If
    
End Function
