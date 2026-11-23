Attribute VB_Name = "AZ298BCatListing"
'AZ298B_CatalogListing
Public Function TestCatalogNumber(CatNo As String) As Boolean

'On Error GoTo ErrorOccurred

    Dim rstCatNos As New ADODB.Recordset
    Dim SQL As String
    

    SQL = "SELECT * FROM AZ298B_CatalogListing_Rev0 WHERE CatalogNumber = '" & CatNo & "'"
    
    rstCatNos.Open SQL, dbconSetupData, adOpenDynamic, adLockBatchOptimistic
    
    If rstCatNos.EOF = False Then
        TestCatalogNumber = True
    Else
        TestCatalogNumber = False
    End If
    
    If rstCatNos.State = 1 Then
        rstCatNos.Close
    End If
    
    Exit Function
    
ErrorOccurred:
    TestCatalogNumber = False
End Function
