Attribute VB_Name = "modDBCode"
Option Explicit
Option Base 0

'-------------------------------MODULE COMMENT------------------------------------
'Created:   02/11/00    By:Ray Raf
'Purpose:   This module contains all of the code for Access Database
'                   functions.
'               At initial creation these functions are designed to work
'               with Access97.  This requires a reference to be created to
'               the Microsoft Access 8.0 Object Library.  In VB6, select Project,
'               References, then find the above library and mark the
'               check box.
'-------------------------------------------------------------------------------------------------

Public Function MakeNewBatchTable(strBatchName As String) As Integer
    '------------------------------------------------------------------------------------------------
    'Created:   02/11/00    By:Ray Raf
    'Purpose:   This function will copy the default batch table to another
    '                table and name the new table "BCH_<batchname>_DAT"
    'Inputs:     strBatchName - the name of the batch entered by the operator.
    'Outputs:   MakeNewBatchTable - function returns 0 if all went well,
    '                the error code for standard errors, and a negative number for
    '               errors specific to this routine.
    '-------------------------------------------------------------------------------------------------
    Dim appAccess As New Access.Application
    
    'select the Results database
    appAccess.OpenCurrentDatabase globResultsDBLocation, False
    'copy table
    appAccess.DoCmd.CopyObject globResultsDBLocation, "BCH_" & strBatchName & "_DAT", acTable, "Blank_Batch"

    appAccess.CloseCurrentDatabase
End Function

Public Sub LoadSetupHeaderToGlobals(rst As Recordset)
    '------------------------------------------------------------------------------------------------
    'Created:   04/05/00    By:Ray Raf
    'Purpose:   This sub will copy the setup info in the current record of
    '                the passed recordset into the CurrentSetup global variable
    'Inputs:     rst - a recordset pulled from the DistrSetup table and set
    '                       to point to the desired record in the set.
    'Rev'd      04/12/00        By:Ray Raf
    'Notes: added the Setup input variable
    '           Setup - variable of type SetupInfo that will be filled with the
    '                       data retreived from the file.  Done this way the routine
    '                       can be used to pick the data out of the recordset and
    '                       put it anywhere, not just CurrentSetup.  It also allows
    '                       picking the data out of both the BatchLog in Results DB
    '                       and DistrSetup in the Setup DB.
    '   Rev'd:  08/15/00    By: Ray Raf
    '   Note:   Changed StartupVoltage to InitV1mACurrent and added
    '               ChargeVoltage field
    '------------------------------------------------------------------------------------------------

    With Setup
        .Name = rst![SetupID]
        .Description = rst![Description]
        .CreateDate = rst![CreationDate]
        .RevDate = rst![RevDate]
        .NumOfCats = rst![NumberOfCategories]
        .MaximumPriority = rst![MaxPriority]
        .DiskType = rst![DiskType]
        .InitV1mACurrent = rst![InitV1mACurrent]
        .ChargeVoltage = rst![ChargeVoltage]
        .ImpulseCurrent = rst![ImpulseCurrent]
        .DischargeVoltage = rst![DischargeVoltage]
        .CatTableName = rst![CatTable]
    End With

End Sub

Function LoadCategories(rst As Recordset) As Integer
    '------------------------------------------------------------------------------------------------
    'Created:   04/05/00    By:Ray Raf
    'Purpose:   This sub will copy the category info in the passed
    '                recordset into the CategoryParam variable array.  It could be
    '                just the Major Cats or the Minor or any combination depending
    '                upon the recordset passed.
    'Inputs:     rst - a recordset pulled from a categories table
    'Rev'd      04/12/00        By:Ray Raf
    'Notes: added the Setup input variable
    '           MC - variable of type CategoryParams that will be filled with the
    '                       data retreived from the file.  Done this way the routine
    '                       can be used to pick the data out of the recordset and
    '                       put it anywhere, not just MajorCategories.  It also allows
    '                       picking the data out of both the Results DB
    '                       and the Setup DB category tables.
    'Outputs:   function returns the number of major categories found
    ' Rev'd:    08/23/00    By: Ray Raf
    '   Note:   deleted the copy of the maxwattsloss to the category arrays
    '------------------------------------------------------------------------------------------------
    
    Dim iNumOfCats As Integer
    Dim iLoop As Integer
    
    iNumOfCats = rst.RecordCount
    ReDim Cats(iNumOfCats - 1)
    rst.MoveFirst
    For iLoop = 0 To iNumOfCats - 1
        With Cats(iLoop)
            .Category = rst![CatLabel]
            .V1mAMin = rst![MinV1mA]
            .V1mAMax = rst![MaxV1mA]
            .V10kAMin = rst![MinV10kA]
            .V10kAMax = rst![MaxV10kA]
            .RatioMin = rst![MinRatio]
            .RatioMax = rst![MaxRatio]
            .MinWattsLossVoltage = rst![MinWLVoltage]
            .Priority = rst![Priority]
        End With
        rst.MoveNext
    Next iLoop
    
    LoadCategories = iNumOfCats
    
End Function


Function CreateNewBatch(strBatchName As String) As Integer
    '------------------------------------------------------------------------------------------------
    'Created:   04/10/00    By:Ray Raf
    'Purpose:   This sub will copy the setup info from the setup database
    '                to the Results database.  It will store the header info in the
    '                BatchLog table, copy the categories table to CAT_BCHxxxxx
    '                where xxxxx is the batch id, and create the
    '   Return: CreateNewBatch - 0 = success, 1=batch already exists.
    '   Rev'd:  06/15/00    By:Ray Raf
    '   Note:   Changed this to a function to return an error code if the
    '              passed batch number already exists.
    '   Rev'd   08/15/00    By: Ray Raf
    '   Note:   Changed startupvoltage field to InitV1maCurrent to match DB
    '              Also added save of ChargeVoltage to batchlog table
    '   Rev'd:  08/23/00    By: Ray Raf
    '   Note:   Added ImpulseCurrent and DischargeVoltage
    '------------------------------------------------------------------------------------------------

    Dim rstSet As Recordset
    Dim rstRes As Recordset
    Dim iLoop As Integer
    Dim fld As Field
    Dim wrkJet As Workspace
    Dim dbSetup As Database
    Dim dbResults As Database
    Dim appAccess As New Access.Application
    
    CreateNewBatch = 2      'set to return unknown error if exit prematurely
    
    'set tabledef object to point to 'DistrSetup" table in DistrSetup database
    Set wrkJet = CreateWorkspace("", "admin", "", dbUseJet)
    Set dbSetup = wrkJet.OpenDatabase(globSetupDBLocation)
    Set dbResults = wrkJet.OpenDatabase(globResultsDBLocation)
    
    'Set rstSet = dbSetup.OpenRecordset("DistrSetup")
    Set rstRes = dbResults.OpenRecordset("BatchLog", dbOpenDynaset)
    rstRes.FindFirst "[BatchID] = '" & strBatchName & "'"
    If rstRes.NoMatch = False Then
        'the batch name already exists
        MsgBox "The batch name you have selected already exists. " & vbCrLf & _
                     "Please select another name."
        CreateNewBatch = 1  'flag that batch number already exists
        GoTo CloseAndExit
    End If
    
    'put basic data in batch log table
    rstRes.AddNew
    With CurrentSetup
        rstRes![BatchID] = strBatchName
        rstRes![StartDate] = Date
        rstRes![StartTime] = Time
        rstRes![SetupID] = .Name
        rstRes![Description] = .Description
        rstRes![CreationDate] = .CreateDate
        rstRes![RevDate] = .RevDate
        rstRes![NumberOfCategories] = .NumOfCats
        rstRes![MaxPriority] = .MaximumPriority
        rstRes![DiskType] = .DiskType
        rstRes![InitV1mACurrent] = .InitV1mACurrent
        rstRes![ChargeVoltage] = .ChargeVoltage
        rstRes![ImpulseCurrent] = .ImpulseCurrent
        rstRes![DischargeVoltage] = .DischargeVoltage
        rstRes![CatTable] = "BCH_" & strBatchName & "_CAT"
        rstRes![BatchTableName] = "BCH_" & strBatchName & "_DAT"
        rstRes![QDiskTableName] = "BCH_" & strBatchName & "_Q"
    End With
    rstRes.Update
    
    'copy categories table from setup db to results db
    appAccess.OpenCurrentDatabase globSetupDBLocation, False
    'copy table
    appAccess.DoCmd.CopyObject globResultsDBLocation, "BCH_" & strBatchName & "_CAT", acTable, _
                    CurrentSetup.CatTableName
    appAccess.CloseCurrentDatabase
    
    
    'make a new batch data table in the Results data base
    'select the Results database
    appAccess.OpenCurrentDatabase globResultsDBLocation, False
    'copy table
    appAccess.DoCmd.CopyObject globResultsDBLocation, "BCH_" & strBatchName & "_DAT", acTable, _
                    "Batch_Data_Blank"
    appAccess.CloseCurrentDatabase

    'make a new Q disk data table in the Results data base
    'select the Results database
    appAccess.OpenCurrentDatabase globResultsDBLocation, False
    'copy table
    appAccess.DoCmd.CopyObject globResultsDBLocation, "BCH_" & globBatchName & "_Q", acTable, _
                    "Q_Disks_Blank"
    appAccess.CloseCurrentDatabase
    
    CreateNewBatch = 0
CloseAndExit:
    Set rstSet = Nothing
    Set rstRes = Nothing
    Set dbSetup = Nothing
    Set dbResults = Nothing
    wrkJet.Close
    Set appAccess = Nothing
End Function

Public Sub SaveQDiskData()
'------------------------------------------------------------------------------------------------
'Created:   05/09/00    By:Ray Raf
'Purpose:  Writes Q disk data to the results DB for the current batch
'-------------------------------------------------------------------------------------------------
    Dim rstBatch As Recordset
    Dim rstRes As Recordset
    Dim iLoop As Integer
    Dim fld As Field
    Dim wrkJet As Workspace
    Dim dbResults As Database
    Dim appAccess As New Access.Application
    Dim iNumRecords As Integer
    
    'set db object to point to results database
    Set wrkJet = CreateWorkspace("", "admin", "", dbUseJet)
    Set dbResults = wrkJet.OpenDatabase(globResultsDBLocation)
        
    Set rstBatch = dbResults.OpenRecordset("BatchLog", dbOpenDynaset)
    rstBatch.FindFirst ("[BatchID] = '" & globBatchName & "'")
    If rstBatch.NoMatch = True Then
        ErrorLogger "SaveQDiskData", "Error finding batch name '" & globBatchName & "' in database"
        MsgBox "Error accessing database for Q disk Data. " & vbCrLf & _
                    "Contact your Supervisor!"
        GoTo CloseAndExit
    End If
        
    'open q disk table for this batch
    On Error Resume Next
    Set rstRes = dbResults.OpenRecordset("BCH_" & globBatchName & "_Q", dbOpenDynaset)
    If Err.Number <> 0 Then
        Err.Clear
        'make a new Q disk data table in the Results data base
        'select the Results database
        appAccess.OpenCurrentDatabase globResultsDBLocation, False
        'copy table
        appAccess.DoCmd.CopyObject globResultsDBLocation, "BCH_" & globBatchName & "_Q", acTable, _
                        "Q_Disks_Blank"
        appAccess.CloseCurrentDatabase
        Set rstRes = dbResults.OpenRecordset("BCH_" & globBatchName & "_Q", dbOpenDynaset)
    End If
    On Error GoTo 0
    
    'have an open and valid table for q disks
    On Error Resume Next
    rstRes.MoveLast             'fully populate recordset to get accurate count of records
    If Err.Number <> 0 Then
        Select Case Err.Number
            Case 3021       'no current record means no records exits, which is ok
                iNumRecords = 0
            Case Else
                'don't know what the problem might be, message and exit
                ErrorLogger "SaveQDiskData", " Error populating DB " & Err.Number & " " & Err.Description
                MsgBox "Unknown error trying to save Q disk Data"
                GoTo CloseAndExit
        End Select
    Else  'recordset is populated, get number of records
        iNumRecords = rstRes.RecordCount
    End If
    On Error GoTo 0
    
    If iNumRecords = 0 Then
        'have a fresh, clean table - just put the data in
        With rstRes
            For iLoop = 1 To globNumQDisks
                .AddNew
                ![Id] = iLoop
                ![V1mA] = LabDisks(iLoop).V1mA
                ![Watts] = LabDisks(iLoop).Watts
                ![V10kA] = LabDisks(iLoop).V10kA
                .Update
            Next iLoop
        End With
    Else
        'we have to check to see of the records already exist before adding
        For iLoop = 1 To globNumQDisks
            rstRes.FindFirst "[ID] = " & iLoop
            If rstRes.NoMatch = True Then
                'record doesn't exist yet, have to add
                rstRes.AddNew
                rstRes![Id] = iLoop
            Else
                'just modify existing record
                rstRes.Edit
            End If
            'put data in
            rstRes![V1mA] = LabDisks(iLoop).V1mA
            rstRes![Watts] = LabDisks(iLoop).Watts
            rstRes![V10kA] = LabDisks(iLoop).V10kA
            rstRes.Update
        Next iLoop
        
        'check number of records again
        rstRes.MoveLast
        iNumRecords = rstRes.RecordCount
        If iNumRecords > globNumQDisks Then
            Do
                'delete extra records
                rstRes.FindLast "[ID] > " & globNumQDisks
                If rstRes.NoMatch = False Then
                    rstRes.Delete
                End If
            Loop While rstRes.NoMatch = False
            'set number of records to whats left in table
            
        End If
    End If
    
CloseAndExit:
    Set appAccess = Nothing
    Set rstRes = Nothing
    Set rstBatch = Nothing
    Set dbResults = Nothing
    wrkJet.Close
End Sub

Public Sub LoadQDiskData()
'------------------------------------------------------------------------------------------------
'Created:   05/09/00    By:Ray Raf
'Purpose:  Read Q disk data from the results DB for the current batch
'-------------------------------------------------------------------------------------------------
    Dim rstBatch As Recordset
    Dim rstRes As Recordset
    Dim iLoop As Integer
    Dim wrkJet As Workspace
    Dim dbResults As Database
    Dim appAccess As New Access.Application
    Dim iNumRecords As Integer
    Dim strMsg As String
    Dim strResponse As String
    Dim iDiskNum As Integer
    
    'set db object to point to results database
    Set wrkJet = CreateWorkspace("", "admin", "", dbUseJet)
    Set dbResults = wrkJet.OpenDatabase(globResultsDBLocation)
        
    Set rstBatch = dbResults.OpenRecordset("BatchLog", dbOpenDynaset)
    rstBatch.FindFirst ("[BatchID] = '" & globBatchName & "'")
    If rstBatch.NoMatch = True Then
        ErrorLogger "LoadQDiskData", "Error finding batch name '" & globBatchName & "' in database"
        MsgBox "Error accessing database for Q disk Data. " & vbCrLf & _
                    "Contact your Supervisor!"
        GoTo CloseAndExit
    End If
        
    'open q disk table for this batch
    On Error Resume Next
    Set rstRes = dbResults.OpenRecordset("BCH_" & globBatchName & "_Q", dbOpenDynaset)
    If Err.Number <> 0 Then
        'its ok not to find data, but user will have to enter it
        Err.Clear
        globNeedQDiskData = True
        MsgBox "Cannot access existing Q disk data.  You must " & vbCrLf & _
                     "re-enter the data to proceed with AutoCal."
        GoTo CloseAndExit
    End If
    
    'Get number of disks by checking number of records
    rstRes.MoveLast
    If Err.Number <> 0 Then   ' need to have user enter data
        Err.Clear
        globNeedQDiskData = True
        GoTo CloseAndExit
    End If
                
    globNumQDisks = rstRes.RecordCount
    
    rstRes.MoveFirst
    'get individual disk data
    For iLoop = 1 To globNumQDisks
        iDiskNum = rstRes![Id]
        With LabDisks(iDiskNum)
            .V1mA = rstRes![V1mA]
            .Watts = rstRes![Watts]
            .V10kA = rstRes![V10kA]
        End With
        rstRes.MoveNext
    Next iLoop
    globNeedQDiskData = False
    
CloseAndExit:
    Set appAccess = Nothing
    Set rstRes = Nothing
    Set rstBatch = Nothing
    Set dbResults = Nothing
    wrkJet.Close
    
End Sub

Public Sub SaveDisk()
'------------------------------------------------------------------------------------------------
'Created:   06/21/00    By:Ray Raf
'Purpose:  Save the passed disk data to the results DB
'Rev'd:      06/26/00       By: Ray Raf
'Note:      Changed routine to save actual category number instead of
'              index number.  The index number can change if the Categories
'             table for this batch gets sorted different for some reason.
'Rev'd      08/15/00    By: Ray Raf
'Note:      Added checks using bMajorCatFound and bMinorCatFound
'              to prevent subscript errors when accessing Major and
'              Minor category arrays when the category is an error category.
'-------------------------------------------------------------------------------------------------
    Dim wrkJet As Workspace
    Dim dbResults As Database
    Dim rstRes As Recordset
    
    'set db object to point to results database
    Set wrkJet = CreateWorkspace("", "admin", "", dbUseJet)
    Set dbResults = wrkJet.OpenDatabase(globResultsDBLocation)
        
    'open data table for this batch
    On Error Resume Next
    Set rstRes = dbResults.OpenRecordset("BCH_" & globBatchName & "_DAT", dbOpenDynaset)

    rstRes.AddNew
    rstRes![Disk_ID] = Disk.DiskID
    rstRes![V1mA] = (Int(Disk.V1mA * 100)) / 100
    rstRes![V10kA] = (Int(Disk.V10kA * 100)) / 100
    rstRes![Ratio] = Disk.Ratio
    rstRes![WattsLoss] = Disk.WattsLoss
    If Disk.bMajorCatFound = True Then    '8/15/00
        rstRes![Category] = MajorCategories(Disk.MajorCat).Category
    Else
        rstRes![Category] = Disk.MajorCat
    End If
    If Disk.bMinorCatFound = True Then      '8/15/00
        rstRes![SubCategory] = MinorCategories(Disk.MinorCat).Category
    Else
        rstRes![SubCategory] = 0
    End If
    rstRes.Update
    
CloseAndExit:
    Set rstRes = Nothing
    Set dbResults = Nothing
    wrkJet.Close
End Sub

Public Function DeleteBatch(strBatchName As String) As Integer
    '------------------------------------------------------------------------------------------------
    'Created:   08/17/00    By:Ray Raf
    'Purpose:  to delete the passed batch name from the local results DB
    'Returns:   0 if successful, -1 if batch didn't exist, -2 other probelms
    '-------------------------------------------------------------------------------------------------

    Dim wrkJet As Workspace
    Dim dbResults As Database
    Dim rstRes As Recordset
    Dim strQName As String, strCatName As String, strDatName As String
    Dim appAccess As New Access.Application
    
    'set db object to point to results database
    Set wrkJet = CreateWorkspace("", "admin", "", dbUseJet)
    Set dbResults = wrkJet.OpenDatabase(globResultsDBLocation)
        
    'open data table for this batch
    On Error GoTo OpenFailure
    Set rstRes = dbResults.OpenRecordset("BatchLog", dbOpenDynaset)
    
    rstRes.FindFirst "[BatchID] = '" & strBatchName & "'"
    If rstRes.NoMatch = True Then
        'let calling routine know that batch was not found in batch log
        DeleteBatch = -1
        GoTo CloseAndExit
    End If
        
    strQName = rstRes![QDiskTableName]
    strCatName = rstRes![CatTable]
    strDatName = rstRes![BatchTableName]
        
    On Error Resume Next    'if a table doesn't exist to delete, the DeleteObject
                                        'method will throw an error.  Since we wanted the
                                        'table gone anyhow, we will just go to the next table.
    'select the Results database
    appAccess.OpenCurrentDatabase globResultsDBLocation, False
    'delete tables
    appAccess.DoCmd.DeleteObject acTable, strDatName
    appAccess.DoCmd.DeleteObject acTable, strCatName
    appAccess.DoCmd.DeleteObject acTable, strQName

    appAccess.CloseCurrentDatabase
    
    On Error GoTo 0
    'delete batch data record from BatchLog table
    rstRes.Delete
    
    DeleteBatch = 0
CloseAndExit:
    Set rstRes = Nothing
    Set dbResults = Nothing
    wrkJet.Close
    
    Err.Clear
    Exit Function
    
OpenFailure:
    DeleteBatch = -2
    GoTo CloseAndExit
    
End Function

Public Function AddToBatchNote(strComment) As Integer
    '------------------------------------------------------------------------------------------------
    'Created:   08/17/00    By:Ray Raf
    'Purpose:  to add the passed string to the existing batch note for the
    '               current batch.
    'Returns:   0 if successful, -1 if batch didn't exist, -2 other probelms
    '-------------------------------------------------------------------------------------------------
    Dim wrkJet As Workspace
    Dim dbResults As Database
    Dim rstRes As Recordset
    Dim strNewComment As String
    
    'set db object to point to results database
    Set wrkJet = CreateWorkspace("", "admin", "", dbUseJet)
    Set dbResults = wrkJet.OpenDatabase(globResultsDBLocation)
        
    'open data table for this batch
    Set rstRes = dbResults.OpenRecordset("BatchLog", dbOpenDynaset)
    
    rstRes.FindFirst "[BatchID] = '" & globBatchName & "'"
    If rstRes.NoMatch = True Then
        'let calling routine know that batch was not found in batch log
        AddToBatchNote = -1
        GoTo CloseAndExit
    End If

    strNewComment = Format(Now(), "General Date") & Space(3) & strComment
    rstRes.Edit
    rstRes![BatchComment] = rstRes![BatchComment] & vbCrLf & strNewComment
    rstRes.Update
    
    AddToBatchNote = 0
CloseAndExit:
    Set rstRes = Nothing
    Set dbResults = Nothing
    wrkJet.Close
    
    Err.Clear
    
End Function
