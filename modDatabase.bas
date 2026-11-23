Attribute VB_Name = "modDatabase"
Option Explicit

'--main database connection
'Global adosetupdata As ADODB.Connection
'--adoRS = temp recordset, adoSerials = temp recordset for retrieving sn's
Global adoRS, adoSerials As ADODB.Recordset
Global dbconSetupData As New ADODB.Connection
Global dbconResultsData As New ADODB.Connection

Public Function ConnectDatabases() As Boolean
    '------------------------------------------------------------------------------------
    '   Created:    03/21/01    By:Ray Raf
    '   Purpose:    Establish connection to DistrTester databases:
    '   Returns:    True if connection was established, false otherwise
    '------------------------------------------------------------------------------------
    
    Dim UserName As String
    Dim Password As String
    Dim ServerName As String
    Dim sqlDatabase As String
    Dim ConnectionString As String
    Set adoRS = New ADODB.Recordset
    Set adoSerials = New ADODB.Recordset
    
    On Error GoTo ErrorHandler
    'Get Setup database connection parameters
    ServerName = GetSet(AppName:=App.Title, Section:="DB", _
            Key:="SetupReadOnly_Server", Default:="")
    sqlDatabase = GetSet(AppName:=App.Title, Section:="DB", _
            Key:="SetupReadOnly_Database", Default:="")
    UserName = GetSet(AppName:=App.Title, Section:="DB", _
            Key:="SetupReadOnly_UID", Default:="")
    Password = GetSet(AppName:=App.Title, Section:="DB", _
            Key:="SetupReadOnly_PWD", Default:="")

        'use SQL connection
        ConnectionString = "driver={SQL Server};" & _
            "server=" & ServerName & ";" & _
            "uid=" & UserName & ";" & _
            "pwd=" & Password & ";" & _
            "database=" & sqlDatabase & ";"
        'make sure connection is closed
        If dbconSetupData.State = 1 Then
            dbconSetupData.Close
        End If
        
        With dbconSetupData
            .ConnectionString = ConnectionString
            .Open
        End With
    
        
        
    'Get Results database connection parameters
    ServerName = GetSet(AppName:=App.Title, Section:="DB", _
            Key:="Results_Server", Default:="")
    sqlDatabase = GetSet(AppName:=App.Title, Section:="DB", _
            Key:="Results_Database", Default:="")
    UserName = GetSet(AppName:=App.Title, Section:="DB", _
            Key:="Results_UID", Default:="")
    Password = GetSet(AppName:=App.Title, Section:="DB", _
            Key:="Results_PWD", Default:="")

    ConnectionString = "driver={SQL Server};" & _
        "server=" & ServerName & ";" & _
        "uid=" & UserName & ";" & _
        "pwd=" & Password & ";" & _
        "database=" & sqlDatabase & ";"
    If dbconResultsData.State = 1 Then
        dbconResultsData.Close
    End If

    With dbconResultsData
        .ConnectionString = ConnectionString
        .Open
    End With

    ConnectDatabases = True
    Exit Function
ErrorHandler:
    If Err.Number = -2147467259 Then
        'can't connect to database server
        MsgBox "Cannot connect to Database server!"
        frmOptions.Show vbModal
    Else
    End If
    Err.Clear
    ConnectDatabases = False
End Function



''----------------------------------------------------------------------------------------
''--this will soon be part of your already existing ConnectDatabases routine
''----------------------------------------------------------------------------------------
'Public Sub initDatabase()
'    Set adosetupdata = New ADODB.Connection
'    Set adoRS = New ADODB.Recordset
'    Set adoSerials = New ADODB.Recordset
'    adosetupdata.ConnectionString = "uid=sa;pwd=aminteg;driver={SQL Server}; server=192.168.0.24;database=Cooper;dsn="
'    adosetupdata.Open
'End Sub
'
'------------------------------------------------------------------------------------------------
'Created:   04/09/03    By:Jeff Pierce
'Purpose:   To see if it a work order is in the database, if so populate the globCurrentWO's attributes
'Inputs:    strWO_Number - the name of the work order to look up
'Outputs:   returns true if work order in database, returns false if work order is not
'-------------------------------------------------------------------------------------------------
Public Function adoWorkOrderInDatabase(strWO_Number)
    Dim X As Integer
    
    'check for WO in database
    Set adoRS = dbconResultsData.Execute("SELECT * FROM WO_Info WHERE WO_Number='" & strWO_Number & "' ;")
    If adoRS.EOF Then
        'WO not found in data base
        adoWorkOrderInDatabase = False
    Else
        'WO found, set glob variables for later use.
        adoWorkOrderInDatabase = True
        With globCurrentWO
            .strWO = Trim(adoRS.Fields("WO_Number"))
            .strCatNumber = Trim(adoRS.Fields("CatalogNumber"))
            
            .strType = getType(.strCatNumber)
            .strBuildPath = getBuildPath(.strCatNumber, .strType)
            
            '--retrieve the number of serials
            Set adoSerials = dbconSetupData.Execute("SELECT COUNT(SerialNumber) as NumSerials FROM WO_SerialNumbers WHERE WO_Number='" & strWO_Number & "' ;")
            Dim cntSerials As Integer
            cntSerials = Trim(adoSerials!NumSerials)
            
            .numSN = cntSerials
            
        End With
    End If
    
    'close the open recordsets
    If adoRS.State = 1 Then
        adoRS.Close
    End If
    
    If adoSerials.State > 0 Then
        adoSerials.Close
    End If
    
End Function

'------------------------------------------------------------------------------------------------
'Created:   04/8/03    By:Jeff Pierce
'Purpose:   To see if there are any serial numbers for a work order as in the BuiltUnits table
'Inputs:    strWO_Number - work order number to find serial numbers
'Outputs:   returns true if they're SN's in the BuiltUnits table, false if not
'-------------------------------------------------------------------------------------------------
Public Function adoCheckForExistingSNsByWO(strWO_Number) As Boolean
    Set adoSerials = dbconResultsData.Execute("SELECT SerialNumber FROM BuiltUnits WHERE " & _
        "EXISTS (SELECT * FROM WO_SerialNumbers WHERE WO_Number='" & strWO_Number & "' AND " & _
                "BuiltUnits.SerialNumber=WO_SerialNumbers.SerialNumber) ;")
    If adoSerials.EOF Then
        adoCheckForExistingSNsByWO = False
    Else
        adoCheckForExistingSNsByWO = True
    End If
    If adoSerials.State = 1 Then
        adoSerials.Close
    End If
End Function

Public Function adoCheckForExistingSNsByList(strSN() As String) As Boolean
'------------------------------------------------------------------------------------------------
'Created:   05/1/03    By:Ray Raf
'Purpose:   To see if any serial numbers in the list already exist
'                in the BuiltUnits table
'Inputs:    strWO_Number - work order number to find serial numbers
'Outputs:   returns true if the SN's are in the BuiltUnits table, false if not
'-------------------------------------------------------------------------------------------------
    Dim rstSNs As New ADODB.Recordset
    Dim SQL As String, SQLWhere As String
    Dim iLoop As Integer
    Dim strMsg As String
    Dim bFoundInBuilt As Boolean, bFoundInWO As Boolean
    
    SQL = "SELECT * FROM BuiltUnits "
    SQLWhere = "WHERE SerialNumber = '" & strSN(0) & "' "
    If UBound(strSN) > 0 Then
        'more than one serial number in the list.
        For iLoop = 1 To UBound(strSN)
            'add each serial number to the list
            SQLWhere = SQLWhere & "OR SerialNumber = '" & strSN(iLoop) & "' "
        Next iLoop
    End If
    
    SQL = "SELECT * FROM BuiltUnits " & SQLWhere
    rstSNs.Open SQL, dbconResultsData, adOpenDynamic, adLockBatchOptimistic
    If rstSNs.EOF Then
        bFoundInBuilt = False
    Else
        bFoundInBuilt = True
        'show users which SNs already exist
        strMsg = "The following serial Numbers already exist in the Built Units table:" & vbCrLf & vbCrLf
        Do While rstSNs.EOF = False
            strMsg = strMsg & vbCrLf & rstSNs!SerialNumber & ", Unit " & Chr(rstSNs!Unit + 64)
            rstSNs.MoveNext
        Loop
        MsgBox strMsg
        ErrorLogger "CheckForExistingSNs", strMsg
    End If
    
    If rstSNs.State = 1 Then
        rstSNs.Close
    End If

    SQL = "SELECT * FROM WO_SerialNumbers " & SQLWhere
    rstSNs.Open SQL, dbconResultsData, adOpenDynamic, adLockBatchOptimistic
    If rstSNs.EOF Then
        bFoundInWO = False
    Else
        bFoundInWO = True
        'show users which SNs already exist
        strMsg = "The following serial Numbers are already assigned to a Work Order:" & vbCrLf & vbCrLf
        Do While rstSNs.EOF = False
            strMsg = strMsg & vbCrLf & rstSNs!SerialNumber
            rstSNs.MoveNext
        Loop
        MsgBox strMsg
        ErrorLogger "CheckForExistingSNs", strMsg
    End If
    
    If rstSNs.State = 1 Then
        rstSNs.Close
    End If

    If bFoundInBuilt = False And bFoundInWO = False Then
        adoCheckForExistingSNsByList = False
    Else
        'serial numbers exist in one of the two tables
        adoCheckForExistingSNsByList = True
    End If

End Function

'------------------------------------------------------------------------------------------------
'Created:   04/13/03    By:Jeff Pierce
'Purpose:   To populate the frmWOOptions.lstUnitLR listbox of available units
'Inputs:    <none>
'Outputs:   [inserts data into listbox]
'   Rev'd:      10/22/03    By: Ray Raf
'   Note:        added handling of MCOV table name for new and old process
'-------------------------------------------------------------------------------------------------
Public Sub adoGetUnitLettersAndRatings()
    Dim X As Integer
    Dim strTemp As String
    Dim varMCOVTable  As String, varCompTable As String
    Dim SQL As String
    Dim strMCOV As Single
    Dim strMCOVTable As String
    
    Select Case globCurrentWO.strType
        Case 1 'porc
            Select Case LCase(globCurrentWO.strBuildPath)
                Case "dual rated"
                    'dual rated arresters ALWAYS have exactly two units.
                    SQL = "SELECT * FROM MCOV_Dr_Porc "
                    SQL = SQL & "WHERE CatalogNumber='" & globCurrentWO.strCatNumber & "' "
                    Set adoRS = dbconSetupData.Execute(SQL)
                    If Not adoRS.EOF Then
                        'the catalog number was found in the table
                        strMCOV = adoRS.Fields("MCOV_Unit_1")
                        If strMCOV <> "0" Then
                            frmWOOptions.lstUnitLR.AddItem "A) Unit 1: " & strMCOV & "kV MCOV"
                        End If
                        strMCOV = adoRS.Fields("MCOV_Unit_2_Lower")
                        If strMCOV <> "0" Then
                            frmWOOptions.lstUnitLR.AddItem "B) Unit 2: " & strMCOV & "kV MCOV"
                        End If
                    End If
                Case "cubicle mount"
                    'cubicle mount can have up to 3 units
                    With globCurrentUnit
                        SQL = "SELECT * FROM MCOV_Cube "
                        SQL = SQL & "WHERE Family='" & .Family4Lookup & "' "
                        SQL = SQL & "AND ArresterRating='" & .Rating & "' "
                        SQL = SQL & "AND String_MCOV ='" & .MCOV & "' "
                        SQL = SQL & "AND PorcRingNumber='" & .PorcelainRingCode & "' ;"
                        Set adoRS = dbconSetupData.Execute(SQL)
                    End With
                    
                    If Not adoRS.EOF Then
                        For X = 1 To 3
                            strMCOV = adoRS.Fields("MCOV_Unit_" & X)
                            If strMCOV <> "0" Then
                                frmWOOptions.lstUnitLR.AddItem Chr(X + 64) & ") Unit " & X & ": " & strMCOV & "kV MCOV"
                            End If
                        Next X
                    End If
                Case Else
                ' added test for calcined disk builds - 14Apr2011
                    If globCalcinedDisksUsedInPorc = True Then
                        strMCOVTable = "MCOV_Porc_Rev1"
                    Else
                        strMCOVTable = "MCOV_Porc"
                    End If
                    'all others can have up to 4 units
                    With globCurrentUnit
                        SQL = "SELECT * FROM " & strMCOVTable & " "
                        SQL = SQL & "WHERE Family='" & .Family4Lookup & "' "
                        SQL = SQL & "AND ArresterRating='" & .Rating & "' "
                        SQL = SQL & "AND string_MCOV='" & .MCOV & "' "
                        SQL = SQL & "AND PorcRingNumber='" & .PorcelainRingCode & "' ;"
                        Set adoRS = dbconSetupData.Execute(SQL)
                    End With
                    If Not adoRS.EOF Then
                        For X = 1 To 4
                            strMCOV = adoRS.Fields("MCOV_Unit_" & X)
                            If strMCOV <> "0" Then
                                frmWOOptions.lstUnitLR.AddItem Chr(X + 64) & ") Unit " & X & ": " & strMCOV & "kV MCOV"
                            End If
                        Next X
                    End If
            End Select
        Case 2 'poly
            Select Case LCase(globCurrentWO.strBuildPath)
                Case "dual rated"
                    SQL = "SELECT * FROM MCOV_Dr_Poly "
                    SQL = SQL & "WHERE CatalogNumber = '" & globCurrentWO.strCatNumber & "' ;"
                    Set adoRS = dbconSetupData.Execute(SQL)
                    If Not adoRS.EOF Then
                        strMCOV = adoRS.Fields("MCOV_Unit_1")
                        If strMCOV <> "0" Then
                            frmWOOptions.lstUnitLR.AddItem "A) Unit 1: " & strMCOV & "kV MCOV"
                        End If
                        strMCOV = adoRS.Fields("MCOV_Unit_2_Lower")
                        If strMCOV <> "0" Then
                            frmWOOptions.lstUnitLR.AddItem "B) Unit 2: " & strMCOV & "kV MCOV"
                        End If
                    End If
                Case "1in", "12mm"
                    GetPolyTableNames globCurrentUnit, varCompTable, varMCOVTable 'added 10/22/03
'                    If globCurrentWO.strBuildPath = "1in" Then
'                        varMCOVTable = "MCOV_1in"
'                    Else
'                        varMCOVTable = "MCOV_12mm"
'                    End If
                    SQL = "SELECT * FROM " & varMCOVTable & " "
                    SQL = SQL & "WHERE Family='" & globCurrentUnit.Family & "' "
                    SQL = SQL & "AND String_MCOV='" & globCurrentUnit.MCOV & "' "
                    SQL = SQL & "AND ShedCode='" & globCurrentUnit.PolyShedRingCode & "' "
                    SQL = SQL & "AND ArresterRating = " & globCurrentUnit.Rating & " "
                    Set adoRS = dbconSetupData.Execute(SQL)
                    If Not adoRS.EOF Then
                        'arrester was found.  Get Unit data
                        For X = 1 To 6
                            strMCOV = adoRS.Fields("MCOV_Unit_" & X)
                            If strMCOV <> "0" Then
                                frmWOOptions.lstUnitLR.AddItem Chr(X + 64) & ") Unit " & X & ": " & strMCOV & "kV MCOV"
                            End If
                        Next
                    End If
            End Select
    End Select
    If adoRS.State = 1 Then
        adoRS.Close
    End If
End Sub

'------------------------------------------------------------------------------------------------
'Created:   04/13/03    By:Jeff Pierce
'Purpose:   To call the appropriate checkAllowedDisks function, based on type (poly, porc)
'Inputs:    varNumDisks = number of disks for the arrestor (entered by user)
'Outputs:   true = allowed disks is valid, false otherwise
'-------------------------------------------------------------------------------------------------
'Public Function adoCheckAllowedDisks(iNumDisks As Integer, iNum3kVDisks As Integer, iNum5kVDisks As Integer) As Boolean
' DLimburg Change 2010_1 3Feb2010 Modified and added code to allow for the 3 and 5 calcined disks
Public Function adoCheckAllowedDisks(iNumDisks As Integer, CP_3kVDiskCnt As Integer, CP_5kVDiskCnt As Integer, _
                                                                 Gray_3kVDiskCnt As Integer, Gray_5kVDiskCnt As Integer) As Boolean
    Select Case globCurrentWO.strType
        Case 1: adoCheckAllowedDisks = adoCheckAllowedDisksPorc(iNumDisks)
'        Case 2: adoCheckAllowedDisks = adoCheckAllowedDisksPoly(iNumDisks, iNum3kVDisks, iNum5kVDisks)
        Case 2: adoCheckAllowedDisks = adoCheckAllowedDisksPoly(iNumDisks, CP_3kVDiskCnt, CP_5kVDiskCnt, Gray_3kVDiskCnt, Gray_5kVDiskCnt)
    End Select
End Function

'------------------------------------------------------------------------------------------------
'Created:   04/13/03    By:Jeff Pierce
'Purpose:   To do several checks on the number of disks allowed for the arrestor
'Inputs:    varNumDisks = number of disks for the arrestor (entered by user)
'Outputs:   true = allowed disks is valid, false otherwise
'   Rev'd:      10/21/03    By: Ray Raf
'   Note:       changed lookup for stud and disktype to use function
'                  also added changes for new and old processes for component and MCOV
'-------------------------------------------------------------------------------------------------
'Public Function adoCheckAllowedDisksPoly(varNumDisks As Integer, varNum3kVDisks As Integer, varNum5kVDisks As Integer)
' DLimburg Change 2010_1 3Feb2010 Modified and added code to allow for the 3 and 5 calcined disks
Public Function adoCheckAllowedDisksPoly(varNumDisks As Integer, CP_3kVDiskCnt As Integer, CP_5kVDiskCnt As Integer, _
                                                                 Gray_3kVDiskCnt As Integer, Gray_5kVDiskCnt As Integer)
    Dim rstNameplate As New ADODB.Recordset
    Dim rstNumDisks As New ADODB.Recordset
    Dim rstHousing As New ADODB.Recordset
    Dim rstSheds As New ADODB.Recordset
    Dim SQL As String
    Dim varFound As Boolean
    Dim varUnitNum As Integer
    Dim iLowLim As Integer, iHighLim As Integer
    Dim strMsg As String
    Dim iAllowedDisks As Integer
    Dim strFieldName As String
    Dim iTotalDisks As Integer
    Dim varMCOVTable As String
    Dim varCompTable As String
    Dim varNumSheds As Integer
    Dim varHousingSQL As String
    Dim bResult As Boolean

    '--get the diskType and stud, based on family and rating
'    SQL = "SELECT DiskType, Stud FROM NameplatePoly "
'    SQL = SQL & "WHERE Family='" & globCurrentUnit.Family & "' "
'    SQL = SQL & "AND MinRating <= '" & globCurrentUnit.Rating & "' "
'    SQL = SQL & "AND MaxRating >= '" & globCurrentUnit.Rating & "' "
'    rstNameplate.Open SQL, dbconSetupData, adOpenDynamic, adLockBatchOptimistic
'    If rstNameplate.EOF Then
'        MsgBox "The disk type and stud could not be retrieved for this family and rating."
'        adoCheckAllowedDisksPoly = False
'        Exit Function
'    End If
'
'    globCurrentUnit.DiskType = Trim(rstNameplate!DiskType)
'    globCurrentUnit.Stud = Trim(rstNameplate!Stud)
'
'    If rstNameplate.State = 1 Then
'        rstNameplate.Close
'    End If
    
    With globCurrentUnit
' DLimburg Change 2009_1 11Feb2009 Added .DiskType2 to proccedure call
        GetPolyStudDiskandProcess globCurrentWO.strCatNumber, .Rating, .Stud, _
                        .DiskType, .DiskType2, .Process
    End With
    
    bResult = GetPolyTableNames(globCurrentUnit, varCompTable, varMCOVTable)
    If bResult = False Then
        'something is wrong
        strMsg = "could not resolve Component and MCOV table names!"
        ErrorLogger "adoCheckAllowedDisksPoly", strMsg      'this is a data error. log it so someone can follow up on it
        adoCheckAllowedDisksPoly = False
        Exit Function
    End If
'    If globCurrentUnit.Stud = "12mm" Then
'        varMCOVTable = "MCOV_12mm"
'        varCompTable = "COMPONENTS_12mm"
'    Else
'        varMCOVTable = "MCOV_1in"
'        varCompTable = "COMPONENTS_1in"
'    End If

    '--get diskquantity based on diskType and string_mcov
' D Limburg 10Feb2009 Changed follow line to accommodate 3kV and 5kV AT32G disk implimenation
' ** Start of change
'    SQL = "SELECT DiskQuantity FROM DiskQuanPoly "
'    SQL = "SELECT DiskQuantity, DiskQuantity_3kV, DiskQuantity_5kV FROM DiskQuanPoly_rev1 "
' DLimburg Change 2010_1 3Feb2010 Modified and added code to allow for the 3 and 5 calcined disks
    SQL = "SELECT DiskQuantity, DiskQuantity_3kV, DiskQuantity_5kV FROM DiskQuanPoly_rev2 "
' ** End of change
    SQL = SQL & "WHERE DiskType='" & globCurrentUnit.DiskType & "' "
    SQL = SQL & "AND Design_MCOV ='" & globCurrentUnit.UnitMCOV & "' "
    SQL = SQL & "AND Family = '" & globCurrentUnit.Family & "' "
    
    Set rstNumDisks = dbconSetupData.Execute(SQL)
    
    varFound = False
    '--see if any of the results matches the number of disks entered
    If rstNumDisks.EOF = False Then
        iHighLim = 0
        iLowLim = 100
        Do Until rstNumDisks.EOF
            'check each record returned to see if any match the number of disks
            'that the user wants to use.
            iAllowedDisks = rstNumDisks.Fields("DiskQuantity")
            If varNumDisks = iAllowedDisks Then
                varFound = True
                'since a match was found we will not need the high and low limits
                Exit Do
            End If
            'find allowable high and low limits in case no match found
            If iHighLim < iAllowedDisks Then
                iHighLim = iAllowedDisks
            End If
            If iLowLim > iAllowedDisks Then
                iLowLim = iAllowedDisks
            End If
            rstNumDisks.MoveNext
        Loop
    Else
        'no records were returned.  This most likely means that there is an error in the data
        strMsg = "No Disk Quantities were found for Disk Type:" & globCurrentUnit.DiskType
        strMsg = strMsg & vbCrLf & "at " & globCurrentUnit.DesignMCOV & " kV"
        MsgBox strMsg
        ErrorLogger "adoCheckAllowedDisksPoly", strMsg      'this is a data error. log it so someone can follow up on it
        adoCheckAllowedDisksPoly = False
        Exit Function
    End If


    If varFound = False Then
        strMsg = "The number of disks entered is invalid for this arrestor." & vbCrLf
        strMsg = strMsg & "The allowable quantities are " & iLowLim & " to " & iHighLim
        MsgBox strMsg, vbInformation, "Error"
        adoCheckAllowedDisksPoly = False
        Exit Function
    End If


    '--determine the number of sheds for this unit
    Select Case globCurrentWO.strBuildPath
        Case "Dual Rated"
            strFieldName = "Unit_" & globCurrentUnit.UnitNumber & "_Housing"
            varHousingSQL = "SELECT " & strFieldName & " "
            varHousingSQL = varHousingSQL & "FROM HousingDrPoly WHERE CatalogNumber='" & globCurrentWO.strCatNumber & "' "
        Case Else
            strFieldName = "Housing_" & globCurrentUnit.UnitNumber
            varHousingSQL = "SELECT " & strFieldName & " FROM " & varMCOVTable
            varHousingSQL = varHousingSQL & " WHERE Family= '" & globCurrentUnit.Family & "' "
            varHousingSQL = varHousingSQL & "AND string_MCOV='" & globCurrentUnit.MCOV & "' "
            varHousingSQL = varHousingSQL & "AND ShedCode='" & globCurrentUnit.PolyShedRingCode & "' "
            varHousingSQL = varHousingSQL & "AND ArresterRating = " & globCurrentUnit.Rating & " "
    End Select

    rstHousing.Open varHousingSQL, dbconSetupData, adOpenDynamic, adLockBatchOptimistic
    If rstHousing.EOF = True Then
        MsgBox "No housing was found for this unit!.", vbInformation, "Error"
        adoCheckAllowedDisksPoly = False
        Exit Function
    End If

    globCurrentUnit.UnitSections = CInt(Trim(rstHousing.Fields(strFieldName).Value))

    If rstHousing.State > 0 Then
        rstHousing.Close
    End If
    
    '--retrieve number of sections based on diskType and diskQuantity

'    SQL = "SELECT * FROM " & varCompTable & " WHERE DiskType='" & globCurrentUnit.DiskType & "' "
'    SQL = SQL & "AND DiskQuantity='" & varNumDisks & "' "
'    SQL = SQL & " AND DiskQuantity_3kv = " & varNum3kVDisks
'    SQL = SQL & " AND DiskQuantity_5kV = " & varNum5kVDisks
'    SQL = SQL & " AND Sections = " & globCurrentUnit.UnitSections & " "
' DLimburg Change 2010_1 3Feb2010 Modified and added code to allow for the 3 and 5 calcined disks
    SQL = "SELECT DiskQuanPoly_rev2.*, ModuleComponents_Poly.Sections, ModuleComponents_Poly.ModuleHeight," & _
                " ModuleComponents_Poly.EndElectrodeHgt , ModuleComponents_Poly.MidElectrodes, ModuleComponents_Poly.HeatSinks" & _
                " FROM DiskQuanPoly_rev2 LEFT OUTER JOIN ModuleComponents_Poly ON " & _
                " DiskQuanPoly_rev2.DiskQuantity = ModuleComponents_Poly.TotalDiskCnt" & _
                " AND DiskQuanPoly_rev2.DiskQuantity_3kV = (ModuleComponents_Poly.Total3kVDiskCnt)" & _
                " AND DiskQuanPoly_rev2.DiskQuantity_5kV = (ModuleComponents_Poly.Total5kVDiskCnt)" & _
                " WHERE (DiskQuanPoly_rev2.DiskType = '" & globCurrentUnit.DiskType & "')" & _
                " AND (DiskQuanPoly_rev2.Design_MCOV = '" & globCurrentUnit.UnitMCOV & "')" & _
                " AND (DiskQuanPoly_rev2.Family = '" & globCurrentUnit.Family & "')" & _
                " AND (ModuleComponents_Poly.CP_3kVDiskCnt = " & CP_3kVDiskCnt & ")" & _
                " AND (ModuleComponents_Poly.CP_5kVDiskCnt = " & CP_5kVDiskCnt & ")" & _
                " AND (ModuleComponents_Poly.Gray_3kVDiskCnt = " & Gray_3kVDiskCnt & ")" & _
                " AND (ModuleComponents_Poly.Gray_5kVDiskCnt = " & Gray_5kVDiskCnt & ")" & _
                " AND (ModuleComponents_Poly.Stud = '" & globCurrentUnit.Stud & "')" & _
                " AND (ModuleComponents_Poly.Sections = " & globCurrentUnit.UnitSections & ")"

    rstSheds.Open SQL, dbconSetupData, adOpenDynamic, adLockBatchOptimistic
    If rstSheds.EOF = True Then
        MsgBox "The Number of disks specified will not fit in any housing for this unit!", vbInformation, "error"
        adoCheckAllowedDisksPoly = False
        Exit Function
    End If

    '--if we make it to here, it's a valid quantity
    adoCheckAllowedDisksPoly = True
End Function

'------------------------------------------------------------------------------------------------
'Created:   04/13/03    By:Jeff Pierce
'Purpose:   To do several checks on the number of disks allowed for the arrestor
'Inputs:    varNumDisks = number of disks for the arrestor (entered by user)
'Outputs:   true = allowed disks is valid, false otherwise
'   Rev'd:  04/17/03    By: Ray Raf
'   Note:       added explicit types to input param and function return
'                  added allowable limits for wrong Number Of Disks error mesage
'   Rev'd:  10/31/03      By: Ray Raf
'   Note:   changed disk quantity lookup to use actually family instead of
'               Family4Lookup.
'-------------------------------------------------------------------------------------------------
Public Function adoCheckAllowedDisksPorc(varNumDisks As Integer) As Boolean
    Dim varSQL As String
    Dim varFound As Boolean
    Dim varUnitNum As Integer
    Dim varBuildTable As String
    Dim varHousingSQL As String
    Dim iLowLim As Integer, iHighLim As Integer
    Dim strMsg As String
    Dim iAllowedDisks As Integer
    Dim strFieldName As String
    Dim iTotalDisks As Integer
    Dim adoNumDisks As New ADODB.Recordset
    Dim strMCOVTable As String
    
    '--determine the diskType based on the build path
    Select Case globCurrentWO.strBuildPath
        Case "Dual Rated"
            varSQL = "SELECT DiskType FROM MCOV_Dr_Porc WHERE CatalogNumber='" & globCurrentWO.strCatNumber & "' ;"
        Case "Cubicle Mount"
            varSQL = "Select DiskType FROM MCOV_Cube WHERE Family='" & globCurrentUnit.Family4Lookup & "' "
            varSQL = varSQL & "AND String_MCOV='" & globCurrentUnit.MCOV & "' AND ArresterRating=" & globCurrentUnit.Rating & " "
            varSQL = varSQL & "AND PorcRingNumber = " & globCurrentUnit.PorcelainRingCode & " ;"
        Case Else
            ' added test for calcined disk builds - 14Apr2011
            If globCalcinedDisksUsedInPorc = True Then
                strMCOVTable = "MCOV_Porc_Rev1"
            Else
                strMCOVTable = "MCOV_Porc"
            End If
            varSQL = "Select DiskType FROM " & strMCOVTable & " WHERE Family='" & globCurrentUnit.Family4Lookup & "' "
            varSQL = varSQL & "AND String_MCOV='" & globCurrentUnit.MCOV & "' AND ArresterRating=" & globCurrentUnit.Rating & " "
            varSQL = varSQL & "AND PorcRingNumber = '" & globCurrentUnit.PorcelainRingCode & "' ;"
    End Select
    Set adoRS = dbconSetupData.Execute(varSQL)
    'since the catalog number was verified from the MCOV tables before getting her, there will be a DiskType
    globCurrentUnit.DiskType = adoRS.Fields("DiskType")
    '--retrieve the diskQuantity based on diskType and UnitMCOV
    varSQL = "SELECT DiskQuantity_Total FROM DiskQuanPorc_Rev1 "   ' changed table name from DiskQuanPorc to DiskQuanPorc_Rev1 and DiskQuantity to DiskQuantity_Total - DLimburg16Feb2011
    varSQL = varSQL & "WHERE DiskType='" & globCurrentUnit.DiskType & "' "
    varSQL = varSQL & "AND Design_MCOV='" & globCurrentUnit.UnitMCOV & "' "
    varSQL = varSQL & "AND Family = '" & globCurrentUnit.Family & "' ORDER BY DiskQuantity_Total;"  ' changed field DiskQuantity to DiskQuantity_Total
    Set adoNumDisks = dbconSetupData.Execute(varSQL)
    
    varFound = False
    '--see if any of the results matches the number of disks entered
    If adoNumDisks.EOF = False Then
        iHighLim = 0
        iLowLim = 100
        Do Until adoNumDisks.EOF
            'check each record returned to see if any match the number of disks
            'that the user wants to use.
            iAllowedDisks = adoNumDisks.Fields("DiskQuantity_Total")
            If varNumDisks = iAllowedDisks Then
                varFound = True
                'since a match was found we will not need the high and low limits
                Exit Do
            End If
            'find allowable high and low limits in case no match found
            If iHighLim < iAllowedDisks Then
                iHighLim = iAllowedDisks
            End If
            If iLowLim > iAllowedDisks Then
                iLowLim = iAllowedDisks
            End If
            adoNumDisks.MoveNext
        Loop
    Else
        'no records were returned.  This most likely means that there is an error in the data
        strMsg = "No Disk Quantities were found for Disk Type:" & globCurrentUnit.DiskType
        strMsg = strMsg & vbCrLf & "at " & globCurrentUnit.DesignMCOV & " kV"
        MsgBox strMsg
        ErrorLogger "adoCheckAllowedDisksPorc", strMsg      'this is a data error. log it so someone can follow up on it
        adoCheckAllowedDisksPorc = False
        Exit Function
    End If
    
    If varFound = False Then
        strMsg = "The number of disks entered is invalid for this arrestor." & vbCrLf
        strMsg = strMsg & "The allowable quantities are " & iLowLim & " to " & iHighLim
        MsgBox strMsg, vbInformation, "Error"
        adoCheckAllowedDisksPorc = False
        Exit Function
    End If
    
    varUnitNum = globCurrentUnit.UnitNumber
    
    '--determine the build table and housing query string based on build path
    strFieldName = "Unit_" & varUnitNum & "_Housing"
    Select Case globCurrentWO.strBuildPath
        Case "Dual Rated"
            varBuildTable = "BuildDrPorc"
            varHousingSQL = "SELECT " & strFieldName & " FROM HousingDrPorc WHERE CatalogNumber='" & globCurrentWO.strCatNumber & "' ;"
        Case "Cubicle Mount"
            varBuildTable = "BuildTableCube_Rev1"   ' Chg'd from BuildTableCube to BuildTableCube_Rev1 DLimburg 2/16/2011
            varHousingSQL = "SELECT " & strFieldName & " FROM HousingPorc WHERE RingCode='" & globCurrentUnit.PorcelainRingCode & globCurrentUnit.HousingCode & "';"
        Case Else
            varBuildTable = "BuildTablePorc_Rev1"   ' Chg'd from BuildTablePorc to BuildTablePorc_Rev1 DLimburg 2/16/2011
            varHousingSQL = "SELECT " & strFieldName & " FROM HousingPorc WHERE RingCode='" & globCurrentUnit.PorcelainRingCode & globCurrentUnit.HousingCode & "';"
    End Select
    Set adoRS = dbconSetupData.Execute(varHousingSQL)
    
    If adoRS.EOF Then
        strMsg = "There wasn't a Unit " & varUnitNum & " housing code for Ring code '" & globCurrentUnit.PorcelainRingCode & "'!"
        MsgBox strMsg, vbInformation, "error"
        ErrorLogger "adoCheckAllowedDisksPorc", strMsg      'this is a data error. log it so someone can follow up on it
        adoCheckAllowedDisksPorc = False
        Exit Function
    End If
    
    globHousingCode = adoRS.Fields(strFieldName)
    
    '--retrieve the total number of disks based on PorcelainHousing
    varSQL = "SELECT TotDisks FROM " & varBuildTable & " "
    varSQL = varSQL & "WHERE PorcelainHousing='" & globHousingCode & "' "
    varSQL = varSQL & "And DiskType = '" & globCurrentUnit.DiskType & "' "
    Set adoRS = dbconSetupData.Execute(varSQL)
    
    If adoRS.EOF Then
        strMsg = "The porcelain housing code '" & globHousingCode & "' was not found in the build table!"
        MsgBox strMsg, vbInformation, "Error"
        ErrorLogger "adocheckAllowedDisksPorc", strMsg      'this is a data error. log it so someone can follow up on it
        adoCheckAllowedDisksPorc = False
        Exit Function
    End If
        
    varFound = False
    '--see if any of the returned results match the number of disks entered
    If adoRS.EOF = False Then
        iLowLim = 100
        iHighLim = 0
        Do Until adoRS.EOF
            iTotalDisks = adoRS.Fields("TotDisks")
            If varNumDisks = iTotalDisks Then
                varFound = True
                Exit Do
            End If
            'find allowable high and low limits in case no match found
            If iHighLim < iTotalDisks Then
                iHighLim = iTotalDisks
            End If
            If iLowLim > iTotalDisks Then
                iLowLim = iTotalDisks
            End If
            adoRS.MoveNext
        Loop
    End If
    
    If varFound = False Then
        strMsg = "The number of disks you've entered does not fit in Housing '" & globHousingCode & "'." & vbCrLf
        strMsg = strMsg & "The allowable range is " & iLowLim & " to " & iHighLim & "."
        MsgBox strMsg, vbInformation, "Error"
        ErrorLogger "adocheckAllowedDisksPorc", strMsg      'this is a data error. log it so someone can follow up on it
        adoCheckAllowedDisksPorc = False
        Exit Function
    End If
    
    '--if we make it to here, the number of disks entered is valid
    adoCheckAllowedDisksPorc = True
End Function

'------------------------------------------------------------------------------------------------
'Created:   03/28/03    By:Jeff Pierce
'Purpose:   An easy way to set the catalog number (for my temporary purposes only)
'Inputs:    <none>
'Outputs:   <none>
'-------------------------------------------------------------------------------------------------
Public Sub adoSetWorkOrder()
    dbconSetupData.Execute "UPDATE WO_Info SET CatalogNumber='AZG2E02E003123';"
End Sub

Public Function VerifyCatalogNumber(strCatNumber As String) As Boolean
    '------------------------------------------------------------------------------------------------
    'Created:   04/18/03    By:Ray Raf
    'Purpose:   To verify that this is a valid catalog number.  Check is
    '                 made on the MCOV tables only.
    'Inputs:    The Catalog Number to check
    'Returns:   True if an entry found for the catalog number, False otherwise.
    '-------------------------------------------------------------------------------------------------
    Dim strFamily As String, iRating As Integer, strMCOV As String, strPorcRingCode As String
    Dim strPolyShedCode As String
    Dim iType As Integer
    Dim strMsg As String
    Dim strBuildPath As String
    Dim rstMCOV As New ADODB.Recordset
    Dim strMCOVTable As String
    Dim SQL As String, SQLFamily As String
    Dim strCrossedFamily As String
    Dim iLength As Integer
    Dim strStud As String, strDiskType As String, strDiskType2 As String, strProcess As String
    Dim strCompTable As String
    
    iType = getType(strCatNumber)
    If iType <> 1 And iType <> 2 Then
        strMsg = "The catalog number is incorrect because it does not start with an A or U!"
        MsgBox strMsg, vbInformation, "Error In Catalog Number"
        VerifyCatalogNumber = False
        Exit Function
    End If
    
    'check lengths to make sure we don't error when trying to extract data from catalog number
    If iType = 1 Then
        If Len(strCatNumber) <> 14 Then
            strMsg = "Catalog Number '" & strCatNumber & "' is not the required length of 14 characters!"
            MsgBox strMsg, vbInformation, "Error In Catalog Number"
            VerifyCatalogNumber = False
            Exit Function
        End If
    Else
        'check poly lent
        iLength = Len(strCatNumber)
        If iLength < 10 Or iLength > 15 Then
            strMsg = "Catalog Number '" & strCatNumber & "' is not the right length!"
            MsgBox strMsg, vbInformation, "Error In Catalog Number"
            VerifyCatalogNumber = False
            Exit Function
        End If
    End If
    
    'thats a start.  It could be porcelain or polymer.
    'get the build path because this will check out the rest of the family name
    strBuildPath = getBuildPath(strCatNumber, iType)
    If LCase(strBuildPath) = "not found" Then
        'the basics of the family could not be obtained for this catalog number
        strMsg = "There is an unknown error in the catalog number!"
        MsgBox strMsg, vbInformation, "Error In Catalog Number"
        VerifyCatalogNumber = False
        Exit Function
    End If
        
    'A build path was found.  at least the basic structure of the catalog number is right
    'From here, any errors could be in either the catalog number OR the database.
    Select Case iType
        Case 1 ' porc
            Select Case LCase(strBuildPath)
                Case "dual rated"
                    SQL = "SELECT * FROM MCOV_Dr_Porc "
                    SQL = SQL & "WHERE CatalogNumber='" & strCatNumber & "' "
                    Set rstMCOV = dbconSetupData.Execute(SQL)
                    If rstMCOV.EOF Then
                        'the catalog number not was found in the table
                        strMsg = "Dual Rated catalog number " & strCatNumber & " was not found in the database!"
                        MsgBox strMsg, vbInformation, "Error Verifying Catalog Number"
                        VerifyCatalogNumber = False
                        Exit Function
                    End If
                    'that's all for dual rated verify
                Case "cubicle mount", "aze/azg"
                    'only difference between standard and Cubicle is the table to get data from
                    If LCase(strBuildPath) = "cubicle mount" Then
                        strMCOVTable = "MCOV_Cube"
                    Else
                    ' added test for calcined disk builds - 14Apr2011
                        If globCalcinedDisksUsedInPorc = True Then
                            strMCOVTable = "MCOV_Porc_Rev1"
                        Else
                            strMCOVTable = "MCOV_Porc"
                        End If
                    End If
                    
                    strFamily = getFamily(iType, strCatNumber)
                    iRating = getRating(iType, strCatNumber)
                    strMCOV = getMCOV(iType, strCatNumber)
                    strPorcRingCode = getRingCode(iType, strCatNumber)
                        
                    SQL = "SELECT * FROM " & strMCOVTable & " "
                    SQL = SQL & "WHERE ArresterRating='" & iRating & "' "
                    SQL = SQL & "AND String_MCOV ='" & strMCOV & "' "
                    SQL = SQL & "AND PorcRingNumber='" & strPorcRingCode & "' "
                    SQLFamily = "AND Family='" & strFamily & "' ;"
                    Set rstMCOV = dbconSetupData.Execute(SQL & SQLFamily)
                    
                    If rstMCOV.EOF Then
                        'no record found with original family, try using cross referenced family
                        strCrossedFamily = PorcCrossRef(strFamily, iRating)
                        SQLFamily = "AND Family='" & strCrossedFamily & "' ;"
                        Set rstMCOV = dbconSetupData.Execute(SQL & SQLFamily)
                                              
                        If rstMCOV.EOF Then
                            'it wasn't found crossed either.  Bad number
                            strMsg = "Catalog Number '" & strCatNumber & "' could not be found in table '" & strMCOVTable & "' "
                            MsgBox strMsg, vbInformation, "Error Verifying Catalog Number"
                            VerifyCatalogNumber = False
                            Exit Function
                        End If
                    End If
                    'number found in the MCOV table, move on to next check
                Case Else
                    'all others can have up to 4 units
                    strMsg = "No Valid Porcelain Build Path found.  Catalog Number '" & strCatNumber & "' is invalid!            "
                    MsgBox strMsg, vbInformation, "Invalid Catalog Number"
                    VerifyCatalogNumber = False
                    Exit Function
            End Select  'for porc build path
        Case 2 'poly
            Select Case LCase(strBuildPath)
                Case "dual rated"
                    SQL = "SELECT * FROM MCOV_Dr_Poly "
                    SQL = SQL & "WHERE CatalogNumber='" & strCatNumber & "' "
                    Set rstMCOV = dbconSetupData.Execute(SQL)
                    If rstMCOV.EOF Then
                        'the catalog number not was found in the table
                        strMsg = "Dual Rated catalog number " & strCatNumber & " was not found in the database!"
                        MsgBox strMsg, vbInformation, "Error Verifying Catalog Number"
                        VerifyCatalogNumber = False
                        Exit Function
                    End If
                    'that's all for dual rated verify
                Case "1in", "12mm"
                    
                    strFamily = getFamily(iType, strCatNumber)
                    iRating = getRating(iType, strCatNumber)
                    strMCOV = getMCOV(iType, strCatNumber)
                    strPolyShedCode = getRingCode(iType, strCatNumber)
                        
                    'only realy need the process
' DLimburg Change 2009_1 11Feb2009 Added strDiskType2 to proccedure call
                    GetPolyStudDiskandProcess strCatNumber, iRating, strStud, _
                                        strDiskType, strDiskType2, strProcess
                    
                    'build mcov table name
                    If LCase(strBuildPath) = "12mm" Then
                        strMCOVTable = "MCOV_12mm_" & strProcess
                    Else
                        strMCOVTable = "MCOV_1in"
                    End If
                        
                    SQL = "SELECT * FROM " & strMCOVTable & " "
                    SQL = SQL & "WHERE ArresterRating='" & iRating & "' "
                    SQL = SQL & "AND String_MCOV ='" & strMCOV & "' "
                    SQL = SQL & "AND ShedCode = '" & strPolyShedCode & "' "
                    SQL = SQL & "AND Family ='" & strFamily & "' "
                    Set rstMCOV = dbconSetupData.Execute(SQL & SQLFamily)
                    
                    If rstMCOV.EOF Then
                        'no entry in the MCOV table, verify failed.
                        strMsg = "Catalog Number '" & strCatNumber & "' could not be found in table '" & strMCOVTable & "' "
                        MsgBox strMsg, vbInformation, "Error Verifying Catalog Number"
                        VerifyCatalogNumber = False
                        Exit Function
                    End If
                    'number found in the MCOV table, move on to next check
                Case Else
                    'not a legit build path
                    strMsg = "No Valid Porcelain Build Path found.  Catalog Number '" & strCatNumber & "' is invalid!            "
                    MsgBox strMsg, vbInformation, "Invalid Catalog Number"
                    VerifyCatalogNumber = False
                    Exit Function
            End Select  'for porc build path
    End Select  'for type
    
    VerifyCatalogNumber = True
            
    
End Function

Public Function PorcCrossRef(strFamily As String, iRating As Integer) As String
    '------------------------------------------------------------------------------------------------
    'Created:   04/18/03    By:Ray Raf
    'Purpose:   to do a cross reference lookup on a porcelain family
    'Inputs:    the family to cross reference
    'Returns:   the family that should be used for lookups.  If no cross
    '                found, the function returns the passed family (because
    '                that is what will have to be used for lookups anyway.
    '-------------------------------------------------------------------------------------------------
    Dim rstCross As New ADODB.Recordset
    Dim SQL As String
    
    SQL = "SELECT * FROM CrossRefPorC "
    SQL = SQL & "WHERE FamilyToCross = '" & strFamily & "' "
    SQL = SQL & "AND LowRating <= " & iRating & " "
    SQL = SQL & "AND HighRating >= " & iRating & " "
    
    rstCross.Open SQL, dbconSetupData, adOpenStatic, adLockReadOnly
    If rstCross.EOF Then
        'cross family not found.  return original family
        PorcCrossRef = strFamily
    Else
        PorcCrossRef = rstCross!CrossedFamily
    End If
    
    Set rstCross = Nothing
End Function

Public Function LookupDesignMCOV(strMCOV As String) As String
    '------------------------------------------------------------------------------------------------
    'Created:   04/18/03    By:Ray Raf
    'Purpose:   Since exact matches can not be done on string values, all
    '                fractional kV MCOV values (ex 2.55) that need to be used
    '                as keys for data lookup have been coded in the database to
    '                match the codes used in the catalog numbers.  This routine
    '                with find the actual MCOV value based on this code.
    'Inputs:    the MCOV Value to lookup.
    'Returns:   the actual MCOV value to use, as a string so that proper
    '               number of decimal places can be used for display.  If no
    '               entry is found, then the MCOV value has no decimal portion
    '               and can be used as is.  In that case the string originally passed
    '               with be returned (minus any leading 0's)
    '-------------------------------------------------------------------------------------------------
    Dim rstActualMCOV As New ADODB.Recordset
    Dim SQL As String
    
    SQL = "SELECT * FROM Actual_MCOV WHERE String_MCOV = '" & strMCOV & "' "
    rstActualMCOV.Open SQL, dbconSetupData, adOpenStatic, adLockReadOnly
    
    If rstActualMCOV.EOF = False Then
        'a lookup value was found, return it
        LookupDesignMCOV = rstActualMCOV!Actual_MCOV
    Else
        'convert passed string to a number then back to a string to eliminate any leading zeros
        On Error Resume Next
        LookupDesignMCOV = CStr(Val(strMCOV))
        If Err.Number <> 0 Then
            'somehow a non-numeric string was passed to this routine.  send back a 0 MCOV
            LookupDesignMCOV = "0"
            ErrorLogger "LookupDesignMCOV", "Tried to lookup an MCOV of " & strMCOV
            Err.Clear
        End If
    End If
    
End Function

Public Function adoCopyDiskData() As Boolean
    '------------------------------------------------------------------------------------------------
    'Created:   04/03    By:Jeff Pierce
    'Purpose:   To copy the disks entered for an arrester from the Available
    '                 Disks table to the Consumed Disks table
    'Returns:   True if all disks are moved properly, false if any errors are encountered.
    '               If there are any errors, all data entered into the database will be deleted.
    '-------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strMsg As String
    Dim rstConsumed As New ADODB.Recordset
    Dim X As Integer
    
    On Error GoTo errCopyDiskData
    
    dbconResultsData.BeginTrans
    'open blank recordset with right columns
    rstConsumed.Open "select * from ConsumedDisks WHERE serialnumber='1'", dbconResultsData, adOpenDynamic, adLockOptimistic
    
    With frmDiskQuantities.grdDiskNum
        For X = 1 To .Rows - 1
            strSQL = "SELECT * FROM AvailableDisks WHERE Disk_Number='" & .TextMatrix(X, 2) & "' and Batch_Number='" & .TextMatrix(X, 1) & "' ;"
            Set adoRS = dbconResultsData.Execute(strSQL)
            If adoRS.EOF Then
                strMsg = "Error trying to copy batch number " & .TextMatrix(X, 1) & " and disk number " & .TextMatrix(X, 2) & " from the AvailableDisks table to the ConsumedDisks table."
                MsgBox strMsg, vbInformation, "Error"
                ErrorLogger "adoCopyDiskData", strMsg      'this is a data error. log it so someone can follow up on it
                adoCopyDiskData = False
                Exit Function
            End If
            rstConsumed.AddNew
            rstConsumed.Fields("SerialNumber") = globCurrentUnit.SN
            rstConsumed.Fields("UnitNumber") = globCurrentUnit.UnitNumber
            rstConsumed.Fields("Batch_Number") = adoRS.Fields("Batch_Number")
            rstConsumed.Fields("Disk_Number") = adoRS.Fields("Disk_Number")
            rstConsumed.Fields("TimeTested") = adoRS.Fields("TimeTested")
            rstConsumed.Fields("TimeConsumed") = Now
            rstConsumed.Fields("SetupID") = adoRS.Fields("SetupID")
            rstConsumed.Fields("DiskType") = adoRS.Fields("DiskType")
            rstConsumed.Fields("Energy_1_Current") = adoRS.Fields("Energy_1_Current")
            rstConsumed.Fields("Energy_2_Current") = adoRS.Fields("Energy_2_Current")
            rstConsumed.Fields("Energy_3_Current") = adoRS.Fields("Energy_3_Current")
            rstConsumed.Fields("E1_Charge_Setpoint") = adoRS.Fields("E1_Charge_Setpoint")
            rstConsumed.Fields("E2_Charge_Setpoint") = adoRS.Fields("E2_Charge_Setpoint")
            rstConsumed.Fields("E3_Charge_Setpoint") = adoRS.Fields("E3_Charge_Setpoint")
            rstConsumed.Fields("V1mA_Volts") = adoRS.Fields("V1mA_Volts")
            rstConsumed.Fields("V1mA_Current") = adoRS.Fields("V1mA_Current")
            rstConsumed.Fields("Vref_Volts") = adoRS.Fields("Vref_Volts")
            rstConsumed.Fields("Watts_Watts") = adoRS.Fields("Watts_Watts")
            rstConsumed.Fields("Watts_Temp") = adoRS.Fields("Watts_Temp")
            rstConsumed.Fields("V10kA_Volts") = adoRS.Fields("V10kA_Volts")
            rstConsumed.Fields("V10kA_Current") = adoRS.Fields("V10kA_Current")
            rstConsumed.Fields("Test_Failed") = adoRS.Fields("Test_Failed")
            rstConsumed.Fields("Category") = adoRS.Fields("Category")
            rstConsumed.Fields("CntrHold") = adoRS.Fields("CntrHold")
            rstConsumed.Update
        Next X
    End With
    
    If rstConsumed.State > 0 Then
        rstConsumed.Close
    End If
    
    With frmDiskQuantities.grdDiskNum
        For X = 1 To .Rows - 1
            strSQL = "DELETE FROM AvailableDisks WHERE Disk_Number='" & .TextMatrix(X, 2) & "' and Batch_Number='" & .TextMatrix(X, 1) & "' ;"
            dbconResultsData.Execute (strSQL)
        Next
    End With
    
'DLimburg 27F3eb2009 modified following line, as the exact number of 3 and 5kV disks used is important once you get to the Arrester Tester Program
  '  strSQL = "INSERT INTO BuiltUnits VALUES('" & globCurrentUnit.SN & "','" & globCurrentUnit.UnitNumber & "','" & Now & "','" & frmDiskQuantities.m_NumDisksEntered & "'," & CInt(globCurrentUser) & ") ;"
  ' DLimburg 2/4/2010 add values for 3kV and 5kV calcined disk quanties which have been added to table.
    strSQL = "INSERT INTO BuiltUnits VALUES('" & globCurrentUnit.SN & "','" & globCurrentUnit.UnitNumber & "','" & Now & "','" & frmDiskQuantities.m_NumDisksEntered & "'," & CInt(globCurrentUser) & "," & globCurrentUnit.NumOf3kVDisks_Gray & "," & globCurrentUnit.NumOf5kVDisks_Gray & "," & globCurrentUnit.NumOf3kVDisks_Calcined & "," & globCurrentUnit.NumOf5kVDisks_Calcined & ");"
    dbconResultsData.Execute (strSQL)
    
    strSQL = "UPDATE WO_Info SET Unit" & globCurrentUnit.UnitNumber & "NumBuilt=Unit" & globCurrentUnit.UnitNumber & "NumBuilt + 1 WHERE WO_Number='" & globCurrentWO.strWO & "' and CatalogNumber='" & globCurrentWO.strCatNumber & "' ;"
    dbconResultsData.Execute (strSQL)
    
    dbconResultsData.CommitTrans
    
    adoCopyDiskData = True
    Exit Function
errCopyDiskData:
    strMsg = "There was an error while try to copy from the AvailableDisks table to the ConsumedDisks table. Unit cannot be built."
    MsgBox strMsg, vbInformation, "Error"
    ErrorLogger "adoCopyDiskData", strMsg      'this is a data error. log it so someone can follow up on it
    dbconResultsData.RollbackTrans
    adoCopyDiskData = False
    Exit Function
End Function

Public Function DeleteWO(strWO As String) As Boolean
    '------------------------------------------------------------------------------------------------
    'Created:   04/03    By:Jeff Pierce
    'Purpose:   To perform all of the database modifications necessary to delete an
    '                entire work order and all units already built for it.  Each built unit will
    '                be unbuilt and all disks returned to inventory.
    '   Inputs:     strSN - The serial Number of the unit to unbuild.
    '                  iUnit - the unit number (A-F) to unbuild.
    'Returns:   True if any existing units are unbuilt properly and all data deleted,
    '                false if any errors are encountered.
    '-------------------------------------------------------------------------------------------------
    Dim SQL As String, SQLsn As String, SQLBuilt As String
    Dim rstWO_Info As New ADODB.Recordset
    Dim rstWO_SN As New ADODB.Recordset
    Dim rstBuiltUnits As New ADODB.Recordset
    Dim strMsg As String
    Dim iButton As Integer
    Dim bResult As Boolean
    Dim strUnitLetter As String
    
    SQL = "SELECT * FROM WO_Info "
    SQL = SQL & "WHERE WO_Number = '" & strWO & "' "
    rstWO_Info.Open SQL, dbconResultsData, adOpenDynamic, adLockOptimistic
    If rstWO_Info.EOF = True Then
        'wo not listed in database so we can't delete it
        strMsg = "The Work order '" & strWO & "' does not exist in the database" & vbCrLf
        strMsg = strMsg & "and therefore can't be deleted."
        MsgBox strMsg
        DeleteWO = False
        Exit Function
    End If
    
    DeleteWO = True 'default to all OK
    'now find out which serial numbers are attached to WO
    SQLsn = "SELECT SerialNumber FROM WO_SerialNumbers "
    SQLsn = SQLsn & "WHERE WO_Number = '" & strWO & "' "
    rstWO_SN.Open SQLsn, dbconResultsData, adOpenDynamic, adLockOptimistic
    If rstWO_SN.EOF = False Then
        'there are SN's listed for the wo.
        'see if any are built.
        SQLBuilt = "SELECT * FROM BuiltUnits WHERE SerialNumber IN "
        SQLBuilt = SQLBuilt & "(" & SQLsn & ")"
        rstBuiltUnits.Open SQLBuilt, dbconResultsData, adOpenDynamic, adLockOptimistic
        If rstBuiltUnits.EOF = False Then
            'units are already built, query user on deleting units.
            strMsg = "The following Units have already been built for the work order." & vbCrLf
            strMsg = strMsg & "Do you wish to Unbuild all of these units?"
            Do While rstBuiltUnits.EOF = False
                strMsg = strMsg & vbCrLf & rstBuiltUnits!SerialNumber & ", Unit " & rstBuiltUnits!Unit
                rstBuiltUnits.MoveNext
            Loop
            iButton = MsgBox(strMsg, vbYesNo, "Unbuild Units?")
            If iButton = vbNo Then
                'user elected not to unbuild units.
                MsgBox "Units will not be unbuilt and Work order will not be deleted"
                DeleteWO = False
                Exit Function
            End If
            'if we reach here, then the user wishes to to unbuild all built units
            rstBuiltUnits.MoveFirst
            DeleteWO = True
            Do While rstBuiltUnits.EOF = False
                'convert unit number to letter
                strUnitLetter = Chr(rstBuiltUnits!Unit + 64)
                bResult = UnBuildUnit(rstBuiltUnits!SerialNumber, strUnitLetter)
                If bResult = False Then
                    DeleteWO = False
                End If
                rstBuiltUnits.MoveNext
            Loop
        End If
            
        If DeleteWO = True Then
            'units all unbuilt, now delete info from serial number table
            SQL = "DELETE FROM WO_SerialNumbers "
            SQL = SQL & "WHERE WO_Number = '" & strWO & "' "
            dbconResultsData.Execute SQL
            
            'delete wo
            SQL = "DELETE FROM WO_Info "
            SQL = SQL & "WHERE WO_Number = '" & strWO & "' "
            dbconResultsData.Execute SQL
        
            MsgBox "Work Order Successfully Deleted"
        End If
    End If

End Function

Public Function UnBuildUnit(strSN As String, strUnit As String) As Boolean
    '------------------------------------------------------------------------------------------------
    'Created:   04/03    By:Jeff Pierce
    'Purpose:   To perform all of the database modifications necessary when unbuilding
    '                an arrester.  This routine ASSUMES that the serial number and unit
    '                have been validated (but not necessarily checked to see if the unit exists)
    '   Inputs:     strSN - The serial Number of the unit to unbuild.
    '                  iUnit - the unit number (A-F) to unbuild.
    'Returns:   True if all unbuild functions happen properly, false if any errors are encountered.
    '               If there are any errors, all changes to the data will be un-done.
    '-------------------------------------------------------------------------------------------------
    Dim bResult As Boolean
    Dim dbconUnbuild As New ADODB.Connection
    Dim strMsg As String
    Dim rstAvail As New ADODB.Recordset
    Dim rstConsumed As New ADODB.Recordset
    Dim rstBuilt As New ADODB.Recordset
    Dim rstWO As New ADODB.Recordset
    Dim SQL As String
    Dim fld As Field
    Dim strFieldName As String
    Dim strWO As String
    
    'make sure that the unit was built
    bResult = CheckForBuiltUnit(strSN, strUnit)
    'if not exit routine with a False (because unit could not be unbuilt)
    If bResult = False Then
        strMsg = "Unit could not be unbuilt because it does not exist"
        MsgBox strMsg
        UnBuildUnit = False
        Exit Function
    End If
    
    If dbconUnbuild.State > 0 Then
        dbconUnbuild.Close
    End If
    
    'we are going to create a separate connection for working in this routine.
    'Since transaction commands are going to be used, we don't want them
    'to effect any database mods that may be made by other routines.
    dbconUnbuild.ConnectionString = dbconSetupData.ConnectionString
    dbconUnbuild.Open
    
    'MArk the begining of a block of related transactions to the DB
    dbconUnbuild.BeginTrans
    
    SQL = "SELECT * FROM ConsumedDisks "
    SQL = SQL & "WHERE SerialNumber = '" & strSN & "' "
    SQL = SQL & "AND UnitNumber = " & Asc(UCase(strUnit)) - 64 & " "
    rstConsumed.Open SQL, dbconUnbuild, adOpenDynamic, adLockOptimistic
    If rstConsumed.EOF = True Then
        strMsg = "No disk info was found for '" & strSN & "' , Unit " & strUnit & vbCrLf
        strMsg = strMsg & "Unit can not be unbuild automatically!"
        MsgBox strMsg
        ErrorLogger "UnbuildUnit", strMsg
        'undo any DB changes and cancel transaction block
        dbconUnbuild.RollbackTrans
        UnBuildUnit = False
        Exit Function
    End If
    
    'the select statement is configured to return an emtpy recordset that
    'has all of the needed filed for the table
    SQL = "SELECT * FROM AvailableDisks WHERE Batch_Number = '1' "
    rstAvail.Open SQL, dbconUnbuild, adOpenDynamic, adLockOptimistic
    
    '--copy all of the disks back to available disks
    'loop thru each disk found in consumeddisks for this unit
    Do While rstConsumed.EOF = False
        'make new record in rstavail
        rstAvail.AddNew
        For Each fld In rstAvail.Fields
            Select Case LCase(fld.Name)
                Case "msrpl_trans"
                    'any fields listed in this case should not be copied.
                Case Else
                    'anything not explictly excepted using the above case should be copied.
                    fld.Value = rstConsumed.Fields(fld.Name).Value
            End Select
        Next fld
        'all done copying fields
        On Error Resume Next
        rstAvail.Update
        If Err.Number <> 0 Then
            'there was a problem re-inserting the disk into available disks.
            'tell user what problem was & exit
            strMsg = "There was a problem copying data for disk " & rstConsumed!Disk_Number
            strMsg = strMsg & "of batch " & rstConsumed!Batch_Number & vbCrLf
            strMsg = strMsg & "Unit can not be unbuild automatically!"
            strMsg = strMsg & vbCrLf & Err.Number & " " & Err.Description
            GoTo QuitEarly
        End If
        On Error GoTo 0
        rstConsumed.MoveNext
    Loop
    
    'all disks copied OK.  delete the disks from Consumed Disks
    SQL = "DELETE  FROM ConsumedDisks "
    SQL = SQL & "WHERE SerialNumber = '" & strSN & "' "
    SQL = SQL & "AND UnitNumber = " & Asc(UCase(strUnit)) - 64 & " "
    dbconUnbuild.Execute SQL
    
    'decrement the number of units build for this unit type in the WO_Info Table
    strFieldName = "Unit" & Asc(UCase(strUnit)) - 64 & "NumBuilt"
    strWO = GetWONumber(strSN)
    SQL = "SELECT * FROM WO_Info WHERE WO_Number = '" & strWO & "' "
    rstWO.Open SQL, dbconUnbuild, adOpenDynamic, adLockOptimistic
    If rstWO.EOF = True Then
        'there's a problem becuase we can't find the WO for this unit (which we have
        'already confirmed was built.
        strMsg = "The WO '" & strWO & "'  was not found in the database.  The arrester with " & vbCrLf
        strMsg = strMsg & "S/N = '" & strSN & "', Unit " & strUnit & " was confirmed as already built"
        strMsg = strMsg & "Unit can not be unbuild automatically!"
        rstWO.Close
        GoTo QuitEarly
    End If
    
    If rstWO.Fields(strFieldName).Value > 0 Then
        rstWO.Fields(strFieldName).Value = rstWO.Fields(strFieldName).Value - 1
        rstWO.Update
    End If
    rstWO.Close
    
    'delete unit from built units table
    SQL = "DELETE FROM BuiltUnits "
    SQL = SQL & "WHERE SerialNumber = '" & strSN & "' "
    SQL = SQL & "AND Unit = " & Asc(UCase(strUnit)) - 64 & " "
    dbconUnbuild.Execute SQL
    
    'if all DB mods occurred properly, commit the transaction (make permanent in DB)
    dbconUnbuild.CommitTrans
    'return True because  unit sucessfully deleted
    UnBuildUnit = True
    Exit Function
    
QuitEarly:
    'if not all mods completed properly, Rollback transaction (undo all changes to
    'DB since BeginTrans
    Err.Clear
    MsgBox strMsg
    ErrorLogger "UnbuildUnit", strMsg
    'undo any DB changes and cancel transaction block
    dbconUnbuild.RollbackTrans
    'return false because it didn't work.  User should have been notified of specific error
    'when it occured.
    UnBuildUnit = False
    Exit Function
    
        
        
End Function

Public Function adoGetDiskData()
    Dim strSQL As String
    Dim strMsg As String
    Dim X As Integer
    Dim sglVref As Single, strTemp As String
    Dim bVrefFound As Boolean
    Dim DiskTypeHolder As String    ' Add March 4, 2010 to accomodate difference in how categories are handled for 5kV disks
    Dim CategoryHolder As Integer   ' Add March 4, 2010 to accomodate difference in how categories are handled for 5kV disks
    'if nothing sets it false, then all disks passed
    adoGetDiskData = True
    
    With frmDiskQuantities.grdDiskNum
        For X = 1 To .Rows - 1
            strSQL = "SELECT V1mA_Volts, Vref_Volts, V10kA_Volts, Category, DiskType FROM AvailableDisks WHERE Batch_Number='" & .TextMatrix(X, 1) & "' and Disk_Number='" & .TextMatrix(X, 2) & "' ;"
            Set adoRS = dbconSetupData.Execute(strSQL)
            If adoRS.EOF Then
                strSQL = "SELECT * FROM ConsumedDisks WHERE Batch_Number='" & .TextMatrix(X, 1) & "' and Disk_Number='" & .TextMatrix(X, 2) & "' ;"
                Set adoRS = dbconSetupData.Execute(strSQL)
                If adoRS.EOF Then
                    strMsg = "There is no data available for batch number " & .TextMatrix(X, 1) & " and disk number " & .TextMatrix(X, 2) & "."
                    MsgBox strMsg, vbInformation, "Error"
                    ErrorLogger "adoGetDiskData", strMsg      'this is a data error. log it so someone can follow up on it
                    adoGetDiskData = False
                Else
                    strMsg = "Batch number " & .TextMatrix(X, 1) & " and disk number " & .TextMatrix(X, 2) & " have already been used."
                    MsgBox strMsg, vbInformation, "Error"
                    ErrorLogger "adoGetDiskData", strMsg      'this is a data error. log it so someone can follow up on it
                    adoGetDiskData = False
                End If
            Else
                'do Vref first
                bVrefFound = False  'init to not found
                strTemp = Trim(adoRS.Fields("Vref_Volts"))
                If Len(strTemp) > 0 Then
                    'something was found in vref field
                    If IsNumeric(strTemp) Then
                        sglVref = CSng(strTemp)
                    Else
                        'not a number in the field
                        sglVref = 0
                    End If
                    If sglVref > 0 Then bVrefFound = True
                Else
                    sglVref = 0
                End If
                If bVrefFound = False Then
                    'need to check for the V1mA value
                    strTemp = Trim(adoRS.Fields("V1mA_Volts"))
                    If Len(strTemp) > 0 Then
                        'something was found in V1mA field
                        If IsNumeric(strTemp) Then
                            sglVref = CSng(strTemp) * (1 / (2 ^ 0.5)) 'mult v1ma by .707
                        Else
                            'not a number in the field
                            sglVref = 0
                        End If
                        If sglVref > 0 Then bVrefFound = True
                    Else
                        sglVref = 0
                    End If
                End If
                
                .TextMatrix(X, 3) = Format(sglVref, "0.00")
                .TextMatrix(X, 4) = Trim(adoRS.Fields("V10kA_Volts"))
   ' Add March 4, 2010 to accomodate difference in how categories are handled for 5kV disks
                DiskTypeHolder = Left(Trim(adoRS.Fields("DiskType")), 4)
                CategoryHolder = Int(Trim(adoRS.Fields("Category")))
                Select Case DiskTypeHolder
                    Case "AT52"
                        CategoryHolder = CategoryHolder + 1000
                    Case "AT53"
                    ' Hi's and Lo's are denoted with 3 digits and need to have a thousand added to them to achive the correct value
                    ' 500 is a ramdom number that all AT53 Hi's and Lo's fall below
                        If CategoryHolder < 500 Then
                            CategoryHolder = CategoryHolder + 1000
                        End If
                    Case "AT54"
                    ' Hi's and Lo's are denoted with 3 digits and need to have a thousand added to them to achive the correct value
                    ' 500 is a ramdom number that all AT54 Hi's and Lo's fall below
                        If CategoryHolder < 500 Then
                            CategoryHolder = CategoryHolder + 1000
                        End If
                End Select
                '.TextMatrix(X, 5) = Trim(adoRS.Fields("Category"))
                
                .TextMatrix(X, 5) = CategoryHolder
    ' end of add March 4, 2010
                .TextMatrix(X, 6) = Trim(adoRS.Fields("DiskType"))
            End If
        Next
    End With
End Function

Public Function adoCreateWorkOrder(strWO As String, strCatNum As String, _
                strSNs() As String) As Boolean
    '------------------------------------------------------------------------------------------------
    'Created:   05/01/03    By:Ray Raf
    'Purpose:   To perform all actions necessary to create a new work order
    '                in the Database.  Modified from Jeff's original to be
    '                independent of frmAddWO (So it can be used by Cooper if needed)
    '                It assumes that the database has already be check for the
    '                existence of the WO to create.
    'Inputs:    strWO - The WO Number to create
    '              strCatNumber - the catalog number of the arrester to be built for this WO.
    '              strSNs - an array of serial numbers to associate with this WO
    'Returns:   True if WO is created properly , false if any errors are encountered.
    '               If there are any errors, all data entered into the database will be deleted.
    '-------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strMsg As String
    Dim iLoop As Integer
    Dim bResult As Boolean
    Dim iNumSerials As Integer
    
    On Error GoTo errCreateWorkOrder
    
    '1st confirm that the serial numbers don't exist
    bResult = adoCheckForExistingSNsByList(strSNs)
    If bResult = True Then
        MsgBox "Work order Not Created!"
        adoCreateWorkOrder = False
        Exit Function
    End If
    
    'find out how many Serial Numbers associated with this WO
    iNumSerials = UBound(strSNs) + 1       'strSNs is a 0 Lbound array
    
    'create WO entry in main WO table
    strSQL = "INSERT INTO WO_Info "
    strSQL = strSQL & "VALUES('" & strWO & "','" & strCatNum & "','" & Now & "','" & iNumSerials & "','1','0','0','0','0','0','0') ;"
    dbconSetupData.Execute strSQL

    'insert each serial number into the serial number table
    For iLoop = 0 To iNumSerials - 1
        strSQL = "INSERT INTO WO_SerialNumbers VALUES('" & strWO & "','" & strSNs(iLoop) & "') ;"
        dbconSetupData.Execute strSQL
    Next
    
    adoCreateWorkOrder = True
    Exit Function

errCreateWorkOrder:
    strSQL = "DELETE FROM WO_Info WHERE WO_Number='" & strWO & "' ;"
    dbconSetupData.Execute strSQL
    strSQL = "DELETE FROM WO_SerialNumbers WHERE WO_Number='" & strWO & "' ;"
    dbconSetupData.Execute strSQL
    strMsg = "There was an error while saving work order " & strWO & " to the database." & vbCr & "Error Number: " & Err.Number & vbCr & "Error Description: " & Err.Description
    MsgBox strMsg, vbInformation, "Error"
    ErrorLogger "adoCreateWorkOrder", strMsg      'this is a data error. log it so someone can follow up on it
    adoCreateWorkOrder = False
End Function

Public Function LoadUnitInfo(ByRef ThisWO As WorkOrder, ByRef ThisUnit As Unit) As Boolean
    '------------------------------------------------------------------------------------------------
    '   Created:    05/23/03    By:Ray Rafalowski
    '   Purpose:   to load all of the info neccessary to print a unit report into
    '                    WorkOrder and Unit type variables.  It assumes the unit SN
    '                    and ID (1-6) are already loaded into ThisUnit
    '   Inputs:       ThisWO - the work order variable to load info to
    '                   ThisUnit - the unit variable to load data to.
    '-------------------------------------------------------------------------------------------------
    Dim bResult As Boolean
    Dim strMsg As String
    Dim SQL As String
    Dim rstBuiltUnit As New ADODB.Recordset
    
    '1st check to make sure unit was built
    bResult = CheckForBuiltUnit(ThisUnit.SN, CStr(ThisUnit.UnitNumber))
    If bResult = False Then
        strMsg = "No infomation exists in the database for the unit.  Data can't be loaded"
        MsgBox strMsg
        LoadUnitInfo = False
        Exit Function
    End If
        
    'get wo number first
    ThisWO.strWO = GetWONumber(ThisUnit.SN)
    If Len(ThisWO.strWO) = 0 Then
        LoadUnitInfo = False
        Exit Function
    End If
    
    'use wo to get catalog number
    ThisWO.strCatNumber = GetCatalogNumber(ThisWO.strWO)
    If Len(ThisWO.strCatNumber) = 0 Then
        LoadUnitInfo = False
        Exit Function
    End If
    
    'get rest of needed WO info
    With ThisWO
        .strType = getType(.strCatNumber)
        .strBuildPath = getBuildPath(.strCatNumber, CStr(.strType))
    End With
    
    'this gets the basic info for the arrester
    ParseCatNumToUnit ThisWO.strCatNumber, ThisUnit
    
    With ThisUnit
        '-----------now we need to get specific info for the unit.
        'Get number of disks
        SQL = "SELECT * FROM BuiltUnits "
        SQL = SQL & "WHERE SerialNumber = '" & .SN & "' "
        SQL = SQL & "AND Unit = " & .UnitNumber & " "

        rstBuiltUnit.Open SQL, dbconResultsData, adOpenStatic, adLockReadOnly
        If rstBuiltUnit.EOF = True Then
            MsgBox "S/N " & .SN & ", Unit " & .UnitNumber & " has not yet been built!"
            Exit Function
        End If
        globNumDisksToUse = rstBuiltUnit!NumDisks
        .NumOfDisks = rstBuiltUnit!NumDisks
 'DLimburg Change 2009_3 23Sept2009 retreived number of 3kV and 5kV disk that were used in build
        .NumOf3kVDisks_Gray = rstBuiltUnit!NumOf3kVDisks    ' gray disk use old name
        .NumOf5kVDisks_Gray = rstBuiltUnit!NumOf5kVDisks    ' gray disk use old name
        .NumOf3kVDisks_Calcined = rstBuiltUnit!NumOf3kVDisksCalcined
        .NumOf5kVDisks_Calcined = rstBuiltUnit!NumOf5kVDisksCalcined
        rstBuiltUnit.Close
        
        'get unit MCOV
        .UnitMCOV = GetUnitMCOV(ThisWO, .UnitNumber, ThisUnit)
        
        If ThisWO.strType = 2 Then  'polymer arrester
            'get stud, disk, and process for poly arrester  'modified 10/21/03
' DLimburg Change 2009_1 11Feb2009 Added .DiskType2 to proccedure call
            GetPolyStudDiskandProcess ThisWO.strCatNumber, .Rating, .Stud, .DiskType, .DiskType2, .Process
            If Len(Trim(.Stud)) < 1 Then
                'stud not found
                LoadUnitInfo = False
                Exit Function
            End If
            
            'get the sections (sheds) for the unit
            GetUnitSheds ThisUnit, ThisWO.strBuildPath, ThisWO.strCatNumber
        Else
            'porcelain arrester
            GetPorcDisk ThisUnit, ThisWO.strCatNumber, ThisWO.strBuildPath
        End If
        
    End With
    
    'now we should have everything that we need.
    LoadUnitInfo = True
    
End Function

Public Function GetWONumber(strSN As String) As String
    '------------------------------------------------------------------------------------------------
    '   Created:    05/23/03    By:Ray Rafalowski
    '   Purpose:   To get the work order number associated with a serial number.
    '   Inputs:       strSN - The arrester serial number to lookup the Work order for.
    '   Returns:    The work order number associated with the selected arrester.  If the serial number
    '                    is not found, a blank string will be returned.
    '-------------------------------------------------------------------------------------------------
    Dim rstWO As New ADODB.Recordset
    Dim SQL As String
    Dim strMsg As String
    
    SQL = "SELECT * FROM WO_SerialNumbers "
    SQL = SQL & "WHERE SerialNumber = '" & strSN & "' "
    
    rstWO.Open SQL, dbconResultsData, adOpenStatic, adLockReadOnly
    
    If rstWO.EOF = False Then
        'record found .  Load WO number
        GetWONumber = rstWO!WO_Number
    Else
        'This S/N doesn't exist in the database.  Either it dates before this system, It has
        'not yet been created, or it was purged from the online data because of it's age.
        GetWONumber = ""
        strMsg = "This S/N doesn't exist in the database.  Either it dates before this system, it has" & vbCrLf
        strMsg = strMsg & "not yet been created, or it was purged from the online data because of it's age." & vbCrLf
        strMsg = strMsg & vbCrLf & "No Work Order info found for Arrester with S/N of '" & strSN & "' "
        MsgBox strMsg
        Exit Function
    End If

End Function

Public Function GetCatalogNumber(strWO As String) As String
    '------------------------------------------------------------------------------------------------
    '   Created:    05/23/03    By:Ray Rafalowski
    '   Purpose:   To get the Catalog Number associated with a Work Order.
    '   Inputs:       strWO - The Work order to lookup the catalog number for.
    '   Returns:    The catalog number associated with the selected arrester.  If the WO number
    '                    is not found, a blank string will be returned.
    '-------------------------------------------------------------------------------------------------
    Dim rstCat As New ADODB.Recordset
    Dim SQL As String
    Dim strMsg As String
    
    SQL = "SELECT * FROM WO_Info "
    SQL = SQL & "WHERE WO_Number = '" & strWO & "' "
    
    rstCat.Open SQL, dbconResultsData, adOpenStatic, adLockReadOnly
    
    If rstCat.EOF = False Then
        'wo info was found.
        GetCatalogNumber = rstCat!CatalogNumber
        Exit Function
    Else
        strMsg = "No information for the work order '" & strWO & "' was found in the database"
        MsgBox strMsg
        GetCatalogNumber = ""
        Exit Function
    End If
    
End Function

Public Function CheckForBuiltUnit(ByVal strSN As String, ByVal strUnit As String) As Boolean
    '------------------------------------------------------------------------------------------------
    '   Created:    05/23/03    By:Ray Rafalowski
    '   Purpose:   To see if the S/N , Unit combination has been built and has data available.
    '   Inputs:       strSN - the serial number of the unit in question
    '                   strUnit - the Unit number (A-F) of the unit in question
    '   Returns:    True if data for the unit exists in the database, false otherwise
    '-------------------------------------------------------------------------------------------------
    Dim rstUnit As New ADODB.Recordset
    Dim SQL As String
    Dim iUnitNumber As Integer
    Dim strMsg As String
    
    'get rrid of any extra spaces
    strUnit = Trim(strUnit)
    
    If Len(strUnit) <> 1 Then
        'an invalid unit ID was passed.
        strMsg = "The Unit ID '" & strUnit & "' is invalid.  No check can be performed!"
        MsgBox strMsg
        CheckForBuiltUnit = False
        Exit Function
    End If
    
    If Not (IsNumeric(strUnit)) Then
        'get the ascii code for the unit ID passed
        iUnitNumber = Asc(UCase(strUnit))
        If iUnitNumber < 65 Or iUnitNumber > 70 Then
            strMsg = "The entered Unit code '" & strUnit & "' is outside the allowed range"
            MsgBox strMsg
            CheckForBuiltUnit = False
            Exit Function
        End If
            
        'we know that the unit is one of the allowable (A-F), convert to a number
        'in the range 1-6 by manipulating the ascii code for the unit ID
        iUnitNumber = iUnitNumber - 64
    Else
        'a number was passed, check it to see if it is 1 to 6
        iUnitNumber = CInt(strUnit)
        If iUnitNumber < 1 Or iUnitNumber > 6 Then
            strMsg = "The entered Unit code '" & strUnit & "' is outside the allowed range"
            MsgBox strMsg
            CheckForBuiltUnit = False
            Exit Function
        End If
    End If
    
    SQL = "SELECT * FROM BuiltUnits "
    SQL = SQL & "WHERE SerialNumber = '" & strSN & "' "
    SQL = SQL & "AND Unit = " & iUnitNumber & " "
    
    rstUnit.Open SQL, dbconResultsData, adOpenStatic, adLockReadOnly
    
    'if eof=true , no record found ie no data for unit
    CheckForBuiltUnit = Not (rstUnit.EOF)
    
    If rstUnit.State > 0 Then
        rstUnit.Close
    End If
    
End Function

Public Function GetUnitMCOV(WO As WorkOrder, UnitNumber As Integer, _
                            ThisUnit As Unit) As String
    '------------------------------------------------------------------------------------------------
    'Created:   05/23/03    By:Ray Rafalowski
    'Purpose:   To get the MCOV for the specific unit
    'Inputs:    WO - variable containing type & catalog number
    '               UnitNumber - the unit number (1-6) to get the MCOV for
    '               ThisUnit - a unit variable holding info about the unit to get
    '                               the unit MCOV for.
    'Returns:   the MCOV for the unit or 0 if unit is not valid for arrester.
    '   Rev'd:      10/22/03    By: Ray Raf
    '   Note:       changed logic to handle new and old process for MCOV table name
    '   Rev'd:      10/31/03    By: Ray Raf
    '   Note:       added ThisUnit variable to pass required info for MCOV lookup
    '                   when dealing with porcelain arresters.
    '-------------------------------------------------------------------------------------------------
    Dim strTemp As String
    Dim varMCOVTable  As String, varCompTable As String
    Dim SQL As String
    Dim TempUnit As Unit
    Dim strMCOVTable As String
    
    Select Case WO.strType
        Case 1 'porc
            Select Case LCase(WO.strBuildPath)
                Case "dual rated"
                    'dual rated arresters ALWAYS have exactly two units.
                    SQL = "SELECT * FROM MCOV_Dr_Porc "
                    SQL = SQL & "WHERE CatalogNumber='" & WO.strCatNumber & "' "
                    Set adoRS = dbconSetupData.Execute(SQL)
                    If Not adoRS.EOF Then
                        'the catalog number was found in the table
                        Select Case UnitNumber
                            Case 1
                                GetUnitMCOV = adoRS.Fields("MCOV_Unit_1")
                            Case 2
                                GetUnitMCOV = adoRS.Fields("MCOV_Unit_2_Lower")
                            Case Else
                                GetUnitMCOV = "0"
                        End Select
                    Else
                        'nothing found so the unit mcov has to be 0
                        GetUnitMCOV = "0"
                    End If
                Case "cubicle mount"
                    'cubicle mount can have up to 3 units
                    With ThisUnit
                        SQL = "SELECT * FROM MCOV_Cube "
                        SQL = SQL & "WHERE Family='" & .Family4Lookup & "' "
                        SQL = SQL & "AND ArresterRating='" & .Rating & "' "
                        SQL = SQL & "AND String_MCOV ='" & .MCOV & "' "
                        SQL = SQL & "AND PorcRingNumber='" & .PorcelainRingCode & "' ;"
                        Set adoRS = dbconSetupData.Execute(SQL)
                    End With
                    
                    If Not adoRS.EOF Then
                        If UnitNumber > 0 And UnitNumber < 4 Then
                            GetUnitMCOV = adoRS.Fields("MCOV_Unit_" & UnitNumber)
                        Else
                            GetUnitMCOV = "0"
                        End If
                    Else
                        GetUnitMCOV = "0"
                    End If
                Case Else
                    'all others can have up to 4 units
                    ' added test for calcined disk builds - 14Apr2011
                    If globCalcinedDisksUsedInPorc = True Then
                        strMCOVTable = "MCOV_Porc_Rev1"
                    Else
                        strMCOVTable = "MCOV_Porc"
                    End If
                    With ThisUnit
                        SQL = "SELECT * FROM " & strMCOVTable & " "
                        SQL = SQL & "WHERE Family='" & .Family4Lookup & "' "
                        SQL = SQL & "AND ArresterRating='" & .Rating & "' "
                        SQL = SQL & "AND string_MCOV='" & .MCOV & "' "
                        SQL = SQL & "AND PorcRingNumber='" & .PorcelainRingCode & "' ;"
                        Set adoRS = dbconSetupData.Execute(SQL)
                    End With
                    If Not adoRS.EOF Then
                        If UnitNumber > 0 And UnitNumber < 4 Then
                            GetUnitMCOV = adoRS.Fields("MCOV_Unit_" & UnitNumber)
                        Else
                            GetUnitMCOV = "0"
                        End If
                    Else
                        GetUnitMCOV = "0"
                    End If
            End Select
        Case 2 'poly
            Select Case LCase(WO.strBuildPath)
                Case "dual rated"
                    SQL = "SELECT * FROM MCOV_Dr_Poly "
                    SQL = SQL & "WHERE CatalogNumber = '" & WO.strCatNumber & "' ;"
                    Set adoRS = dbconSetupData.Execute(SQL)
                    If Not adoRS.EOF Then
                        Select Case UnitNumber
                            Case 1
                                GetUnitMCOV = adoRS.Fields("MCOV_Unit_1")
                            Case 2
                                GetUnitMCOV = adoRS.Fields("MCOV_Unit_2_Lower")
                            Case Else
                                GetUnitMCOV = "0"
                        End Select
                    Else
                        GetUnitMCOV = "0"
                    End If
                Case "1in", "12mm"
                    'don't have process yet
                    With TempUnit
                        .UnitNumber = UnitNumber
                        ParseCatNumToUnit WO.strCatNumber, TempUnit
                        GetPolyTableNames TempUnit, varCompTable, varMCOVTable
    '                    If WO.strBuildPath = "1in" Then
    '                        varMCOVTable = "MCOV_1in"
    '                    Else
    '                        varMCOVTable = "MCOV_12mm"
    '                    End If
                        SQL = "SELECT * FROM " & varMCOVTable & " "
                        SQL = SQL & "WHERE Family='" & .Family & "' "
                        SQL = SQL & "AND String_MCOV='" & .MCOV & "' "
                        SQL = SQL & "AND ShedCode='" & .PolyShedRingCode & "' "
                        SQL = SQL & "AND ArresterRating = " & .Rating & " "
                        Set adoRS = dbconSetupData.Execute(SQL)
                    End With
                    If Not adoRS.EOF Then
                        'arrester was found.  Get Unit data
                        If UnitNumber > 0 And UnitNumber < 7 Then
                            GetUnitMCOV = adoRS.Fields("MCOV_Unit_" & UnitNumber)
                        Else
                            'arrester not found
                            GetUnitMCOV = "0"
                        End If
                    Else
                        'invalid unit number
                        GetUnitMCOV = "0"
                    End If
            End Select
    End Select
    If adoRS.State = 1 Then
        adoRS.Close
    End If
End Function


Public Sub GetPolyStudDiskandProcess(ByVal strCatNum As String, ByVal Rating As Integer, _
                                ByRef strStud As String, ByRef strDiskType As String, _
                                ByRef strDiskType2 As String, ByRef strProcess As String)
    '------------------------------------------------------------------------------------------------
    '   Created:    05/23/03    By:Ray Rafalowski
    '   Purpose:   To get the stud and disk size based on family and rating
    '   Inputs:       strCatNum - the complete catalog number for the arrester
    '                   Rating - the rating of the entire arrester
    '   Outputs:   strStud - The stud string as it is stored in the table
    '                   strDiskType - the disk type stored in the table
    ' DLimburg Change 2009_1 11Feb2009 Added strDiskType2 to proccedure call
    '                   strDiskType2 - the disk type stored in the table
    '                   strProcess - the process stored in the table
    '   Rev'd:      10/21/03    by: Ray Raf
    '   Note:        Changed sub name and added handling for 3 different Nameplate
    '                   tables.
    '   Rev'd:      10/7/05     by: DLimburg
    '   Noted:      Added test for cubicle mount options F,G,J,K & L
    '-------------------------------------------------------------------------------------------------
    Dim SQL As String
    Dim rstNameplate As New ADODB.Recordset
    Dim Family As String
    Dim strNameplateTableName As String
    
    Family = getFamily(2, strCatNum)
    
    If Mid(strCatNum, 3, 1) = "D" Then
        'dual rated arrester
        SQL = "SELECT * FROM NameplateDrPoly "
        SQL = SQL & "WHERE CatalogNumber='" & strCatNum & "' "
    Else
        'either cubicle mount or normal.  Either way need family
        Family = getFamily(2, strCatNum)
        'now chekc to see if it is a cube mount arrester
        '10/7/05  added test for cubicle mount options F,G,J,K & L
        If Mid(strCatNum, 13, 1) = "B" Or _
            Mid(strCatNum, 13, 1) = "F" Or _
            Mid(strCatNum, 13, 1) = "G" Or _
            Mid(strCatNum, 13, 1) = "J" Or _
            Mid(strCatNum, 13, 1) = "K" Or _
            Mid(strCatNum, 13, 1) = "L" Then
            'it's a cube mount alright!
            strNameplateTableName = "NameplateCubePoly"
        Else
            'just a normal poly arrester (be nice if they were all like this)
            strNameplateTableName = "NameplatePoly"
        End If
        'setup query string
        SQL = "SELECT * FROM " & strNameplateTableName & " "
        SQL = SQL & "WHERE Family='" & Family & "' "
        SQL = SQL & "AND MinRating <= '" & Rating & "' "
        SQL = SQL & "AND MaxRating >= '" & Rating & "' "
    End If
    
    'now everythings the same for all arresters
    rstNameplate.Open SQL, dbconSetupData, adOpenDynamic, adLockBatchOptimistic
    If rstNameplate.EOF Then
        MsgBox "The disk type and stud could not be retrieved for this arrester."
        Exit Sub
    End If

    strDiskType2 = GetPartnerDisk(DiskType)
    strStud = Trim(rstNameplate!Stud)
    strDiskType = Trim(rstNameplate!DiskType)
    strProcess = Trim(rstNameplate!Process)
    

' DLimburg Change 2009_1 11Feb2009 added code to close existing recordset, should not cause a problem farther on
' in code becaue recordset defined with this subroutine
    If rstNameplate.State = 1 Then
        rstNameplate.Close
    End If
    
    Set rstNameplate = Nothing
    
End Sub

Public Function GetPartnerDisk(DiskType As String) As String
    Dim rstDiskPartner As New ADODB.Recordset
    '  get 5kV partner disk
    
        SQL = "SELECT * FROM DiskPartners "
        SQL = SQL & " WHERE DiskPartNo_3kV = '" & DiskType & "'"
        
        rstDiskPartner.Open SQL, dbconSetupData, adOpenDynamic, adLockBatchOptimistic
        
        If rstDiskPartner.EOF = True Then
            strDiskType2 = ""
        Else
            strDiskType2 = Trim(rstDiskPartner!DiskPartNo_5kv)
        End If

        If rstDiskPartner.State = 1 Then
            rstDiskPartner.Close
        End If
        Set rstDiskPartner = Nothing
        
Exit Sub
ErrorFindingPartnerDisk:
    strDiskType2 = ""
End Function

Public Sub GetUnitSheds(ThisUnit As Unit, strBuildPath As String, strCatNumber As String)
    '------------------------------------------------------------------------------------------------
    'Created:   05/23/03    By:Ray Rafalowski
    'Purpose:   To get the number of sheds for the given polymer unit
    '                This logic assums that the stud for the unit has already been determined
    'Inputs:       ThisUnit - a unit variable that contains the information about the unit
    '                 strBuildPath - the build path for the unit so we don't have to recreate it
    '                 strCatNumber - The catalog number - needed only for dual rated arresters
    'Outputs:    this routine places the Shed count for this unit in .UnitShed of ThisUnit
    '   Rev'd:      10/22/03    By: Ray Raf
    '   Note:       added handling of old and new process for MCOV table
    '-------------------------------------------------------------------------------------------------
    Dim rstHousing As New ADODB.Recordset
    Dim SQL As String
    Dim strFieldName As String
    Dim strMCOVTable As String, strCompTable As String
    
    'determine which MCOV table to use based on the stud of the unit
    'the stud value is valid for dual rated arresters becuase they can be either
    '12mm or 1 in.  the build path would not give the proper results
    
    'we have process info at this point
    GetPolyTableNames ThisUnit, strCompTable, strMCOVTable
'    If ThisUnit.Stud = "12mm" Then
'        strMCOVTable = "MCOV_12mm"
'    Else
'        strMCOVTable = "MCOV_1in"
'    End If
    
    'set query string to ready data from proper table for unit
    Select Case strBuildPath
        Case "Dual Rated"
            strFieldName = "Unit_" & ThisUnit.UnitNumber & "_Housing"
            SQL = "SELECT " & strFieldName & " "
            SQL = SQL & "FROM HousingDrPoly WHERE CatalogNumber='" & strCatNumber & "' "
        Case Else
            strFieldName = "Housing_" & ThisUnit.UnitNumber
            SQL = "SELECT " & strFieldName & " FROM " & strMCOVTable
            SQL = SQL & " WHERE Family= '" & ThisUnit.Family & "' "
            SQL = SQL & "AND string_MCOV='" & ThisUnit.MCOV & "' "
            SQL = SQL & "AND ShedCode='" & ThisUnit.PolyShedRingCode & "' "
            SQL = SQL & "AND ArresterRating = " & ThisUnit.Rating & " "
    End Select


    rstHousing.Open SQL, dbconSetupData, adOpenDynamic, adLockBatchOptimistic
    'sheck to make sure a housing was found
    If rstHousing.EOF = True Then
        MsgBox "No housing was found for this unit!.", vbInformation, "Error"
        Exit Sub
    End If

    ThisUnit.UnitSections = CInt(Trim(rstHousing.Fields(strFieldName).Value))

    If rstHousing.State > 0 Then
        rstHousing.Close
    End If

End Sub

Public Function GetPorcDisk(ThisUnit As Unit, strCatNumber As String, strBuildPath As String) As String
    Dim rstDiskType As New ADODB.Recordset
    Dim SQL As String
    
    '--determine the diskType based on the build path
    Select Case strBuildPath
        Case "Dual Rated"
            SQL = "SELECT DiskType FROM MCOV_Dr_Porc WHERE CatalogNumber='" & strCatNumber & "' ;"
        Case "Cubicle Mount"
            SQL = "Select DiskType FROM MCOV_Cube WHERE Family='" & ThisUnit.Family4Lookup & "' "
            SQL = SQL & "AND String_MCOV='" & ThisUnit.MCOV & "' AND ArresterRating=" & ThisUnit.Rating & " "
            SQL = SQL & "AND PorcRingNumber = " & ThisUnit.PorcelainRingCode & " ;"
        Case Else
            SQL = "Select DiskType FROM MCOV_Porc_Rev1 WHERE Family='" & ThisUnit.Family4Lookup & "' "   ' Changed table name from MCOV_Porc to MCOV_Porc_Rev1 DLimburg 21Feb2011
            SQL = SQL & "AND String_MCOV='" & ThisUnit.MCOV & "' AND ArresterRating=" & ThisUnit.Rating & " "
            SQL = SQL & "AND PorcRingNumber = '" & ThisUnit.PorcelainRingCode & "' ;"
    End Select
    
    rstDiskType.Open SQL, dbconSetupData, adOpenStatic, adLockReadOnly
    'since the catalog number was verified from the MCOV tables before getting her, there will be a DiskType
    ThisUnit.DiskType = rstDiskType.Fields("DiskType")
    ThisUnit.DiskType2 = GetPartnerDisk(ThisUnit.DiskType)  ' added routine to get 5kV disk part number
End Function

Public Function LookupNickname(strNickname As String) As String
    '------------------------------------------------------------------------------------------------
    'Created:   10/21/03    By:Ray Rafalowski
    'Purpose:   to cross reference a catalog number 'nickname' to a full catalog number
    'Inputs:       strNickname - the nickname to look up
    'Outputs:    the catalog number for the nickname if found, a blank string if
    '                 nickname not found
    '-------------------------------------------------------------------------------------------------
    Dim rstNickname As New ADODB.Recordset
    Dim SQL As String
    
    On Error GoTo ErrorHandler
   'DLimburg 2Feb2009 changed table name to be consistent with othere programs and eliminate
   ' duplicate SQL tables
   ' SQL = "SELECT * FROM Nicknames WHERE Nickname = '" & strNickname & "' "
    SQL = "SELECT * FROM ShortCutListing WHERE ShortCut = '" & strNickname & "' "
    rstNickname.Open SQL, dbconSetupData, adOpenStatic, adLockReadOnly
    If rstNickname.EOF = True Then
        'nickname not found
        LookupNickname = ""
    Else
        LookupNickname = rstNickname!CatNo
    End If
    
    'close recordset
    If rstNickname.State > 0 Then
        rstNickname.Close
    End If
    Exit Function
ErrorHandler:
    'mark lookup as failed
    LookupNickname = ""
    ErrorLogger "lookupNickname", Err.Description
    Err.Clear
    If rstNickname.State > 0 Then
        rstNickname.Close
    End If

End Function


Public Function Test4CalcinedDiskInPorc(SN As String, UnitNumber As Integer) As Boolean
' this routine created to test if calcined disks have been used in the currect arrester.
Dim SQL As String
Dim rstBuiltUnit As New ADODB.Recordset

        SQL = "SELECT * FROM BuiltUnits "
        SQL = SQL & "WHERE SerialNumber = '" & SN & "' "
        SQL = SQL & "AND Unit = " & UnitNumber & " "

        rstBuiltUnit.Open SQL, dbconResultsData, adOpenStatic, adLockReadOnly
        If rstBuiltUnit.EOF = False Then
            If rstBuiltUnit!NumOf3kVDisks = 0 And _
                rstBuiltUnit!NumOf5kVDisks = 0 And _
                (rstBuiltUnit!NumOf3kVDisksCalcined > 0 Or _
                 rstBuiltUnit!NumOf5kVDisksCalcined > 0) Then
                 globCalcinedDisksUsedInPorc = True ' set gllobal veriable that arrester is a calcined build
                 Test4CalcinedDiskInPorc = True ' valid answer
            Else
                If (rstBuiltUnit!NumOf3kVDisks > 0 Or _
                    rstBuiltUnit!NumOf5kVDisks > 0) And _
                    (rstBuiltUnit!NumOf3kVDisksCalcined = 0 And _
                     rstBuiltUnit!NumOf5kVDisksCalcined = 0) Then
                     globCalcinedDisksUsedInPorc = False ' set gllobal veriable that arrester is NOT a calcined build
                     Test4CalcinedDiskInPorc = True ' valid answer
                Else
                    MsgBox ("Error has occurred, this build contains a mix of gray and calcined disk. Do not proceed with build.")
                    Test4CalcinedDiskInPorc = False ' list build is not valid for a porcelain arrester, if this error happens ther is a mistake in the code that allows a disk build
                End If
            End If
        Else
            Test4CalcinedDiskInPorc = False ' serial number not found, should not get to this point in code
        End If
End Function

