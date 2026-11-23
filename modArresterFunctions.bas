Attribute VB_Name = "modArresterFunctions"
Option Explicit
    '---------------------------MODULE COMMENT--------------------------------------------------------------
    '   This module should contain code that handles extracting data from the
    '   catalog number.  This module is common with both the tester and
    '   build programs.
    '------------------------------------------------------------------------------------------------

Public Type WorkOrder
    strWO As String
    strCatNumber As String
    strType As Integer
    strBuildPath As String
    strSN() As String
    numSN As Integer
    bArrestComp() As String
    globNumDiskDefault As Integer
    SerialNumbersChecked As Integer
End Type

Public Type Unit
    Family  As String
    Nameplate   As String
    Rating  As Integer
    MCOV    As String   '3 chars
    PorcelainRingCode   As String
    PorcelainHousing    As String 'B or G
    PolyShedRingCode    As String
    Family4Lookup   As String
    Type    As Integer      '(1=Porc., 2=Poly)
    DesignMCOV As Single
    SN As String    'serial number
    UnitNumber As Integer '(1 to 6)
    UnitMCOV As String      'can be used for lookups using MCOV as a Key
    UnitSections As Integer     '   the number of poly shed sections for the unit
    UnitEndElectrodeCount As Integer
    UnitMidElectrodeCount As Integer
    UnitHeatSinkCount As Integer
    UnitORingCount As Integer
    HousingCode As String
    DiskType As String          'The disk type to use as a lookup
    DiskType2 As String          'The disk type to use as a lookup
    Stud As String                 'indicates the stud size to use in polymer lookups
    Process As String           'indicates whether to use '_old' or '_new' tables for polymer lookups
    NumOfDisks As Integer       'number of disks to be used in unit
    NumOf3kVDisks As Integer    ' number of 3kV disks used in design
    NumOf5kVDisks As Integer    ' number of 5kV disks used in design
    NumOf3kVDisks_Gray As Integer    ' number of 3kV disks used in design ' DLimburg Change 2010_1 3Feb2010 Modified and added code to allow for the 3 and 5 calcined disks
    NumOf5kVDisks_Gray As Integer    ' number of 5kV disks used in design ' DLimburg Change 2010_1 3Feb2010 Modified and added code to allow for the 3 and 5 calcined disks
    NumOf3kVDisks_Calcined As Integer    ' number of 3kV disks used in design ' DLimburg Change 2010_1 3Feb2010 Modified and added code to allow for the 3 and 5 calcined disks
    NumOf5kVDisks_Calcined As Integer    ' number of 5kV disks used in design ' DLimburg Change 2010_1 3Feb2010 Modified and added code to allow for the 3 and 5 calcined disks

End Type
Public globCalcinedDisksUsedInPorc As Boolean ' Set to True is calcined disks are used in a porcelain arrester

Public Sub ParseCatNumToUnit(strCatNumber As String, ThisUnit As Unit)
    '------------------------------------------------------------------------------------------------
    'Created:   04/18/03    By:Jeff Pierce
    'Purpose:   parse the catalog number into it's components and put the
    '                data into the passed unit variable (typically globCurrentUnit)
    '                This routine ASSUMES! that the catalog number has already been verified.
    'Inputs:    strCatNumber - the catalog number to parse.
    '              ThisUnit - a Unit type variable to put the data in
    '   Rev'd:      10/23/03    By: Ray Raf
    '   Note:        added code for placing type in Unit for all arresters.
    '                   added code for placing stud, disk, and process in unit for poly arresters
    '-------------------------------------------------------------------------------------------------
    Dim iType As Integer
    Dim strBuildPath As String
    
    iType = getType(strCatNumber)
    strBuildPath = getBuildPath(strCatNumber, iType)
    Select Case iType
        Case 1  'porc
            
            Select Case LCase(strBuildPath)
                Case "cubicle mount", "aze/azg"
                    With ThisUnit
                        .Rating = getRating(iType, strCatNumber)
                        .Family = getFamily(iType, strCatNumber)
                        .Family4Lookup = PorcCrossRef(.Family, .Rating)
                        .MCOV = getMCOV(iType, strCatNumber)
                        .HousingCode = getHousing(iType, strCatNumber)
                        .Nameplate = getNameplate(iType, strCatNumber)
                        .PorcelainRingCode = getRingCode(iType, strCatNumber)
                        .Type = iType
                    End With
                Case "dual rated"
                    With ThisUnit
                        .Rating = getRating(iType, strCatNumber)
                        .Family = getFamily(iType, strCatNumber)
                        .HousingCode = getHousing(iType, strCatNumber)
                        .MCOV = getMCOV(iType, strCatNumber)
                        .Nameplate = getNameplate(iType, strCatNumber)
                        .Type = iType
                    End With
            End Select
        Case 2 'polymer arresters
            Select Case LCase(strBuildPath)
                Case "1in", "12mm"
                    With ThisUnit
                        .Family = getFamily(iType, strCatNumber)
                        .Rating = getRating(iType, strCatNumber)
                        .MCOV = getMCOV(iType, strCatNumber)
                        .PolyShedRingCode = getRingCode(iType, strCatNumber)
                        .Type = iType
' DLimburg Change 2009_1 11Feb2009 Added .DiskType2 to proccedure call
                        GetPolyStudDiskandProcess strCatNumber, .Rating, .Stud, .DiskType, .DiskType2, .Process
                    End With
                Case "dual rated"
                    With ThisUnit
                        .Family = getFamily(iType, strCatNumber)
                        .Rating = getRating(iType, strCatNumber)
                        .MCOV = getMCOV(iType, strCatNumber)
                        .Type = iType
' DLimburg Change 2009_1 11Feb2009 Added .DiskType2 to proccedure call
                        GetPolyStudDiskandProcess strCatNumber, .Rating, .Stud, .DiskType, .DiskType2, .Process
                    End With
            End Select
    End Select
End Sub

Public Function getFamily(iType As Integer, strCatNumber As String) As String
    '------------------------------------------------------------------------------------------------
    'Created:   03/28/03    By:Jeff Pierce
    'Purpose:   To get the family from the catalog number
    'Inputs:    strCatNumber - the catalog number to get the family from.
    '              iType - the arrester type (1=Porc., 2=Polymer)
    'Outputs:  the family string
    '-------------------------------------------------------------------------------------------------
        If iType = 1 Then
            getFamily = Mid(strCatNumber, 1, 4)
        Else
            getFamily = Mid(strCatNumber, 1, 2)
        End If
    
End Function

Public Function getRating(iType As Integer, strCatNumber As String) As Integer
    '------------------------------------------------------------------------------------------------
    'Created:   03/28/03    By:Jeff Pierce
    'Purpose:   To get the Arrester Rating from catalog number
    'Inputs:    strCatNumber - the catalog number to get the rating from.
    '              iType - the arrester type (1=Porc., 2=Polymer)
    'Outputs:  the rating string
    '-------------------------------------------------------------------------------------------------
    Dim strNameplate As String
    
    If iType = 1 Then   'porc
        strNameplate = getNameplate(iType, strCatNumber)
        If strNameplate = "D" Then
            'its a dual rated arrester, rating is in a different spot.
            getRating = Mid(strCatNumber, 9, 3)
        Else
            getRating = Mid(strCatNumber, 12, 3)
        End If
    Else    'poly
        strNameplate = Mid(strCatNumber, 3, 1)
        If strNameplate = "D" Then
            'it's a dual rated arrester.  rating is in a different spot
            getRating = Mid(strCatNumber, 7, 3)
        Else
            getRating = Mid(strCatNumber, 3, 3)
        End If
    End If
End Function

Public Function getNameplate(iType As Integer, strCatNumber As String) As String
    '------------------------------------------------------------------------------------------------
    'Created:   03/28/03    By:Jeff Pierce
    'Purpose:   To get the Nameplate value from catalog number (porc Only)
    'Inputs:    strCatNumber - the catalog number to get the Nameplate from.
    '              iType - the arrester type (1=Porc., 2=Polymer)
    'Outputs:  the nameplate string or 'N/A' for poly arresters
    '-------------------------------------------------------------------------------------------------
    If iType = 1 Then
        getNameplate = Mid(strCatNumber, 5, 1)
    Else
        'shouldn't have gotten here with a poly arrester type.  This could cause things to blow up
        getNameplate = "N/A"
    End If

End Function

Public Function getRingCode(iType As Integer, strCatNumber As String) As String
    '------------------------------------------------------------------------------------------------
    'Created:   03/28/03    By:Jeff Pierce
    'Purpose:   To get the Ring Code (porc) or shed code (poly) value from catalog number
    'Inputs:    strCatNumber - the catalog number to get the Nameplate from.
    '              iType - the arrester type (1=Porc., 2=Polymer)
    'Outputs:  the ring code string
    '-------------------------------------------------------------------------------------------------
    
        'everything but dual rated arresters have a ring or shed code embedded in the cat number
        If iType = 1 Then
            'porcelain
            If LCase(Mid(strCatNumber, 5, 1)) <> "d" Then
                'everything but dual rated arresters have a ring or shed code embedded in the cat number
                getRingCode = Mid(strCatNumber, 6, 2)
            Else
                getRingCode = "0"
            End If
        Else
            If LCase(Mid(strCatNumber, 3, 1)) <> "d" Then
                'everything but dual rated arresters have a ring or shed code embedded in the cat number
                getRingCode = Mid(strCatNumber, 9, 2)
            Else
                getRingCode = "0"
            End If
        End If
End Function

Public Function getHousing(iType As Integer, strCatNumber As String) As String
    '------------------------------------------------------------------------------------------------
    'Created:   03/28/03    By:Jeff Pierce
    'Purpose:   To get the Housing code (B or G) from catalog number (porc Only)
    'Inputs:    strCatNumber - the catalog number to get the Housing code from from.
    '              iType - the arrester type (1=Porc., 2=Polymer)
    'Outputs:  the nameplate string or 'N/A' for poly arresters
    '-------------------------------------------------------------------------------------------------
    Dim strBuildPath As String
    
    If iType = 1 Then   'porc
        strBuildPath = getBuildPath(strCatNumber, iType)
        If LCase(strBuildPath) = "dual rated" Then
            getHousing = Mid(strCatNumber, 12, 1)
        Else
            getHousing = Mid(strCatNumber, 8, 1)
        End If
    Else
        getHousing = "N/A"
    End If
End Function

Public Function getMCOV(iType As Integer, strCatNumber As String)
    '------------------------------------------------------------------------------------------------
    'Created:   03/28/03    By:Jeff Pierce
    'Purpose:   To get the String_MCOV value from catalog number
    'Inputs:    strCatNumber - the catalog number to get the string MCOV from.
    '              iType - the arrester type (1=Porc., 2=Polymer)
    'Outputs:  the string_MCOV value
    '-------------------------------------------------------------------------------------------------
    Dim strBuildPath As String
    
    strBuildPath = getBuildPath(strCatNumber, iType)
    If iType = 1 Then   'porc
        If LCase(strBuildPath) = "dual rated" Then
            getMCOV = Mid(strCatNumber, 6, 3)
        Else
            getMCOV = Mid(strCatNumber, 9, 3)
        End If
    Else
        'polymer arrester
        If LCase(strBuildPath) = "dual rated" Then
            getMCOV = Mid(strCatNumber, 4, 3)
        Else
            'normal poly arrester
            getMCOV = Mid(strCatNumber, 6, 3)
        End If
    End If
End Function

Public Function getType(strCatNumber As String) As Integer
    '------------------------------------------------------------------------------------------------
    'Created:   03/28/03    By:Jeff Pierce
    'Purpose:   To get the arrester type (1=Porc, 2 = Polymer
    'Inputs:    strCatalogNumber - the catalog number to get the type for.
    'Outputs:  the type or -1 if doesn't match either type
    '-------------------------------------------------------------------------------------------------
        
        Select Case UCase(Left(strCatNumber, 1))
            Case "A": getType = 1 'porc
            Case "U": getType = 2 'poly
            Case Else: getType = -1
        End Select
End Function

Public Function getTypeName(iType As Integer) As String
    '------------------------------------------------------------------------------------------------
    'Created:   03/28/03    By:Jeff Pierce
    'Purpose:   To get the arrester type as a string
    'Inputs:    iType - the numeric type to convert to a string
    'Outputs:  the arrester type as a string
    '-------------------------------------------------------------------------------------------------
    If iType = 1 Then
        getTypeName = "Porcelain"
    Else
        getTypeName = "Polymer"
    End If
End Function

'--return the build path
Public Function getBuildPath(strCatNumber As String, iType As Integer) As String
    '------------------------------------------------------------------------------------------------
    '   Created:   03/28/03    By:Jeff Pierce
    '   Purpose:   To get the build path
    '   Inputs:    strCatNumber - the catalog number to get the build path from .
    '                  iType - the arrester type (1=Porc., 2=Polymer)
    '   Outputs:  the build path name or "Not Found" if no build path could be determined.
    '   Rev'd:      10/22/03    By: Ray RAf
    '   Note:       added handling for separate polymer cubicle nameplate table
    '   Rev'd:      10/7/05     by: DLimburg
    '   Noted:      Added test for cubicle mount options F,G,J,K & L
    '-------------------------------------------------------------------------------------------------
    Dim strRingCode As String
    Dim strFamily As String
    Dim iRating As Integer
    Dim SQL As String
    Dim rstStud As New ADODB.Recordset
    Dim strStud As String
    Dim strNameplateTable As String
    
    'set default to path not found
    getBuildPath = "Not Found"
    Select Case iType
        Case 1  'porc
            'every legit porc arrester will fall into this 1st grouping
            If Mid(strCatNumber, 1, 3) = "AZE" Or Mid(strCatNumber, 1, 3) = "AZG" Then
                getBuildPath = "AZE/AZG"
            End If
            
            If getNameplate(iType, strCatNumber) = "D" Then
                'this will over-ride generic AZE/AZG path
                getBuildPath = "Dual Rated"
            End If
                            
            strRingCode = getRingCode(iType, strCatNumber)
            Select Case strRingCode
                Case "91", "92", "93", "94", "95"
                    'If strRingCode >= 91 And strRingCode <= 95 Then
                    'this will over-ride generic AZE/AZG path
                    getBuildPath = "Cubicle Mount"
                    'End If
            End Select
            'if none of the above matches were made, the default value will be returned
        Case 2      'poly
            'polymer build cases were modified 10/21/03 to reflect addition of cubicle
            'mount polymer arresters and new/old processes.  Dual rated nameplate
            'table now has stud info in it even though (at this time) all dual rated poly
            'arresters will be 12mm.
            If Mid(strCatNumber, 3, 1) = "D" Then
                'it's dual rated
                getBuildPath = "Dual Rated"
                'added additional lookups 10/21/03
                SQL = "SELECT * FROM NameplateDrPoly "
                SQL = SQL & "WHERE CatalogNumber = '" & strCatNumber & "' "
                rstStud.Open SQL, dbconSetupData, adOpenDynamic, adLockReadOnly
            Else
                'need to look in Nameplate table to determine if arrester is 1in or 12mm
                '10/7/05 Added test for cubicle mount options F,G,J,K & L
                strFamily = getFamily(iType, strCatNumber)
                iRating = Int(getRating(iType, strCatNumber))
                If Trim(Mid(strCatNumber, 13, 1)) = "B" Or _
                    Trim(Mid(strCatNumber, 13, 1)) = "F" Or _
                    Trim(Mid(strCatNumber, 13, 1)) = "G" Or _
                    Trim(Mid(strCatNumber, 13, 1)) = "J" Or _
                    Trim(Mid(strCatNumber, 13, 1)) = "K" Or _
                    Trim(Mid(strCatNumber, 13, 1)) = "L" Then
                    'cubicle mount
                    strNameplateTable = "NameplateCubePoly"
                Else
                    'normal arrester
                    strNameplateTable = "NameplatePoly"
                End If
                SQL = "SELECT * FROM " & strNameplateTable & " "
                SQL = SQL & "WHERE Family = '" & strFamily & "' "
                SQL = SQL & "AND MinRating <= " & iRating & " "
                SQL = SQL & "AND MaxRating >= " & iRating & " "
                rstStud.Open SQL, dbconSetupData, adOpenDynamic, adLockReadOnly
                If rstStud.EOF = False Then
                    'a record was returned
                    strStud = Trim(rstStud!Stud)
                    If InStr(1, strStud, "12mm") > 0 Then
                        getBuildPath = "12mm"
                    Else
                        getBuildPath = "1in"
                    End If
                End If
            End If
    End Select
    
    If rstStud.State > 0 Then
        rstStud.Close
    End If
    
End Function

Public Function GetPolyTableNames(ThisUnit As Unit, strCompTable As String, _
                            strMCOVTable As String) As Boolean
    '------------------------------------------------------------------------------------------------
    '   Created:   10/22/03    By:Ray Raf
    '   Purpose:   To determine the proper tables to use for retrieving polymer
    '                    arrester components and MCOV related data.
    '   Inputs:      ThisUnit - a unit variable that is fully populated
    '   Outputs:    strCompTable - the component table name to use
    '                          (Components_1in, Components_12mm_new, Components_12mm_old)
    '                   strMCOVTable - the MCOV Table to use (MCOV_1in,
    '                               MCOV_12mm_New, MCOV_12mm_old)
    '   Returns:    False if not a polymer arrester
    '------------------------------------------------------------------------------------------------
    If ThisUnit.Type <> 2 Then
        GetPolyTableNames = False
        Exit Function
    End If
    
    With ThisUnit
        If .Stud = "12mm" Then
            'need to check for new or old process
' DLimburg Change 2009_1 11Feb2009, changed table name to Components_12mm_Rev1 from Components_1in
' as part of the implementation of the AT52G
            strCompTable = "Components_12mm_" & Trim(.Process) & "_Rev1"
            strMCOVTable = "MCOV_12mm_" & Trim(.Process)
            GetPolyTableNames = True
            Exit Function
        Else
            If .Stud = "1in" Then
' DLimburg Change 2009_1 11Feb2009, changed table name to Components_1in_Rev1 from Components_1in
' as part of the implementation of the AT52G
                strCompTable = "Components_1in_Rev1"
                strMCOVTable = "MCOV_1in"
                GetPolyTableNames = True
            Else
                'if its not 12mm and not 1in somethings wrong
                GetPolyTableNames = False
                Exit Function
            End If
        End If
    End With
    
End Function

Public Function CrossRefDiskType(strListedDisk As String, _
                                    strActualDisk As String, iType As Integer) As Boolean
    '------------------------------------------------------------------------------------------------
    '   Created:   10/22/03    By:Ray Raf
    '   Purpose:   To determine if the actual disk type is compatable with the polymer arrester
    '                    even though it is not an exact match to the table
    '   Inputs:      strListedDisk - the disk listed as the proper one  in the database
    '                   strActualDisk - the disk type for the batch that the disk is from
    '   Returns:    true if an allowable disk, false otherwise
    '   Rev'd:      12/19/03    By: Rray RAf
    '   Note        added check of arrester type (porc or Poly) for some xrefs
    '   Rev'd:      06/22/06    By: Ray Raf
    '   Note        added check for glass./epoxy substitution flag and checks for G disk types
    '------------------------------------------------------------------------------------------------
    
    'default to disk type NOT being able to be used
    CrossRefDiskType = False
    
    Select Case strListedDisk
        Case "AT32E"
            If InStr(1, strActualDisk, "AT118G") Then
                'old style disks are allowed in old process
                CrossRefDiskType = True
                Exit Function
            End If
            If iType = 1 Then       'only check for porc arresters
                If InStr(1, strActualDisk, "AT118X") Then
                    'old style disks are allowed in old process
                    CrossRefDiskType = True
                    Exit Function
                End If
            End If
            
            'no other disk types match
            CrossRefDiskType = False
            Exit Function
        Case "AT33E"
            If InStr(1, strActualDisk, "AT112G") Then
                'old style disks are allowed in old process
                CrossRefDiskType = True
                Exit Function
            End If
            If iType = 1 Then       'only check for porc arresters
                If InStr(1, strActualDisk, "AT112X") Then
                    'old style disks are allowed in old process
                    CrossRefDiskType = True
                    Exit Function
                End If
            End If
            'no other disk types match
            CrossRefDiskType = False
            Exit Function
        Case "AT34E"
            If InStr(1, strActualDisk, "AT108G") Then
                'old style disks are allowed in old process
                CrossRefDiskType = True
                Exit Function
            End If
            If iType = 1 Then       'only check for porc arresters
                If InStr(1, strActualDisk, "AT108X") Then
                    'old style disks are allowed in old process
                    CrossRefDiskType = True
                    Exit Function
                End If
            End If
            'no other disk types match
            CrossRefDiskType = False
            Exit Function
        Case "AT32G"
            If InStr(1, strActualDisk, "AT32E") And globGandESubs = 1 Then
                'when flag is set, Epoxy can be substituted for glass collars
                CrossRefDiskType = True
                Exit Function
            End If
            
            'no other disk types match
            CrossRefDiskType = False
            Exit Function
        Case "AT33G"
            If InStr(1, strActualDisk, "AT33E") And globGandESubs = 1 Then
                'when flag is set, Epoxy can be substituted for glass collars
                CrossRefDiskType = True
                Exit Function
            End If
            
            'no other disk types match
            CrossRefDiskType = False
            Exit Function
        Case "AT34G"
            If InStr(1, strActualDisk, "AT34E") And globGandESubs = 1 Then
                'when flag is set, Epoxy can be substituted for glass collars
                CrossRefDiskType = True
                Exit Function
            End If
            
            'no other disk types match
            CrossRefDiskType = False
            Exit Function
    End Select          'for strListedDisk
    
End Function

