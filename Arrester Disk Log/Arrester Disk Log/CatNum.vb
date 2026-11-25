Option Strict On
Imports System.Data.SqlClient
Public Class CatNum

#Region "Internal Data Structures"
    Enum Types
        PORC
        POLY
    End Enum

	Enum BuildPaths
        StdAZG = 0
        ' 2019-01-22 RVA - following line was commented out giving many
        ' compile errors
        DualRatedAZG = 1
		TriRatedAZG = 2
		StdAZG5r1 = 3
		StdAZG5r0 = 4
		StdPolymer = 5
		DualRatedPolymer = 6
		UltraSIL = 7
		DualRatedPolymerNew = 8
	End Enum



	Private PRIMARY_MCOV_STARTINDEX() As Integer = {8, 5, 5, 4, 7, 5, 3, 7, 7}

	Private RATING_STARTINDEX() As Integer = {11, 8, 8, 7, 4, 2, 6, 4, 4}

	Private PORCELAIN_COLOR_INDEX() As Integer = {7, 11, 13, 12, 10, 0, 0, 0, 0}

    Private POLY_CUBICLE_MOUNT_CODES() As String = {"B", "F", "G", "J", "K", "L"}


    'Fields
    '-3 = NA, -2 = Not Parsed. -1 = Error
    Private m_type As Integer = -2
    Private m_catNumStr As String
    Private m_givencatNumStr As String
    Private m_buildPath As Integer = -2
    Private m_mcov As Double = -2
    Private m_family As String = Nothing
    Private m_rating As String = Nothing
    Private m_classification As String = Nothing
    Private m_RingShedCode As String = Nothing
    Private m_porcColorCode As String = Nothing
    '-2: Not Parsed, -1 = Error, 0 - False, 1 = True
    Private m_isCubicleMount As Integer = -2
    Private m_stud As String = Nothing
    Private m_diskType As String = Nothing
    Private m_partnerDiskType As String = Nothing
    Private m_process As String = Nothing
    Private m_units As New List(Of ArresterUnit)
    Private modules_set As Boolean = False
    Private m_componentTable As String = Nothing
    Private m_mcovTable As String = Nothing
    Private m_buildTable As String = Nothing
    Private m_mounting As String = Nothing
    Private m_namePlate As Nameplate = Nothing
    Private m_mcovStr As String = Nothing
    Private m_serNum As Long = 0
    Private m_workOrder As WorkOrder
    Private m_crossedFamily As String = Nothing
    Private m_diskDia As String = Nothing
    Private m_IRCode As String = Nothing
    Private m_cantiStrength As String = Nothing
    Private m_HsgGrdRing As String = Nothing    ' DLimburg added this veriable - 25March2014

#End Region

#Region "Properties"
	ReadOnly Property Units As List(Of ArresterUnit)
		Get
			If (Not modules_set) Then
				fetchModules()
			End If
			Return m_units
		End Get
	End Property

	Property SerialNumber() As Long
		Get
			Return m_serNum
		End Get
		Set(ByVal value As Long)
			m_serNum = value
		End Set
	End Property

	Property CatNumString() As String
		Get
			If (m_catNumStr.Length = 6) Or (m_catNumStr.Length = 8) Then	 '   20Apr2012   Added test for new polymer shortcuts
				LookupNickname()
			End If
			Return m_catNumStr
		End Get
		Set(ByVal value As String)
			m_catNumStr = value.ToUpper
			ParseCatNumBasic()
		End Set
	End Property

	Property GivenCatNum() As String
		Get
			Return m_givencatNumStr
		End Get
		Set(ByVal value As String)
			m_givencatNumStr = value.ToUpper
		End Set
	End Property

	ReadOnly Property Type() As Integer
		Get
			If m_type = -2 Then
				ParseType()
			End If
			Return m_type
		End Get
	End Property

	ReadOnly Property Rating() As String
		Get
			If (m_rating = Nothing) Then
				ParseRating()
			End If
			Return m_rating
		End Get
	End Property

	ReadOnly Property BuildPath() As Integer
		Get
			If (m_buildPath = -2) Then
				ParseBuildPath()
			End If
			Return m_buildPath
		End Get
	End Property

	ReadOnly Property Classification() As String
		Get
			If (m_classification = Nothing) Then
				ParseClassification()
			End If
			Return m_classification
		End Get
	End Property

	ReadOnly Property RingShedCode() As String
		Get
			If (m_RingShedCode = Nothing) Then
				ParseRingShedCode()
			End If
			Return m_RingShedCode
		End Get
	End Property

	ReadOnly Property IsCubicleMount() As Integer
		Get
			If (m_isCubicleMount = -2) Then
				ParseCubicleMount()
			End If
			Return m_isCubicleMount
		End Get
	End Property

	ReadOnly Property Family() As String
		Get
			If (m_family = Nothing) Then
				ParseFamily()
			End If
			Return m_family
		End Get
	End Property

	ReadOnly Property CrossedFamily As String
		Get
			If (m_crossedFamily Is Nothing) Then
				lookupPorcelainCrossRef()
			End If
			Return m_crossedFamily
		End Get
	End Property

	ReadOnly Property MCOV() As Double
		Get
			If (m_mcov = -2) Then
				ParseMCOV()
			End If
			Return m_mcov
		End Get
	End Property

	ReadOnly Property MCOVStr() As String
		Get
			If (m_mcovStr = Nothing) Then
				ParseMCOV()
			End If
			Return m_mcovStr
		End Get
	End Property

	ReadOnly Property TypeString() As String
		Get
			If (Type = Types.PORC) Then
				Return "Porcelain"
			ElseIf (Type = Types.POLY) Then
				Return "Polymer"
			Else
				Return "Unknown"
			End If
		End Get
	End Property

	ReadOnly Property Stud() As String
		Get
			If (m_stud = Nothing) Then
				LookupPolyStudDisk()
			End If
			Return m_stud
		End Get
	End Property

	ReadOnly Property DiskType() As String
		Get
			If (m_diskType = Nothing) Then
				LookupPolyStudDisk()
			End If
			Return m_diskType
		End Get
	End Property

	ReadOnly Property PartnerDiskType As String
		Get
			If (m_partnerDiskType = Nothing) Then
				LookupPartnerDisk()
			End If
			Return m_partnerDiskType
		End Get
	End Property

	ReadOnly Property Process() As String
		Get
			If (m_process = Nothing) Then
				LookupPolyStudDisk()
			End If
			Return m_process
		End Get
	End Property

	ReadOnly Property IRCode() As String
		Get
			If (m_IRCode Is Nothing) Then
				ParseIRCode()
			End If
			Return m_IRCode
		End Get
    End Property

    ReadOnly Property HsgGrdRingCode() As String 'add this routine DLimburg 25March2014
        Get
            If (m_HsgGrdRing Is Nothing) Then
                ParseHsgGradingRingCode()
            End If
            Return m_HsgGrdRing
        End Get
    End Property

	ReadOnly Property CantileverStrength As String
		Get
			If (m_cantiStrength Is Nothing) Then
				ParseCantiStrength()
			End If
			Return m_cantiStrength
		End Get
	End Property

	Property WorkOrder As WorkOrder
		Get
			Return m_workOrder
		End Get
		Set(ByVal value As WorkOrder)
			m_workOrder = value
		End Set
	End Property

	ReadOnly Property DiskDia As String
		Get
			If (m_diskDia = Nothing) Then
				LookupPolyStudDisk()
			End If
			Return m_diskDia
		End Get
	End Property

	ReadOnly Property PorcColorCode() As String
		Get
			If m_porcColorCode = Nothing Then
				ParsePorcColorCode()
			End If
			Return m_porcColorCode
		End Get
	End Property

	ReadOnly Property ComponentTable() As String
		Get
			If (m_componentTable = Nothing) Then
				GetTableNames()
			End If
			Return m_componentTable
		End Get
	End Property

	ReadOnly Property MCOVTable() As String
		Get
			If (m_mcovTable = Nothing) Then
				GetTableNames()
			End If
			Return m_mcovTable
		End Get
	End Property

	ReadOnly Property BuildTable() As String
		Get
			If (m_buildTable = Nothing) Then
				GetTableNames()
			End If
			Return m_buildTable
		End Get
	End Property

	ReadOnly Property Mounting() As String
		Get
			If (m_mounting Is Nothing) Then
				ParseMounting()
			End If
			Return m_mounting
		End Get
	End Property

	ReadOnly Property Nameplate() As Nameplate
		Get
			If (m_namePlate Is Nothing) Then
				m_namePlate = New Nameplate(Me)
			End If
			Return m_namePlate
		End Get
	End Property

	ReadOnly Property SqlFamily As String
		Get
			If (Type = Types.PORC) Then
				Return " ( Family = '" & Family & "' Or Family = '" & CrossedFamily & "' ) "
			Else
				Return " Family = '" & Family & "'"
			End If

		End Get
	End Property


	ReadOnly Property NumUnits() As Integer
		Get
			Return m_units.Count
		End Get
	End Property

	ReadOnly Property NumModules() As Integer
		Get
			Dim size As Integer = 0
			Dim unit As ArresterUnit
			For Each unit In Units
				size += unit.Modules.Count
			Next unit
			Return size
		End Get
	End Property


#End Region

#Region "Parse Methods"
    Private Sub ParseCatNumBasic()
        If (Not (m_catNumStr.Length = 6 Or m_catNumStr.Length = 8 Or m_catNumStr.Length = 14 Or m_catNumStr.Length = 15 Or m_catNumStr.Length = 18)) Then   ' 20Apr2012 DLimburg - Added valid length for polymer shortcuts.
            MsgBox("Invalid Part Number: Insufficient number of characters in given catalog number.", MsgBoxStyle.Critical)  ' 20Apr2012 DLimburg modified error message 
            m_catNumStr = "ERR"
        End If
    End Sub

    Private Sub ParseType()
        If m_catNumStr.Chars(0) = "U" Then
            m_type = Types.POLY
        ElseIf m_catNumStr.Chars(0) = "A" Then
            m_type = Types.PORC
        Else
            MsgBox("Invalid Catalog Number: Must start with A or U. Could not parse type.", MsgBoxStyle.Critical)
            m_type = -1
        End If
    End Sub

    Private Sub ParseBuildPath()
        
		If (m_catNumStr.Chars(0) = "U"c AndAlso m_catNumStr.Chars(2) = "D"c AndAlso m_catNumStr.Length = 15) Then
			m_buildPath = BuildPaths.DualRatedPolymer
		ElseIf (m_catNumStr.Chars(0) = "U"c AndAlso m_catNumStr.Chars(2) = "D"c AndAlso m_catNumStr.Length = 18) Then
			m_buildPath = BuildPaths.DualRatedPolymerNew
		ElseIf (m_catNumStr.Length = 18 AndAlso m_catNumStr.Chars(0) = "U"c) Then		' Changed .Substring test from "US" DLimburg 26March2012
			m_buildPath = BuildPaths.UltraSIL
		ElseIf (m_catNumStr.Chars(0) = "U"c) Then
			m_buildPath = BuildPaths.StdPolymer
		ElseIf (m_catNumStr.Substring(0, 5) = "AZEST") Then
			m_buildPath = BuildPaths.TriRatedAZG
        ElseIf (m_catNumStr.Substring(0, 3) = "AZG" And m_catNumStr.Substring(4, 1) = "D") Or (m_catNumStr.Substring(0, 3) = "AZE" And m_catNumStr.Substring(4, 1) = "D") Then ' Need to add test for dual rated AZE, ' DLimburg 11Dec2013  added test for AZE, not sure why it was not added at time it was identified by previous comment
            m_buildPath = BuildPaths.DualRatedAZG
		ElseIf (m_catNumStr.Substring(0, 3) = "AZG" Or m_catNumStr.Substring(0, 3) = "AZE") Then
			m_buildPath = BuildPaths.StdAZG
		ElseIf (m_catNumStr.Substring(0, 4) = "AZG5" And m_catNumStr.Length = 15) Then
			m_buildPath = BuildPaths.StdAZG5r0
		ElseIf (m_catNumStr.Substring(0, 4) = "AZG5" And m_catNumStr.Length = 16) Then
			m_buildPath = BuildPaths.StdAZG5r1
		Else
			m_buildPath = -1
			MsgBox("Invalid Catalog Number: Couldn't identify the build path.")
		End If
	End Sub

	Private Sub ParseMCOV()
		Dim tmpMCOV As Integer
		Dim success As Boolean = Integer.TryParse(m_catNumStr.Substring(PRIMARY_MCOV_STARTINDEX(BuildPath), 3), tmpMCOV)
		If (success = False) Then
			MsgBox("Part Number Not Valid: Could not parse the MCOV Value.", MsgBoxStyle.Critical)
			m_mcov = -1
		Else
			Dim actMCOVQuery As String = "SELECT * FROM " & ACTUAL_MCOV & " WHERE String_MCOV = " & tmpMCOV
			Dim actMCOVCommand As New SqlCommand(actMCOVQuery, GRoomConnection)

			Dim actMCOVReader As SqlDataReader = actMCOVCommand.ExecuteReader
			Try
				If (actMCOVReader.Read()) Then
					Double.TryParse(CStr(actMCOVReader!Actual_MCOV), m_mcov)
				Else
					m_mcov = tmpMCOV
				End If
			Finally
				actMCOVReader.Close()
			End Try
			m_mcovStr = m_catNumStr.Substring(PRIMARY_MCOV_STARTINDEX(BuildPath), 3)
		End If

	End Sub

    Private Sub ParseFamily()
        If (Type = Types.PORC) Then
            m_family = Mid(m_catNumStr, 1, 4)
        ElseIf (Type = Types.POLY) Then
            m_family = Mid(m_catNumStr, 1, 2)
        End If

    End Sub

    Private Sub ParseRating()
        m_rating = m_catNumStr.Substring(RATING_STARTINDEX(BuildPath), 3)
    End Sub

    Private Sub ParseClassification()
        If (Type = Types.POLY) Then
            m_classification = m_catNumStr.Substring(1, 1)
        Else
            m_classification = "NA"
        End If
    End Sub

    Private Sub ParseRingShedCode()
        If (BuildPath = BuildPaths.StdAZG) Then
            m_RingShedCode = m_catNumStr.Substring(5, 2)
        ElseIf (BuildPath = BuildPaths.StdPolymer) Then
            m_RingShedCode = m_catNumStr.Substring(8, 2)
        ElseIf (m_buildPath = BuildPaths.UltraSIL) Then
            m_RingShedCode = m_catNumStr.Substring(11, 2)
        Else
            m_RingShedCode = "NA"
        End If

    End Sub

    Private Sub ParseCubicleMount()
        If (Type = Types.PORC) Then
            Dim ringCode As Integer = 0
            Integer.TryParse(RingShedCode, ringCode) 'To prevent RingShedCode from being parsed twice on an error
            If (ringCode >= 91 And ringCode <= 95) Then
                m_isCubicleMount = CInt(True)
            Else
                m_isCubicleMount = CInt(False)
            End If
        ElseIf (Type = Types.POLY) Then
            If (BuildPath = BuildPaths.StdPolymer And Array.IndexOf(POLY_CUBICLE_MOUNT_CODES, Mounting) <> -1) Then
                m_isCubicleMount = CInt(True)
            Else
                m_isCubicleMount = CInt(False)
            End If
        End If

    End Sub

    Private Sub ParsePorcColorCode()
        If (Type = Types.PORC) Then
            m_porcColorCode = m_catNumStr.Substring(PORCELAIN_COLOR_INDEX(BuildPath), 1)
            If (m_porcColorCode <> "B" And m_porcColorCode <> "G") Then
                MsgBox("Invalid color code: Must be a B or a G.")
                m_porcColorCode = "ERR"
            End If

        Else
            m_porcColorCode = "NA"
        End If

    End Sub

    Private Sub ParseMounting()
        If (Type = Types.POLY And m_catNumStr.Length >= 13) Then
            m_mounting = Mid(m_catNumStr, 13, 1)
        Else
            m_mounting = "NA"
        End If
    End Sub

    Private Sub ParseIRCode()
        If (BuildPath = BuildPaths.UltraSIL) Then
            m_IRCode = CatNumString.Substring(2, 1)  'DLimburg changed string position to 2 from 10  - 25March2014
		ElseIf BuildPath = BuildPaths.DualRatedPolymerNew Then
			m_IRCode = CatNumString.Substring(3, 1)
		Else
			m_IRCode = "NA"
        End If
    End Sub

    Private Sub ParseHsgGradingRingCode()
        'DLimburg added this routine  - 25March2014
        If (BuildPath = BuildPaths.UltraSIL) Or (BuildPath = BuildPaths.DualRatedPolymerNew) Then
            m_HsgGrdRing = CatNumString.Substring(10, 1)
        Else
            m_HsgGrdRing = "NA"
        End If
    End Sub

    Private Sub ParseCantiStrength()
		If (BuildPath = BuildPaths.UltraSIL) Then
            m_cantiStrength = CatNumString.Substring(3, 1)
		ElseIf BuildPath = BuildPaths.DualRatedPolymerNew Then
			m_cantiStrength = "A"
		Else
			m_cantiStrength = "NA"
		End If
    End Sub

#End Region

#Region "DB Interaction Methods"

    Public Function Validate() As Boolean
        Dim validated As Boolean = True

        If (m_catNumStr = "ERR") Then
            validated = False
        End If

        If (validated = True And Type = -1) Then
            validated = False
        End If

        If (validated = True And Family = "ERR") Then
            validated = False
        End If

        If (validated = True And Rating = "ERR") Then
            validated = False
        End If

        If (validated = True And Classification = "ERR") Then
            validated = False
        End If

        If (validated = True And RingShedCode = "ERR") Then
            validated = False
        End If

        '        If (validated = True And IsCubicleMount = -1) Then ' 23Apr2012 DLimburg removed this test as it prevents a cubicle mount from being built
        '       validated = False
        '       End If

        If (validated = True And Stud = "ERR") Then
            validated = False
        End If

        If (validated = True And DiskType = "ERR") Then
            validated = False
        End If
        If (validated = True And Process = "ERR") Then
            validated = False
        End If

        If (validated = True And BuildPath = -1) Then
            validated = False
        End If

        If (validated = True And MCOV = -2.0) Then
            validated = False
        End If



        Return validated

    End Function
    Private Sub GetTableNames()
        If (Stud = "12mm" And Process = "new") Then
            m_mcovTable = MCOV_12MM_NEW
            m_componentTable = COMPONENTS_12MM_NEW
        ElseIf (Stud = "12mm" And Process = "old") Then
            m_mcovTable = MCOV_12MM
            m_componentTable = COMPONENTS_12MM_OLD
        ElseIf (Stud = "1in") Then
            m_mcovTable = MCOV_1IN
            m_componentTable = COMPONENTS_1IN
        End If
    End Sub
    Private Function CheckAgainstMCOVTables() As Boolean
        Dim success As Boolean
        Dim query As String = "Select 0"
        If Type = Types.PORC Then
            If (BuildPath = BuildPaths.DualRatedAZG) Then
                query = "SELECT * FROM " & MCOV_DR_PORC & " WHERE CatalogNumber='" & m_catNumStr & "' "


            ElseIf (CBool(IsCubicleMount)) Then
                query = "SELECT * FROM " & MCOV_CUBE & " " & _
                  "WHERE ArresterRating='" & Rating & "' " & _
                  "AND String_MCOV ='" & MCOV & "' " & _
                  "AND PorcRingNumber='" & RingShedCode & "' " & _
                  "AND Family='" & Family & "' ;"
            Else
                query = "SELECT * FROM " & MCOV_PORC_CALCINED & " " & _
                  "WHERE ArresterRating='" & Rating & "' " & _
                  "AND String_MCOV ='" & MCOV & "' " & _
                  "AND PorcRingNumber='" & RingShedCode & "' " & _
                  "AND Family='" & Family & "' ;"
            End If
        ElseIf Type = Types.POLY Then
            If (BuildPath = BuildPaths.DualRatedPolymer) Then
                query = "SELECT * FROM " & MCOV_DR_POLY & _
                "WHERE CatalogNumber='" & m_catNumStr & "' "
            ElseIf (BuildPath = BuildPaths.StdPolymer) Then
                Dim table As String
                table = "MCOV_" & Stud & Process
                query = "SELECT * FROM " & table & " " & _
                "WHERE ArresterRating='" & Rating & "' " & _
                "AND String_MCOV ='" & MCOV & "' " & _
                "AND ShedCode = '" & RingShedCode & "' " & _
                "AND Family ='" & Family & "' "
            End If
        End If

        Dim MCOVLookupCommand As New SqlCommand(query, GRoomConnection)
        Dim MCOVLookupReader As SqlDataReader = MCOVLookupCommand.ExecuteReader()
        Try
            If (Not MCOVLookupReader.HasRows()) Then
                MsgBox("Unable to find a matching entry for MCOV Lookup. Therefore the catalog " & _
                "number entered is invalid.", MsgBoxStyle.Critical)
                success = False
            Else
                success = True
            End If
        Finally
            MCOVLookupReader.Close()
        End Try
        Return success
    End Function
    Private Function LookupPolyStudDisk() As Boolean
        Dim NameplateTable As String
        If (Type = Types.POLY) Then
            If (BuildPath = BuildPaths.UltraSIL) Then
				Dim query As String = "SELECT * FROM " & MCOV_HOUSING_ULTRASIL & " " & _
				"WHERE Family='" & Family & "' " & _
				" AND Housing_Code = '" & RingShedCode & "' " & _
				" AND Rating_Arrester = " & Rating & _
				" AND MCOV_Arrester = " & MCOV & _
				" AND Cantilever_Strength_Option='" & CantileverStrength & "'"

                Dim ultraSILCommand As New SqlCommand(query, PolymerStationConnection)
                Dim ultraSILReader As SqlDataReader = ultraSILCommand.ExecuteReader()
                Try
                    If (ultraSILReader.HasRows()) Then
                        ultraSILReader.Read()
                        m_stud = CStr(ultraSILReader!Stud)
                        m_diskDia = CStr(ultraSILReader!Disk_Type)
                        m_process = "NA"
                        m_diskType = "NA"

                    Else
                        MsgBox("The disk type and stud could not be retrieved for this arrester.", MsgBoxStyle.Critical)
                        m_stud = "ERR"
                        m_diskType = "ERR"
                        m_process = "ERR"
                    End If
                Finally
                    ultraSILReader.Close()
                End Try
                LookupPartnerDisk()

            Else
                If (BuildPath = BuildPaths.DualRatedPolymer) Then
                    NameplateTable = NAMEPLATE_DR_POLY
                    '12June2012 RPost Added following code, there was nothing for this condition
                    Dim sql As String = "SELECT * FROM " & NameplateTable & " WHERE CatalogNumber ='" & CatNumString & "'"
                    Dim myCommand As SqlCommand = New SqlCommand(sql, GRoomConnection)
                    Dim myReader As SqlDataReader = myCommand.ExecuteReader()
                    Try
                        If myReader.Read() Then
                            m_stud = CStr(myReader!stud)
                            m_diskType = CStr(myReader!DiskType)
                            m_process = CStr(myReader!Process)
                            m_diskDia = "NA"
                        Else
                            MsgBox("The disk type and stud could not be retrieved for this arrester.", MsgBoxStyle.Critical)
                            m_stud = "ERR"
                            m_diskType = "ERR"
                            m_process = "ERR"
                        End If
                    Finally
                        myReader.Close()
					End Try
				ElseIf (BuildPath = BuildPaths.DualRatedPolymerNew) Then
					ProcessNewDualRated()
				Else
					If (CBool(IsCubicleMount)) Then
						NameplateTable = NAMEPLATE_CUBE_POLY
					Else
						NameplateTable = NAMEPLATE_POLY
					End If
					Dim query As String = "SELECT * FROM " & NameplateTable & " " & _
					"WHERE Family='" & Family & "' " & _
					"AND MinRating <= '" & Rating & "' " & _
					"AND MaxRating >= '" & Rating & "' "
					Dim cmd As SqlCommand = New SqlCommand(query, GRoomConnection)
					Dim rdr As SqlDataReader = cmd.ExecuteReader()
					Try
						If (rdr.Read()) Then

							m_stud = CStr(rdr!stud)
							m_diskType = CStr(rdr!DiskType)
							m_process = CStr(rdr!Process)
							m_diskDia = "NA"
						Else
							MsgBox("The disk type and stud could not be retrieved for this arrester.", MsgBoxStyle.Critical)
							m_stud = "ERR"
							m_diskType = "ERR"
							m_process = "ERR"
						End If
					Finally
						rdr.Close()
					End Try

                End If
            End If
        ElseIf (Type = Types.PORC) Then
            Dim query As String
            If (BuildPath = BuildPaths.DualRatedAZG) Then
                query = "SELECT DiskType FROM " & MCOV_DR_PORC & " WHERE CatalogNumber='" & CatNumString & "' ;"
            ElseIf (CBool(IsCubicleMount)) Then
                query = "Select DiskType FROM " & MCOV_CUBE & " WHERE (Family='" & Family & "' " & _
                 "Or Family = '" & CrossedFamily & "') AND String_MCOV='" & MCOVStr & "' AND ArresterRating=" & Rating & " " & _
                "AND PorcRingNumber = " & RingShedCode & " ;"
            Else
                query = "Select DiskType FROM " & MCOV_PORC_CALCINED & " WHERE (Family='" & Family & "' Or Family = '" & CrossedFamily & "') " & _
                "AND String_MCOV='" & MCOVStr & "' AND ArresterRating=" & Rating & " " & _
                "AND PorcRingNumber = '" & RingShedCode & "' ;"
            End If
            Dim porcDiskCommand As SqlCommand = New SqlCommand(query, GRoomConnection)
            Dim porcDiskReader As SqlDataReader = porcDiskCommand.ExecuteReader()
            Try
                If (porcDiskReader.HasRows()) Then
                    porcDiskReader.Read()
                    m_diskType = CStr(porcDiskReader!DiskType)
                Else
                    MsgBox("Could not look up porcelain disk type.", MsgBoxStyle.Critical)
                End If
            Finally
                porcDiskReader.Close()
            End Try

        End If

    End Function

	Private Sub ProcessNewDualRated()
		'2/19/12 RPost Added this new DualRated 18digit
		Dim sql As String = "SELECT * FROM " & "Poly_Dual_Rated_Arresters_Rev0" & " WHERE CatalogNumber ='" & CatNumString & "'"
		Dim myCommand As SqlCommand = New SqlCommand(sql, PolymerStationConnection)
		Dim rdr As SqlDataReader = myCommand.ExecuteReader()
		Try
			If rdr.Read() Then

				m_stud = CStr(rdr!stud)
				m_diskDia = CStr(rdr!Disk_Type)	'Ya, who the hell does that?
				m_mcov = CDbl(rdr!MCov_Arrester)
				m_rating = CStr(rdr!Rating_arrester)
				m_givencatNumStr = CStr(rdr!CatalogNumber)
				'Adding unit 1
				m_units.Add(New ArresterUnit(1))
				m_units(0).AddModule(1, Me, CDbl(rdr!MCOV_Bottom_Unit_Mod1))
				m_units(0).m_modules(0).ShedSet(CInt(rdr!HsgPairCnt_Bottom_Mod1))
				'Add Mod2 to Unit 1 if needed
				If Not IsDBNull(rdr!MCOV_Bottom_Unit_Mod2) AndAlso CDbl(rdr!MCOV_Bottom_Unit_Mod2) > 0 Then
					m_units(0).AddModule(2, Me, CDbl(rdr!MCOV_Bottom_Unit_Mod2))
					m_units(0).m_modules(1).ShedSet(CInt(rdr!HsgPairCnt_Bottom_Mod2))
				End If
				'Adding unit 2  DR always have 2
				m_units.Add(New ArresterUnit(2))
				m_units(1).AddModule(1, Me, CDbl(rdr!MCOV_Top_Unit_Mod1))
				m_units(1).m_modules(0).ShedSet(CInt(rdr!HsgPairCnt_Top_Mod1))
				'Add Mod2 if U2 needs it
				If Not IsDBNull(rdr!MCOV_Top_Unit_Mod2) AndAlso CDbl(rdr!MCOV_Top_Unit_Mod2) > 0 Then
					m_units(1).AddModule(2, Me, CDbl(rdr!MCOV_Top_Unit_Mod2))
					m_units(1).m_modules(1).ShedSet(CInt(rdr!HsgPairCnt_Top_Mod2))

				End If

				'Rating_Top_Unit,
				'MCOV_Top_Unit,
				'MCOV_Top_Unit_Mod1,
				'Name_Top_Unit_Mod1,
				'MCOV_Top_Unit_Mod2,
				'Name_Top_Unit_Mod2,

				m_process = "new"

				rdr.Close()
				myCommand.CommandText = "SELECT Part_Number,Part_Description FROM Poly_ModuleParts_Rev0 WHERE Disk_Type ='" & m_diskDia & "' AND Family = '" & m_family & "' AND Part_Description IN ('3kV Disk','5kV Disk')"
				rdr = myCommand.ExecuteReader

				While rdr.Read
					Select Case rdr.GetString(1)
						Case "3kV Disk"
							m_diskType = rdr.GetString(0)
						Case Else
							'5kv
							m_partnerDiskType = rdr.GetString(0)
					End Select
				End While
			Else
				MsgBox("This dual rated arrester has not been loaded.", MsgBoxStyle.Critical)
				m_stud = "ERR"
				m_diskType = "ERR"
				m_process = "ERR"
			End If
		Finally
			rdr.Close()
		End Try
	End Sub

	Private Sub LookupPartnerDisk()

		If (BuildPath = BuildPaths.UltraSIL Or BuildPath = BuildPaths.DualRatedPolymerNew) Then
			Dim sql As String = "SELECT * FROM DiskInformation_Rev0 WHERE Disk_Dia = '" & DiskDia & "'"
			Dim cmd As New SqlCommand(sql, PolymerStationConnection)
			Dim rdr As SqlDataReader = cmd.ExecuteReader()
			Try
				If (rdr.Read()) Then

					m_diskType = CStr(rdr!DiskPartNumber_3kV)
					m_partnerDiskType = CStr(rdr!DiskPartNumber_5kV)
				Else
					MsgBox("Unable to get the disk partner for disk: " & DiskType)
					m_partnerDiskType = "ERR"
				End If
			Finally
				rdr.Close()
			End Try


		Else
			Dim sql As String = "SELECT * FROM " & DISK_PARTNERS & " " & _
			 " WHERE DiskPartNo_3kV = '" & DiskType & "'"


			Dim cmd As New SqlCommand(sql, GRoomConnection)
			Dim rdr As SqlDataReader = cmd.ExecuteReader
			Try
				If (rdr.Read()) Then
					m_partnerDiskType = CStr(rdr!DiskPartNo_5kV)
				Else
					MsgBox("Unable to get the disk partner for disk: " & DiskType)
					m_partnerDiskType = "ERR"
				End If
			Finally
				rdr.Close()
			End Try
		End If


	End Sub

	Private Function LookupNickname() As Boolean
		Dim retVal As Boolean = True
		If (m_catNumStr.Length = 6) Then			 ' old style polymer cat number, goto GRoom for information
			Dim lookupQuery As String = "SELECT * FROM ShortCutListing WHERE ShortCut = '" & m_catNumStr & "' "
			Dim UltraSILCommand As New SqlCommand(lookupQuery, GRoomConnection)
			Dim myReader As SqlDataReader = UltraSILCommand.ExecuteReader()
			'    Dim myReader As SqlDataReader = New SqlCommand(lookupQuery).ExecuteReader()        ' 20Apr2012 DLimburg - Removed and replace with connection strings above
			Try
				If (myReader.HasRows) Then
					myReader.Read()
					m_catNumStr = CStr(myReader!CatNo)
				Else
					MsgBox("Unable to find a catalog number for this catalog number shortcut.", MsgBoxStyle.Critical)
					retVal = False
				End If
			Finally
				myReader.Close()
			End Try
		Else	  ' 20Apr2012 DLimburg, added else portion to handle new polymer, only get to LookupNickname if dealing with a string 6 or 8 characters long.
			' new style polymer cat number, goto PolymerStation for information
			Dim lookupQuery As String = "SELECT * FROM ShortCutListing_Rev0 WHERE ShortCut = '" & m_catNumStr & "' "
			Dim UltraSILCommand As New SqlCommand(lookupQuery, PolymerStationConnection)
			Dim myReader As SqlDataReader = UltraSILCommand.ExecuteReader()
			Try
				If (myReader.HasRows) Then
					myReader.Read()
					m_catNumStr = CStr(myReader!CatNo)
				Else
					MsgBox("Unable to find a catalog number for this catalog number shortcut.", MsgBoxStyle.Critical)
					retVal = False
				End If
			Finally
				myReader.Close()
			End Try
		End If
		Return retVal
	End Function

	Private Function fetchModules() As Boolean
		modules_set = True
		Dim tmpModules As New List(Of ArresterModule)
		Dim query As String = "Select 0"


		If BuildPath = BuildPaths.UltraSIL Then
			query = "SELECT * FROM " & MCOV_HOUSING_ULTRASIL & " WHERE Family = '" & Family & "' And Housing_Code = '" & RingShedCode & _
			"' And Disk_Type = '" & DiskDia & "' And Stud = '" & Stud & "' And Rating_Arrester = " & Rating & " AND MCOV_Arrester = " & MCOV & " AND Cantilever_Strength_Option='" & CantileverStrength & "'" 'RPost 11-28-12 Added MCOV to Query criteria

			Dim UltraSILCommand As New SqlCommand(query, PolymerStationConnection)
			Dim UltraSILReader As SqlDataReader = UltraSILCommand.ExecuteReader()
			Try
				If (UltraSILReader.HasRows()) Then
					UltraSILReader.Read()
					Dim X As Integer
					For X = 1 To 6	'Up to 6 units total
						Dim tmpUnitMCOV As String = "0"
						Dim tmpModuleMCOV As String = "0"
						Try
							tmpUnitMCOV = CStr(UltraSILReader.Item(UltraSILReader.GetOrdinal("MCOV_Unit" & X)))
						Catch ie As IndexOutOfRangeException 'If we can't find the column, that's okay, the unit doesn't have to exist.
						End Try
						' 'Dim tmpMCOV As String = "Test"
						If tmpUnitMCOV <> "0" And tmpUnitMCOV <> "0.00" Then '10May2012 RPost Added 0.00 test
							m_units.Add(New ArresterUnit(X))
							tmpModuleMCOV = CStr(UltraSILReader.Item(UltraSILReader.GetOrdinal("MCOV_Unit" & X & "_Module" & 1)))
							If (tmpModuleMCOV <> "0" And tmpModuleMCOV <> "0.00") Then
								m_units(X - 1).AddModule(1, Me, CDbl(tmpModuleMCOV))
							End If
							tmpModuleMCOV = CStr(UltraSILReader.Item(UltraSILReader.GetOrdinal("MCOV_Unit" & X & "_Module" & 2)))
							If (tmpModuleMCOV <> "0" And tmpModuleMCOV <> "0.00") Then
								m_units(X - 1).AddModule(2, Me, CDbl(tmpModuleMCOV))
							End If
						End If
					Next X

				Else
					MsgBox("Could not fetch the MCOV/Housing entry for the UltraSIL Part upon Module Lookup.", MsgBoxStyle.Critical)
				End If
			Finally
				UltraSILReader.Close()
			End Try
		ElseIf BuildPath = BuildPaths.DualRatedPolymerNew Then
			ProcessNewDualRated()
		Else

			Dim ratingInt As Integer
			Integer.TryParse(Rating, ratingInt)
			If Type = Types.PORC Then
				If (BuildPath = BuildPaths.DualRatedAZG) Then
					query = "SELECT * FROM " & MCOV_DR_PORC & " WHERE CatalogNumber='" & m_catNumStr & "' "


				ElseIf (CBool(IsCubicleMount)) Then
					query = "SELECT * FROM " & MCOV_CUBE & " " & _
					  "WHERE ArresterRating=" & ratingInt & " " & _
					  "AND String_MCOV ='" & MCOVStr & "' " & _
					  "AND PorcRingNumber='" & RingShedCode & "' " & _
					  "AND (Family='" & Family & "' Or Family = '" & CrossedFamily & "') ;"
				Else
					query = "SELECT * FROM " & MCOV_PORC_CALCINED & " " & _
					  "WHERE ArresterRating=" & ratingInt & " " & _
					  "AND String_MCOV ='" & MCOVStr & "' " & _
					  "AND PorcRingNumber='" & RingShedCode & "' " & _
					  "AND (Family='" & Family & "' Or Family = '" & CrossedFamily & "') ;"
				End If
			ElseIf Type = Types.POLY Then
				If (BuildPath = BuildPaths.DualRatedPolymer) Then
					query = "SELECT * FROM " & MCOV_DR_POLY & _
					" WHERE CatalogNumber='" & m_catNumStr & "' "
				ElseIf (BuildPath = BuildPaths.StdPolymer) Then
					Dim table As String = ""
					If (Stud = "12mm") Then
						table = "MCOV_" & Stud & "_" & Process
					ElseIf (Stud = "1in") Then
						table = MCOV_1IN
					End If

					query = "SELECT * FROM " & table & " " & _
					"WHERE ArresterRating=" & ratingInt & " " & _
					"AND String_MCOV = '" & MCOVStr & "' " & _
					"AND ShedCode = '" & RingShedCode & "' " & _
					"AND Family ='" & Family & "' "
				End If
			End If
			Dim ModuleLookupCommand As New SqlCommand(query, GRoomConnection)
			Dim ModuleLookupReader As SqlDataReader = ModuleLookupCommand.ExecuteReader()
			Try
				If (ModuleLookupReader.HasRows()) Then
					ModuleLookupReader.Read()
					If (BuildPath = BuildPaths.DualRatedAZG Or BuildPath = BuildPaths.DualRatedPolymer) Then
						m_units.Add(New ArresterUnit(1))
						m_units.Add(New ArresterUnit(2))
						m_units(0).AddModule(1, Me, CDbl(ModuleLookupReader!Mcov_Unit_1))
						m_units(1).AddModule(1, Me, CDbl(ModuleLookupReader!Mcov_Unit_2_Lower))
					ElseIf (Family = "U4") Then ' For U4's, modules/units are equivalent
						Dim X As Integer
						For X = 1 To 6	'Up to 6 units total
							Dim tmpMCOV As String = "0"
							Try
								tmpMCOV = CStr(ModuleLookupReader.Item(ModuleLookupReader.GetOrdinal("MCOV_Unit_" & X)))
							Catch ie As IndexOutOfRangeException 'If we can't find the column, that's okay, the unit doesn't have to exist.
							End Try
							If tmpMCOV <> "0" Then
								m_units.Add(New ArresterUnit(X))
								m_units(X - 1).AddModule(1, Me, CDbl(tmpMCOV))
							End If
						Next X
					ElseIf (BuildPath = BuildPaths.StdAZG Or BuildPath = BuildPaths.StdAZG5r0 Or BuildPath = BuildPaths.StdAZG5r1) Then
						Dim X As Integer
						For X = 1 To 6	'Up to 6 units total
							Dim tmpMCOV As String = "0"
							Try
								tmpMCOV = CStr(ModuleLookupReader.Item(ModuleLookupReader.GetOrdinal("MCOV_Unit_" & X)))
							Catch ie As IndexOutOfRangeException 'If we can't find the column, that's okay, the unit doesn't have to exist.
							End Try
							' 'Dim tmpMCOV As String = "Test"
							If tmpMCOV <> "0" Then
								m_units.Add(New ArresterUnit(X))
								tmpModules.Add(New ArresterModule(m_units(X - 1), Me, 1, tmpMCOV))
								m_units(X - 1).AddModule(1, Me, CDbl(tmpMCOV))
							End If
						Next X
					Else
						Dim X As Integer
						For X = 1 To 6	'Up to 6 units total
							Dim tmpMCOV As String = "0"
							Try
								tmpMCOV = CStr(ModuleLookupReader.Item(ModuleLookupReader.GetOrdinal("MCOV_Unit_" & X)))
							Catch ie As IndexOutOfRangeException 'If we can't find the column, that's okay, the unit doesn't have to exist.
							End Try
							' 'Dim tmpMCOV As String = "Test"
							If tmpMCOV <> "0" Then
								tmpModules.Add(New ArresterModule(Nothing, Nothing, 1, tmpMCOV))
								'm_units.Add(New ArresterUnit)
								' m_units(X - 1).AddModule(1, Nothing, tmpMCOV)
							End If
						Next X

						Dim unitCounter As Integer = 1

						If (tmpModules.Count Mod 2 = 1) Then 'Odd number of modules
							m_units.Add(New ArresterUnit(1))
							m_units(0).AddModule(1, Me, CDbl(tmpModules(0).ModuleMCOV))
							unitCounter = unitCounter + 1
							tmpModules.Remove(tmpModules(0))
						End If
						Dim mcovTmp1 As Double = 0.0
						Dim mcovTmp2 As Double = 0.0
						While (tmpModules.Count > 0)
							m_units.Add(New ArresterUnit(unitCounter))
							m_units(unitCounter - 1).AddModule(1, Me, CDbl(tmpModules(0).ModuleMCOV))
							m_units(unitCounter - 1).AddModule(2, Me, CDbl(tmpModules(1).ModuleMCOV))
							tmpModules.Remove(tmpModules(0))
							tmpModules.Remove(tmpModules(0))
							unitCounter = unitCounter + 1
						End While
					End If

				Else
					MsgBox("Unable to lookup modules for this catalog number. No modules found.")
				End If
			Finally
				ModuleLookupReader.Close()
			End Try

		End If

		Return True

	End Function

	Private Sub lookupPorcelainCrossRef()
		Dim intRating As Integer = CInt(Rating)
		If (Type = Types.PORC) Then
			Dim PorcCrossRefQuery As String = "SELECT CrossedFamily FROM CrossRefPorc WHERE FamilyToCross = '" & Family & "' AND LowRating <= " & intRating & " AND HighRating >= " & intRating		'    20Apr2012  DLimburg
			Dim PorcCrossRefCommand As New SqlCommand(PorcCrossRefQuery, GRoomConnection)
			Dim crossRef As String = CStr(PorcCrossRefCommand.ExecuteScalar())
			If (Not crossRef Is Nothing) Then
				m_crossedFamily = crossRef
			Else
				m_crossedFamily = "NA"
			End If
		End If
	End Sub



#End Region

#Region "Housekeeping Methods"
	Public Sub SetWorkOrder(ByRef woIn As WorkOrder)
		m_workOrder = woIn
	End Sub


	Public Sub ClearModules()
		m_units.Clear()
		modules_set = False
	End Sub
	Public Sub Reset()
		m_type = -2
		m_catNumStr = Nothing
		m_buildPath = -2
		m_mcov = -2
		m_family = Nothing
		m_rating = Nothing
		m_classification = Nothing
		m_RingShedCode = Nothing
		m_porcColorCode = Nothing
		'-2: Not Parsed, -1 = Error, 0 - False, 1 = True
		m_isCubicleMount = -2
		m_stud = Nothing
		m_diskType = Nothing
		m_partnerDiskType = Nothing
		m_process = Nothing
		m_componentTable = Nothing
		m_mcovTable = Nothing
		m_mounting = Nothing
		m_namePlate = Nothing
		ClearModules()
	End Sub
#End Region

End Class
