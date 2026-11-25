Option Strict On
Imports System.Data.SqlClient

Public Class ArresterModule

    Private m_moduleNumber As Integer
    Private m_moduleMCOV As String
    Private m_totalDisks As Integer
    Private m_num3kVDisks_gray As Integer
    Private m_num5kVDisks_gray As Integer
    Private m_num3kVDisks_calcined As Integer
    Private m_num5kVDisks_calcined As Integer
    Private m_moduleSections As Integer
    '-2: Not Parsed, -1 = Error
    Private m_cumV10CATMax As Integer = -2
    Private m_cumV10CATMin As Integer = -2
    Private m_catNum As CatNum = Nothing
    Private m_disks As New ArrayList
    Private m_serialNumber As Int64 = -2
    Private m_housing As String = Nothing
    Private m_shed As Integer = -2
    Private m_buildNotes As String = Nothing
    Private m_buildDate As Date = Nothing
    Private m_catQtyMapping As New List(Of CatQtyMapping)
    Private m_unit As ArresterUnit




    ReadOnly Property Unit As ArresterUnit
        Get
            Return m_unit
        End Get
    End Property

    ReadOnly Property Shed As Integer
        Get
            If (m_shed = -2) Then
                GetUnitShedAndMCOV()
            End If
            Return m_shed
        End Get
    End Property

    ReadOnly Property DiskQtyMapping As List(Of CatQtyMapping)
        Get
            GetCatQtyMappings()
            Return m_catQtyMapping
        End Get
    End Property

    ReadOnly Property Disks() As Disk()
        Get
            Return CType(m_disks.ToArray(GetType(Disk)), Disk())
        End Get
    End Property

    Property ModuleNumber() As Integer
        Get
            Return m_moduleNumber
        End Get
        Set(ByVal value As Integer)
            m_moduleNumber = value
        End Set
    End Property



    ReadOnly Property UnitNumber() As Integer
        Get
            Return m_unit.UnitNumber
        End Get
    End Property


    Property ModuleMCOV() As String
        Get
            Return m_moduleMCOV
        End Get
        Set(ByVal value As String)
            m_moduleMCOV = value
        End Set
    End Property

    ReadOnly Property ModuleMCOVString() As String
        Get
            Return MCOVToStringMCOV(m_moduleMCOV)
        End Get
    End Property

    Property TotalDisks() As Integer
        Get
            Return m_totalDisks
        End Get
        Set(ByVal value As Integer)
            m_totalDisks = value
        End Set
    End Property

    ReadOnly Property Num3kVDisks() As Integer
        Get
            Return Num3kVDisksGray + Num3kVDisksCalcined
        End Get
    End Property

    ReadOnly Property Num5kVDisks() As Integer
        Get
            Return Num5kVDisksGray + Num5kVDisksCalcined
        End Get
    End Property

    Property Num3kVDisksGray() As Integer
        Get
            Return m_num3kVDisks_gray
        End Get
        Set(ByVal value As Integer)
            m_num3kVDisks_gray = value
        End Set
    End Property

    Property Num5kVDisksGray() As Integer
        Get
            Return m_num5kVDisks_gray
        End Get
        Set(ByVal value As Integer)
            m_num5kVDisks_gray = value
        End Set
    End Property

    Property Num3kVDisksCalcined() As Integer
        Get
            Return m_num3kVDisks_calcined
        End Get
        Set(ByVal value As Integer)
            m_num3kVDisks_calcined = value
        End Set
    End Property

    Property Num5kVDisksCalcined() As Integer
        Get
            Return m_num5kVDisks_calcined
        End Get
        Set(ByVal value As Integer)
            m_num5kVDisks_calcined = value
        End Set
    End Property

    ReadOnly Property Cum_V10CATMax() As Integer
        Get
            If (m_cumV10CATMax = -2) Then
                getV10CatMinMax()
            End If
            Return m_cumV10CATMax
        End Get
    End Property

    ReadOnly Property Cum_V10CATMin() As Integer
        Get
            If (m_cumV10CATMin = -2) Then
                getV10CatMinMax()
            End If
            Return m_cumV10CATMin
        End Get
    End Property

    ReadOnly Property Housing() As String
        Get
            If (m_housing Is Nothing) Then
                GetHousing()
            End If
            Return m_housing
        End Get
    End Property

    ReadOnly Property BuildNotes() As String
        Get
            If (m_buildNotes Is Nothing) Then
                GetBuildNotes()
            End If
            Return m_buildNotes
        End Get
    End Property

    ReadOnly Property BuildDate As Date
        Get
            GetBuildDate()
            Return m_buildDate
        End Get
    End Property

    ReadOnly Property CatNumber As CatNum
        Get
            Return m_catNum
        End Get
    End Property

    ReadOnly Property OldModNumber As Integer
        Get
            Dim retVal As Integer
            If (m_catNum.NumModules Mod 2 = 1) Then
                If (m_unit.UnitNumber = 1 And ModuleNumber = 1) Then
                    retVal = 1
                Else
                    retVal = 2 * m_unit.UnitNumber + ModuleNumber - 3
                End If
            Else
                retVal = (2 * m_unit.UnitNumber) + ModuleNumber - 2
            End If


            Dim unit As ArresterUnit
            Dim modRow As ArresterModule
            Dim oldCounter As Integer = 1
            For Each unit In m_catNum.Units
                For Each modRow In unit.Modules
                    If (Me.moduleString() = modRow.moduleString) Then
                        retVal = oldCounter
                    End If
                    oldCounter = oldCounter + 1
                Next modRow
            Next unit
            Return retVal
        End Get
    End Property




    Function moduleString() As String
        Dim BuiltString As String = ""
        If CDbl(ModuleMCOV) > 0 Then
            BuiltString = m_unit.UnitLetter & "->" & "Unit " & m_unit.UnitNumber & "->" & m_unit.UnitMCOV & " kV MCOV->" & "Module " & ModuleNumber & "->" & ModuleMCOV & "kV MCOV"
        End If
        Return BuiltString
    End Function

    Public Sub New(ByRef unitParent As ArresterUnit, ByRef catNumIn As CatNum, ByVal moduleNumberIn As Integer, ByVal moduleMCOVIn As String)
        m_catNum = catNumIn
        m_moduleNumber = moduleNumberIn
        m_moduleMCOV = moduleMCOVIn
        m_unit = unitParent
    End Sub

    Public Function verifyDiskCounts(ByRef order As WorkOrder) As Boolean
        Dim success As Boolean = True
        If (m_totalDisks <= 0) Then
            success = False
            MsgBox("Must have at least 1 disk!", MsgBoxStyle.Critical)
        End If
        If (success And m_num3kVDisks_gray + m_num5kVDisks_gray + m_num3kVDisks_calcined + m_num5kVDisks_calcined <> m_totalDisks) Then
            success = False
            MsgBox("The disk category counts must add up to the total number of disks.", MsgBoxStyle.Critical)
        End If
        If (success And order.CatNum.Type = CatNum.Types.PORC) Then
            success = verifyPorcelainCounts(order)
        ElseIf (success And order.CatNum.Type = CatNum.Types.POLY) Then
            success = verifyPolyCounts(order)
        ElseIf (success) Then
            MsgBox("Catalog number is not a valid type!", MsgBoxStyle.Critical)
        End If
        Return success
    End Function

    Public Function verifyPorcelainCounts(ByRef order As WorkOrder) As Boolean
        Dim success As Boolean = True
        Dim maxAllowed, minAllowed As Integer
        Dim housingCode As String = Nothing
        Dim foundQuery As String = ""
        Dim minMaxQuery As String = ""

            foundQuery = "SELECT DiskQuantity_Total FROM " & DISK_QTY_PORC_CALCINED _
            & " WHERE DiskType='" & order.CatNum.DiskType & "' " & _
            "AND Design_MCOV='" & ModuleMCOV & "' " & _
            "AND (Family = '" & order.CatNum.Family & "' Or Family = '" & order.CatNum.CrossedFamily & "') AND DiskQuantity_Total = " & m_totalDisks & " AND DiskQuantity_3kV = " & _
            Num3kVDisks & " And DiskQuantity_5kV = " & Num5kVDisks & " ORDER BY DiskQuantity_Total"
        minMaxQuery = "SELECT MAX(DiskQuantity_Total) Max, MIN(DiskQuantity_Total) 'Min' FROM " & DISK_QTY_PORC_CALCINED _
& " WHERE DiskType='" & order.CatNum.DiskType & "' " & _
"AND Design_MCOV='" & ModuleMCOV & "' " & _
"AND (Family = '" & order.CatNum.Family & "' Or Family = '" & order.CatNum.CrossedFamily & " ')"
        ' 2017-07-07 RVA Commented out next two lines displaying queries for debugging
        'MsgBox(foundQuery)
        'MsgBox(minMaxQuery)
            If (Not dbMatch(foundQuery, Nothing)) Then
                Dim minMaxCommand As New SqlCommand(minMaxQuery, GRoomConnection)
                Dim minMaxResultSet As SqlDataReader = minMaxCommand.ExecuteReader()
                Try
                    If (minMaxResultSet.HasRows()) Then
                        minMaxResultSet.Read()
                    maxAllowed = CInt(minMaxResultSet!Max)
                    minAllowed = CInt(minMaxResultSet!Min)
                        MsgBox("The number of disks entered is not valid for this arrestor. Allowable quantities are " & _
                                minAllowed & " To " & maxAllowed, MsgBoxStyle.Critical)
                        success = False
                    Else
                        MsgBox("No disk data could be found for this family in the table.", MsgBoxStyle.Critical)
                        success = False
                    End If
                Finally
                    minMaxResultSet.Close()
                End Try
            End If



            If (success) Then
                Dim buildTable As String = ""
                Dim housingQuery As String = ""
                Dim fieldName As String = "Unit_" & m_unit.UnitNumber & "_Housing"
                If (order.CatNum.BuildPath = CatNum.BuildPaths.DualRatedAZG) Then
                    buildTable = BUILD_TABLE_DR_PORC
                    housingQuery = "SELECT " & fieldName & " FROM " & HOUSING_TABLE_DR_PORC & " WHERE CatalogNumber='" & order.CatNum.CatNumString & "' ;"
                Else
                If (CBool(order.CatNum.IsCubicleMount)) Then
                    buildTable = BUILD_TABLE_CUBE
                Else
                    buildTable = BUILD_TABLE_PORC_CALCINED
                End If
                    housingQuery = "SELECT " & fieldName & " FROM " & HOUSING_TABLE_PORC & " WHERE RingCode='" & order.CatNum.RingShedCode & order.CatNum.PorcColorCode & "';"
                End If

                Dim housingQueryCommand As SqlCommand = New SqlCommand(housingQuery, GRoomConnection)
                Dim housingQueryData As SqlDataReader = housingQueryCommand.ExecuteReader()
                Try
                    If (Not housingQueryData.HasRows()) Then
                        success = False
                        MsgBox("The porcelain housing code" & order.CatNum.RingShedCode & order.CatNum.PorcColorCode & "was not found in the build table.", MsgBoxStyle.Critical)
                    Else
                        housingQueryData.Read()
                    housingCode = CStr(housingQueryData.Item(fieldName))
                    End If
                Finally
                    housingQueryData.Close()
                End Try

                If (success) Then


                    Dim BuildTableQuery As String = "SELECT TotDisks FROM " & buildTable & " " & _
                    "WHERE PorcelainHousing='" & housingCode & "' " & _
                    "And DiskType = '" & order.CatNum.DiskType & "' And TotDisks = " & m_totalDisks
                    Dim BuildTableMinMaxQuery As String = "SELECT Min(TotDisks) 'Min', Max(TotDisks) 'Max' FROM " & buildTable & " " & _
                    "WHERE PorcelainHousing='" & housingCode & "' " & _
                    "And DiskType = '" & order.CatNum.DiskType & "'"

                    If (Not dbMatch(BuildTableQuery, Nothing)) Then
                        Dim BuildTableMinMaxCommand As SqlCommand = New SqlCommand(BuildTableMinMaxQuery, GRoomConnection)
                        Dim BuildTableMinMaxResults As SqlDataReader = BuildTableMinMaxCommand.ExecuteReader()
                        Try
                            BuildTableMinMaxResults.Read()
                        Dim BuildTableMax As Integer = CInt(BuildTableMinMaxResults!Max)
                        Dim BuildTableMin As Integer = CInt(BuildTableMinMaxResults!Min)
                            MsgBox("The number of disks entered does not fit in housing " & housingCode & ". The allowable range is" _
                                    & BuildTableMin & " To " & BuildTableMax & ".", MsgBoxStyle.Critical)
                            success = False
                        Finally
                            BuildTableMinMaxResults.Close()
                        End Try
                    End If

                End If
            End If


            Return success
    End Function

    Private Sub GetHousing()
        Dim housingQuery As String
        Dim fieldName As String = "Unit_" & m_unit.UnitNumber & "_Housing"
        If (m_catNum.BuildPath = CatNum.BuildPaths.DualRatedAZG) Then
            housingQuery = "SELECT " & fieldName & " FROM " & HOUSING_TABLE_DR_PORC & " WHERE CatalogNumber='" & m_catNum.CatNumString & "' ;"
        Else
            housingQuery = "SELECT " & fieldName & " FROM " & HOUSING_TABLE_PORC & " WHERE RingCode='" & m_catNum.RingShedCode & m_catNum.PorcColorCode & "';"
        End If

        Dim housingQueryCommand As SqlCommand = New SqlCommand(housingQuery, GRoomConnection)
        Dim housingQueryData As SqlDataReader = housingQueryCommand.ExecuteReader()
        Try
            If (Not housingQueryData.HasRows()) Then
                MsgBox("The porcelain housing code" & m_catNum.RingShedCode & m_catNum.PorcColorCode & "was not found in the build table.", MsgBoxStyle.Critical)
            Else
                housingQueryData.Read()
                m_housing = CStr(housingQueryData.Item(fieldName))
            End If
        Finally
            housingQueryData.Close()
        End Try

    End Sub

	Public Sub ShedSet(ByVal Sheds As Integer)
		m_shed = Sheds
	End Sub

	Private Sub GetUnitShedAndMCOV()

		If (m_catNum.BuildPath <> CatNum.BuildPaths.UltraSIL) Then


			Dim query As String
			Dim field As String
			Dim mcovField As String = "None"
			If (m_catNum.BuildPath = CatNum.BuildPaths.DualRatedPolymer) Then
				field = "Unit_" & m_unit.UnitNumber & "_Housing"
				query = "SELECT " & field & " " & _
				"FROM HOUSINGDRPOLY WHERE CatalogNumber='" & m_catNum.CatNumString & "' "
			Else
				field = "Housing_" & m_unit.UnitNumber
				mcovField = "MCOV_Unit_" & m_unit.UnitNumber
				query = "SELECT " & field & ", " & mcovField & " FROM " & m_catNum.MCOVTable & _
				" WHERE Family= '" & m_catNum.Family & "' " & _
				"AND string_MCOV='" & m_catNum.MCOVStr & "' " & _
				"AND ShedCode='" & m_catNum.RingShedCode & "' " & _
				"AND ArresterRating = " & m_catNum.Rating & " "
			End If

			Dim UnitShedCommand As New SqlCommand(query, GRoomConnection)
			Dim UnitShedReader As SqlDataReader = UnitShedCommand.ExecuteReader()
			Try
				If (UnitShedReader.Read()) Then

					m_shed = CInt(UnitShedReader.Item(field))
					If (m_catNum.BuildPath = CatNum.BuildPaths.StdPolymer) Then
						' m_unit.UnitMCOV = UnitShedReader.Item(mcovField)
					End If
				Else
					MsgBox("Could not find a housing for this unit!")
				End If
			Finally
				UnitShedReader.Close()
			End Try
		Else
			Dim sql As String = "SELECT HsgPairCnt_Unit" & UnitNumber & "_Module" & ModuleNumber & " ShedCode FROM " & MCOV_HOUSING_ULTRASIL & _
			" WHERE Family = '" & m_catNum.Family & "' And Housing_Code = '" & m_catNum.RingShedCode & "' And Disk_Type = '" & m_catNum.DiskDia & "' AND Cantilever_Strength_Option='" & m_catNum.CantileverStrength & "'"
			Dim cmd As New SqlCommand(sql, PolymerStationConnection)
			Dim rdr As SqlDataReader = cmd.ExecuteReader()
			Try
				If (rdr.Read()) Then
					m_shed = CInt(rdr!ShedCode)
				Else
					MsgBox("Could not find a housing for this unit!")
				End If
			Finally
				rdr.Close()
			End Try

		End If
	End Sub

    Public Function verifyPolyCounts(ByRef order As WorkOrder) As Boolean
        Dim success As Boolean = True

		If (m_catNum.BuildPath = CatNum.BuildPaths.UltraSIL OrElse m_catNum.BuildPath = CatNum.BuildPaths.DualRatedPolymerNew) Then
			Dim sqlDiskQty As String = "SELECT Disk_Quantity_Total, Disk_Quantity_3kV, Disk_Quantity_5kV FROM " & DISK_QTY_ULTRASIL _
			 & " WHERE Family='" & order.CatNum.Family & "' AND IR_Code = '" & m_catNum.IRCode & "' AND Disk_Type ='" & order.CatNum.DiskDia & "' " _
			 & " AND MCOV = '" & ModuleMCOV & "' And Disk_Quantity_Total = " & m_totalDisks & " And Disk_Quantity_3kV = " & Num3kVDisks & " And Disk_Quantity_5kV = " & Num5kVDisks & " AND Cantilever_Strength_Option='" & order.CatNum.CantileverStrength & "'"

            Dim sqlMinMax As String = "SELECT MIN(Disk_Quantity_Total) Min, MAX(Disk_Quantity_Total) Max FROM " & DISK_QTY_ULTRASIL _
             & " WHERE Family='" & order.CatNum.Family & "' AND IR_Code = '" & m_catNum.IRCode & "' AND Disk_Type ='" & order.CatNum.DiskDia & "' " _
             & " AND MCOV = " & ModuleMCOV & " AND Cantilever_Strength_Option='" & order.CatNum.CantileverStrength & "'"

			If (Not dbMatch(sqlDiskQty, PolymerStationConnection, Nothing)) Then
				Dim cmd As SqlCommand = New SqlCommand(sqlMinMax, PolymerStationConnection)
				Dim rdr As SqlDataReader = cmd.ExecuteReader()
				Try
					If rdr.Read() AndAlso Not IsDBNull(rdr!min) Then

						MsgBox("You've entered an invalid number of disks. The valid range is " & CStr(rdr!Min) & _
						  " to " & CStr(rdr!Max) & ".", MsgBoxStyle.Critical)
						success = False
					Else
						MsgBox("Unable to find any data that matches the disk type and quantity of disks entered.", MsgBoxStyle.Critical)
						success = False
					End If
				Finally
					rdr.Close()
				End Try
			End If

		Else

			Dim sqlQty As String = "SELECT DiskQuantity, DiskQuantity_3kV, DiskQuantity_5kV FROM " & DISK_QTY_POLY _
			& " WHERE DiskType='" & order.CatNum.DiskType & "' AND String_MCOV ='" & Me.ModuleMCOVString & "' " _
			& " AND Family = '" & order.CatNum.Family & "' And DiskQuantity = " & m_totalDisks & " AND DiskQuantity_3kV = " & Num3kVDisks & _
			" AND DiskQuantity_5kV = " & Num5kVDisks

			Dim sqlMinMax As String = "SELECT Max(DiskQuantity) 'Max', Min(DiskQuantity) 'Min' FROM " & DISK_QTY_POLY _
			& " WHERE DiskType='" & order.CatNum.DiskType & "' AND String_MCOV ='" & Me.ModuleMCOVString & "' " _
			& " AND Family = '" & order.CatNum.Family & "' "

			If (Not dbMatch(sqlQty, Nothing)) Then
				Dim cmd As SqlCommand = New SqlCommand(sqlMinMax, GRoomConnection)
				Dim rdr As SqlDataReader = cmd.ExecuteReader()
				Try
					If rdr.Read() AndAlso Not IsDBNull(rdr!min) Then

						MsgBox("You've entered an invalid number of disks. The valid range is " & CStr(rdr!Min) & _
						  " to " & CStr(rdr!Max) & ".", MsgBoxStyle.Critical)
						success = False
					Else
						MsgBox("Unable to find any data that matches the disk type and quantity of disks entered.", MsgBoxStyle.Critical)
						success = False
					End If
				Finally
					rdr.Close()
				End Try
			End If

		End If

        If (success) Then

            Dim Connection As SqlConnection = GRoomConnection
            Dim housingQuery As String = "SELECT 0"
            Dim housingField As String = "None"

            If (order.CatNum.BuildPath = CatNum.BuildPaths.DualRatedPolymer) Then
                housingField = "Unit_" & m_unit.UnitNumber & "_Housing"
                housingQuery = "SELECT " & housingField & " " & _
                "FROM " & HOUSING_DR_POLY & " WHERE CatalogNumber='" & order.CatNum.CatNumString & "' "
			ElseIf (order.CatNum.BuildPath = CatNum.BuildPaths.UltraSIL Or order.CatNum.BuildPath = CatNum.BuildPaths.DualRatedPolymerNew) Then
				housingField = "Housing_PartNumber"
                housingQuery = "SELECT " & housingField & " FROM " & MCOV_HOUSING_PARTNUM & " WHERE Cantilever_Strength_Option = '" & _
                  order.CatNum.CantileverStrength & "' And Housing_GradingRing_Option = '" & order.CatNum.HsgGrdRingCode & "' And Disk_Type = '" & order.CatNum.DiskDia & "' AND Cantilever_Strength_Option='" & order.CatNum.CantileverStrength & "'"
				Connection = PolymerStationConnection
            ElseIf (order.CatNum.BuildPath = CatNum.BuildPaths.StdPolymer) Then
                housingField = "Housing_" & m_unit.UnitNumber
                housingQuery = "SELECT " & housingField & " FROM " & order.CatNum.MCOVTable & _
                " WHERE Family= '" & order.CatNum.Family & "' " & _
                "AND string_MCOV='" & order.CatNum.MCOVStr & "' " & _
                "AND ShedCode='" & order.CatNum.RingShedCode & "' " & _
                "AND ArresterRating = " & order.CatNum.Rating & " "
            End If

			Dim cmd As SqlCommand = New SqlCommand(housingQuery, Connection)
			Dim rdr As SqlDataReader = cmd.ExecuteReader
			Try
				If (Not rdr.HasRows()) Then
					MsgBox("Unable to find a housing that matches this part.", MsgBoxStyle.Critical)
					success = False
				Else
					rdr.Read()
					If (m_catNum.BuildPath = CatNum.BuildPaths.UltraSIL Or m_catNum.BuildPath = CatNum.BuildPaths.DualRatedPolymerNew) Then
						m_housing = CStr(rdr.Item(housingField))

					Else
						m_moduleSections = CInt(rdr.Item(housingField))
					End If


				End If
			Finally
				rdr.Close()
			End Try
        End If

		If (success AndAlso (order.CatNum.BuildPath <> CatNum.BuildPaths.UltraSIL And order.CatNum.BuildPath <> CatNum.BuildPaths.DualRatedPolymerNew)) Then

			Dim DiskQtyQuery As String = "SELECT " & DISK_QTY_POLY & ".*, " & MODULE_COMPONENTS_POLY & ".Sections, " & MODULE_COMPONENTS_POLY & ".ModuleHeight," & _
				  " " & MODULE_COMPONENTS_POLY & ".EndElectrodeHgt , " & MODULE_COMPONENTS_POLY & ".MidElectrodes, " & MODULE_COMPONENTS_POLY & ".HeatSinks" & _
				  " FROM " & DISK_QTY_POLY & " LEFT OUTER JOIN " & MODULE_COMPONENTS_POLY & " ON " & _
				  " " & DISK_QTY_POLY & ".DiskQuantity = " & MODULE_COMPONENTS_POLY & ".TotalDiskCnt" & _
				  " AND " & DISK_QTY_POLY & ".DiskQuantity_3kV = (" & MODULE_COMPONENTS_POLY & ".Total3kVDiskCnt)" & _
				  " AND " & DISK_QTY_POLY & ".DiskQuantity_5kV = (" & MODULE_COMPONENTS_POLY & ".Total5kVDiskCnt)" & _
				  " WHERE (" & DISK_QTY_POLY & ".DiskType = '" & order.CatNum.DiskType & "')" & _
				  " AND (" & DISK_QTY_POLY & ".Design_MCOV = '" & m_moduleMCOV & "')" & _
				  " AND (" & DISK_QTY_POLY & ".Family = '" & order.CatNum.Family & "')" & _
				  " AND (" & MODULE_COMPONENTS_POLY & ".CP_3kVDiskCnt = " & m_num3kVDisks_calcined & ")" & _
				  " AND (" & MODULE_COMPONENTS_POLY & ".CP_5kVDiskCnt = " & m_num5kVDisks_calcined & ")" & _
				  " AND (" & MODULE_COMPONENTS_POLY & ".Gray_3kVDiskCnt = " & m_num3kVDisks_gray & ")" & _
				  " AND (" & MODULE_COMPONENTS_POLY & ".Gray_5kVDiskCnt = " & m_num5kVDisks_gray & ")" & _
				  " AND (" & MODULE_COMPONENTS_POLY & ".Stud = '" & order.CatNum.Stud & "')" & _
				  " AND (" & MODULE_COMPONENTS_POLY & ".Sections = " & m_moduleSections & ")"

			If (Not dbMatch(DiskQtyQuery, Nothing)) Then
				MsgBox("The Number of disks specified will not fit in any housing for this unit!", MsgBoxStyle.Critical)
				success = False
			End If
		End If


        Return success


    End Function

    Private Sub getV10CatMinMax()
        Dim quantityTable As String = "None"
		Dim sql As String = "SELECT 0"
		Dim diskQtyField As String = "None"

		If (m_catNum.BuildPath = CatNum.BuildPaths.UltraSIL Or m_catNum.BuildPath = CatNum.BuildPaths.DualRatedPolymerNew) Then

			sql = "SELECT Total_V10_Cat_Min Cum_V10CAT_Min, Total_V10_Cat_Max Cum_V10CAT_Max FROM " & DISK_QTY_ULTRASIL & " WHERE Disk_Type = '" & m_catNum.DiskDia & "' AND Family = '" & m_catNum.Family & _
			  "' AND MCOV = " & ModuleMCOV & " And Disk_Quantity_Total = " & TotalDisks & " And Disk_Quantity_3kV = " & Num3kVDisks & " And Disk_Quantity_5kV = " & Num5kVDisks & " AND Cantilever_Strength_Option='" & m_catNum.CantileverStrength & "'"
		Else
			If (m_catNum.Type = CatNum.Types.PORC) Then
				quantityTable = DISK_QTY_PORC_CALCINED
				diskQtyField = "DiskQuantity_Total"
			ElseIf (m_catNum.Type = CatNum.Types.POLY) Then
				quantityTable = DISK_QTY_POLY
				diskQtyField = "DiskQuantity"
			End If


			sql = "SELECT * FROM " & quantityTable & " WHERE DiskType='" & m_catNum.DiskType & "' " & _
			"AND Design_MCOV='" & ModuleMCOV & "' And " & _
			m_catNum.SqlFamily & " AND " & diskQtyField & " = '" & TotalDisks & "'"

			sql = sql & " AND DiskQuantity_3kV = " & (Num3kVDisksGray + Num3kVDisksCalcined) & _
			" AND DiskQuantity_5kV = " & (Num5kVDisksGray + Num5kVDisksCalcined)
			sql = sql & "  Order By Family ASC"
		End If

		Dim cmd As SqlCommand
		If (m_catNum.BuildPath = CatNum.BuildPaths.UltraSIL Or m_catNum.BuildPath = CatNum.BuildPaths.DualRatedPolymerNew) Then
			cmd = New SqlCommand(sql, PolymerStationConnection)
		Else
			cmd = New SqlCommand(sql, GRoomConnection)
		End If

		Dim rdr As SqlDataReader = cmd.ExecuteReader()

		Try
			If rdr.Read Then

				m_cumV10CATMin = CInt(rdr("Cum_V10CAT_Min"))
				m_cumV10CATMax = CInt(rdr("Cum_V10CAT_Max"))
			Else
				MsgBox("Could not obtain the minimum and maximum V10CAT values from the database.", MsgBoxStyle.Critical)
				m_cumV10CATMin = -1
				m_cumV10CATMax = -1
			End If
		Finally
			rdr.Close()
		End Try

    End Sub

    Public Function AddDisk(ByVal diskBatch As Integer, ByVal diskNum As Integer) As Boolean
        Dim retVal As Boolean = True
        Dim diskToAdd As New Disk(Me, CStr(diskBatch), diskNum)
        If (m_disks.Contains(diskToAdd)) Then
            retVal = False
            MsgBox("Duplicate disks found! Batch:" & diskBatch & " Num:" & diskNum, MsgBoxStyle.Critical)
        ElseIf (diskToAdd.ValidForUse = True And Not CrossRefType(diskToAdd.DiskType)) Then
            MsgBox("Not an allowed disk for this type!")
            retVal = False
        Else
            m_disks.Add(diskToAdd)
            retVal = True
        End If

        Return retVal
    End Function

    Public Function AddDisk(ByVal diskToAdd As Disk) As Boolean
        Dim retVal As Boolean = True
        If (m_disks.Contains(diskToAdd)) Then
            retVal = False
            MsgBox("Duplicate disks found! Batch:" & diskToAdd.Batch & " Num:" & diskToAdd.DiskNum, MsgBoxStyle.Critical)
        ElseIf (diskToAdd.ValidForUse And Not CrossRefType(diskToAdd.DiskType)) Then
            MsgBox("Not an allowed disk for this type!")
            retVal = False
        Else
            m_disks.Add(diskToAdd)
            retVal = True
        End If
        Return retVal
    End Function

    Public Function CrossRefType(ByVal buildType As String) As Boolean
        Dim retVal As Boolean = False
        If (buildType.Length >= 4 AndAlso m_catNum.DiskType.Substring(0, 4) = buildType.Substring(0, 4)) Then
            retVal = True
        ElseIf (m_catNum.PartnerDiskType = buildType) Then
            retVal = True
        ElseIf (m_catNum.DiskType = "AT32E") Then
            If (buildType = "AT118G") Then
                retVal = True
            ElseIf (m_catNum.Type = CatNum.Types.PORC And buildType = "AT118X") Then
                retVal = True
            End If
            If (m_catNum.Type = CatNum.Types.PORC And (buildType = "AT52G" OrElse buildType = "AT52E")) Then
                retVal = True
            End If
        ElseIf (m_catNum.DiskType = "AT33E") Then
            If (buildType = "AT112G") Then
                retVal = True
            ElseIf (m_catNum.Type = CatNum.Types.PORC And buildType = "AT112X") Then
            End If
            If (m_catNum.Type = CatNum.Types.PORC And (buildType = "AT53G" OrElse buildType = "AT53E")) Then
                retVal = True
            End If
        ElseIf (m_catNum.DiskType = "AT34E") Then
            If (buildType = "AT108G") Then
                retVal = True
            ElseIf (m_catNum.Type = CatNum.Types.PORC And buildType = "AT108X") Then
            End If
            If (m_catNum.Type = CatNum.Types.PORC And (buildType = "AT54G" OrElse buildType = "AT54E")) Then
                retVal = True
            End If
        ElseIf (m_catNum.DiskType = "AT32G") Then
            If (buildType = "AT32E" OrElse (m_catNum.Type = CatNum.Types.PORC And (buildType = "AT52G" OrElse buildType = "AT52E"))) Then
                retVal = True
            End If
        ElseIf (m_catNum.DiskType = "AT33G" OrElse (m_catNum.Type = CatNum.Types.PORC And (buildType = "AT53G" OrElse buildType = "AT53E"))) Then
            If (buildType = "AT33E") Then
                retVal = True
            End If
        ElseIf (m_catNum.DiskType = "AT34G" OrElse (m_catNum.Type = CatNum.Types.PORC And (buildType = "AT54G" OrElse buildType = "AT54E"))) Then
            If (buildType = "AT34E") Then
                retVal = True
            End If

        End If
        Return retVal

    End Function

    Public Function Build(ByVal l As Login) As Boolean
        Dim trans As SqlClient.SqlTransaction = GRoomConnection.BeginTransaction(IsolationLevel.ReadCommitted)

        Dim success As Boolean = True
        Dim dsk As Disk
        For Each dsk In Disks
            If (success) Then
                Try 'JP 05/23/2019
                    success = dsk.ConsumeDisk(trans)
                Catch sqlEx As SqlException 'JP 05/23/2019
                    'If there is an SQL exception log it to a file 'JP 05/23/2019
                Catch ex As Exception 'JP 05/23/2019
                    'If there is an exception log it to a file 'JP 05/23/2019
                End Try 'JP 05/23/2019

            End If
        Next

        Dim builtUnitsInsert As String = "INSERT INTO BuiltUnits VALUES('" & m_catNum.SerialNumber & "','" & _
         OldModNumber & "','" & Now & "','" & TotalDisks & _
         "'," & l.EmployeeID & "," & Num3kVDisksGray & "," & _
         Num5kVDisksGray & "," & Num3kVDisksCalcined & _
         "," & Num5kVDisksCalcined & ");"

        Dim builtUnitsCommand As New SqlCommand(builtUnitsInsert, GRoomConnection)
        builtUnitsCommand.Transaction = trans
        If (builtUnitsCommand.ExecuteNonQuery() <> 1) Then
            MsgBox("There was a problem inserting the new unit into the BuiltUnits table.", MsgBoxStyle.Critical)
            success = False
        End If

        Dim updateInfoQuery As String = "UPDATE WO_Info SET Unit" & OldModNumber & _
        "NumBuilt=Unit" & OldModNumber & "NumBuilt + 1 WHERE WO_Number='" & _
        m_catNum.WorkOrder.WorkOrderNum & "' and CatalogNumber='" & m_catNum.CatNumString & "' ;"   ' 30Apr2012 DLimburg - changed from and CatalogNumber='" & m_catNum.GivenCatNum & "'
        '8May2012 RPost SET Unit & !UnitNumber! changed to OldModNumber

        Dim updateInfoCommand As New SqlCommand(updateInfoQuery, GRoomConnection)
        updateInfoCommand.Transaction = trans
        If (updateInfoCommand.ExecuteNonQuery() <> 1) Then
            MsgBox("There was a problem updating the WO_Info table with the new build information.", MsgBoxStyle.Critical)
            success = False
        End If

        If (success) Then
            trans.Commit()
        Else
            trans.Rollback()
        End If

        Return success
    End Function

    Public Function checkDiskDetail() As Boolean
        Dim qtyTable As String = "None"
        Dim query As String
        Dim typeDepQuery As String = ""
        Dim success As Boolean = False
        If (m_catNum.Type = CatNum.Types.PORC) Then
            If (m_num3kVDisks_calcined + m_num5kVDisks_calcined > 0) Then
                qtyTable = DISK_QTY_PORC_CALCINED
                typeDepQuery = "AND DiskQuantity_Total =" & m_totalDisks & " "
            ElseIf (m_num3kVDisks_gray > 0) Then
                typeDepQuery = "AND DiskQuantity=" & m_totalDisks & " "
            End If

            typeDepQuery = "AND DiskQuantity_Total =" & m_totalDisks & " "
        ElseIf (m_catNum.Type = CatNum.Types.POLY) Then
            qtyTable = DISK_QTY_POLY
            typeDepQuery = "AND DiskQuantity=" & m_totalDisks & " "
        End If

        query = "SELECT * FROM " & qtyTable & " WHERE DiskType='" & m_catNum.DiskType & "' " & _
        "AND Design_MCOV='" & ModuleMCOV & "' " & _
        "AND Family = '" & m_catNum.Family & "' " & typeDepQuery & _
        " AND DiskQuantity_3kV = " & Num3kVDisks & " AND DiskQuantity_5kV = " & Num5kVDisks


        Dim diskDetailCommand As SqlCommand = New SqlCommand(query, GRoomConnection)
        Dim diskDetailReader As SqlDataReader = diskDetailCommand.ExecuteReader()
        Try
            If (diskDetailReader.HasRows()) Then
                Dim LowCum As Integer
                Dim HighCum As Integer
                Integer.TryParse(CStr(diskDetailReader!Cum_V10CAT_Min), LowCum)
                Integer.TryParse(CStr(diskDetailReader!Cum_V10CAT_High), HighCum)

            Else
                MsgBox("The Cum_V10CAT_Min and Cum_V10CAT_Max cannot be retrieved for " & vbCrLf & _
                "Disk Type = " & m_catNum.DiskType & " and MCOV = " & m_unit.UnitMCOV & ".", MsgBoxStyle.Critical)
            End If

        Finally
            diskDetailReader.Close()
        End Try

        Return success

    End Function

    Public Sub ClearDisks()
        m_disks.Clear()
    End Sub

    Private Sub GetBuildNotes()
        Dim BuildNotesQuery As String = "SELECT 0"
        If (m_catNum.Type = CatNum.Types.POLY) Then
            If (m_catNum.Mounting = "D") Then
                BuildNotesQuery = "SELECT * FROM " & BUILD_NOTES_POLY & " WHERE CatalogNumber = '" & m_catNum.CatNumString & "' "
            End If
        ElseIf (m_catNum.Type = CatNum.Types.PORC) Then
            BuildNotesQuery = "SELECT * FROM BuildNotesPorc WHERE HousingCode = '" & Housing & "'"
        End If
        Dim BuildNotesCommand As New SqlCommand(BuildNotesQuery, GRoomConnection)
        Dim BuildNotesReader As SqlDataReader = BuildNotesCommand.ExecuteReader()
        Try
            If (BuildNotesReader.HasRows) Then
                BuildNotesReader.Read()
                Try
                    m_buildNotes = CStr(BuildNotesReader!BuildNotes)
                Catch ex As Exception
                    m_buildNotes = "NA"
                End Try

            Else
                m_buildNotes = "NA"
            End If
        Finally
            BuildNotesReader.Close()
        End Try

    End Sub

    Public Function Lookup() As Boolean
        Dim success As Boolean = True
        Dim LookupDisksQuery As String = "SELECT * FROM ConsumedDisks WHERE SerialNumber = '" & m_serialNumber & "' AND UnitNumber = " & m_unit.UnitNumber
        Dim LookupDisksCmd As New SqlCommand(LookupDisksQuery, GRoomConnection)
        Dim LookupDisksReader As SqlDataReader = LookupDisksCmd.ExecuteReader()
        Try
            If (LookupDisksReader.HasRows()) Then
                While (LookupDisksReader.Read())
                    m_disks.Add(New Disk(Me, CStr(LookupDisksReader!Batch_Number), CInt(LookupDisksReader!Disk_Number), True))
                End While
            Else
                MsgBox("Unable to find any disks for this module.")
                success = False
            End If
        Finally
            LookupDisksReader.Close()
        End Try

        Dim BuiltUnitsQuery As String = "SELECT * FROM BuiltUnits WHERE SerialNumber = '" & m_catNum.SerialNumber & "' AND Unit = " & m_unit.UnitNumber
        Dim BuiltUnitsCmd As New SqlCommand(BuiltUnitsQuery, GRoomConnection)
        Dim BuiltUnitsReader As SqlDataReader = BuiltUnitsCmd.ExecuteReader()
        Try
            If (BuiltUnitsReader.HasRows()) Then
                BuiltUnitsReader.Read()
                ' If (Not IsDBNull(BuiltUnitsReader!NumOf3kVDisks)) Then
                m_num3kVDisks_gray = CInt(BuiltUnitsReader!NumOf3kVDisks)
                ' End If
                ' If (Not IsDBNull(BuiltUnitsReader!NumOf5kVDisks)) Then
                m_num5kVDisks_gray = CInt(BuiltUnitsReader!NumOf5kVDisks)
                'End If
                If (Not IsDBNull(BuiltUnitsReader!NumOf3kVDisksCalcined)) Then
                    m_num3kVDisks_calcined = CInt(BuiltUnitsReader!NumOf3kVDisksCalcined)
                End If
                If (Not IsDBNull(BuiltUnitsReader!NumOf5kVDisksCalcined)) Then
                    m_num5kVDisks_calcined = CInt(BuiltUnitsReader!NumOf5kVDisksCalcined)
                End If
            Else
                MsgBox("Unable to find the unit in the BuiltUnits table!")
                success = False
            End If
        Finally
            BuiltUnitsReader.Close()
        End Try

        Return success

    End Function

    Private Sub GetCatQtyMappings()
        m_catQtyMapping.Clear()
        Dim row As Disk
        For Each row In m_disks
            Dim mapping As New CatQtyMapping(row.Category, 1, 0, row.DiskType)
            If (m_catQtyMapping.Contains(mapping)) Then
                Dim foundMapping As CatQtyMapping = m_catQtyMapping.Item(m_catQtyMapping.IndexOf(mapping))
                foundMapping.assertedQty += mapping.assertedQty
                m_catQtyMapping.RemoveAt(m_catQtyMapping.IndexOf(mapping))
                m_catQtyMapping.Add(foundMapping)
            Else
                m_catQtyMapping.Add(mapping)
            End If
        Next
        m_catQtyMapping.Sort()
    End Sub

    Private Sub GetBuildDate()
        Dim lookupDateQuery As String = "SELECT * FROM BuiltUnits WHERE SerialNumber = '" & m_catNum.SerialNumber & "' AND Unit = " & m_unit.UnitNumber
        Dim lookupDateCommand As New SqlCommand(lookupDateQuery, GRoomConnection)
        Dim lookupDateDataReader As SqlDataReader = lookupDateCommand.ExecuteReader
        Try
            If (lookupDateDataReader.HasRows()) Then
                lookupDateDataReader.Read()
                m_buildDate = CDate(lookupDateDataReader!TimeBuilt)
            Else
                MsgBox("Unable to find a build date for the current work order/unit.")
            End If
        Finally
            lookupDateDataReader.Close()
        End Try

    End Sub

    Public Function NumberBuilt() As Integer
        Dim numBuiltField As String = "Unit" & OldModNumber & "NumBuilt"
        Dim numBuiltQuery As String = "SELECT " & numBuiltField & " FROM WO_Info WHERE WO_Number = " & m_catNum.WorkOrder.WorkOrderNum
        Dim numBuiltCommand As SqlCommand = New SqlCommand(numBuiltQuery, GRoomConnection)
        Dim numBuilt As Integer = 0
        Integer.TryParse(CStr(numBuiltCommand.ExecuteScalar()), numBuilt)
        Return numBuilt
    End Function

    Public Function IsBuilt() As Boolean
        Dim query As String = "SELECT * FROM " & BUILT_UNITS & " WHERE SerialNumber = '" & CatNumber.SerialNumber & "' And Unit = " & OldModNumber
        Return dbMatch(query, Nothing)
    End Function

    Private Function MCOVToStringMCOV(ByVal realMCOV As String) As String
        Dim realMCOVInt As Integer = 0
        Dim retMCOV As String = ""
        Dim retMCOVDbl As Double

        '    Integer.TryParse(realMCOV, realMCOVInt)    ' this does not work when dealing with a value with a decimal place
        retMCOVDbl = CDbl(realMCOV)     ' 11Apr2012 Added to convert string to a value
        realMCOVInt = CInt(Math.Truncate(retMCOVDbl))  ' 11Apr2012 Added to convert decimal value to integer for formating in next line.
        Return String.Format("{0:000}", realMCOVInt)

    End Function




End Class
