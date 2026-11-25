Option Strict On
Imports System.Data.SqlClient
Public Class Nameplate

    Private m_catNum As CatNum
    Private m_pressureRelief As Integer = -2
    Private m_class As String = Nothing
    Private m_certification As String = Nothing
    Private m_mcov As String = Nothing
    Private m_rating As String = Nothing
    Private m_activeModule As ArresterModule = Nothing


    Public Sub New(ByRef catNumIn As CatNum)
        m_catNum = catNumIn
    End Sub

    ReadOnly Property PressureRelief() As Integer
        Get
            If (m_pressureRelief = -2) Then
                LookupNameplate()
            End If
            Return m_pressureRelief
        End Get
    End Property

    ReadOnly Property ArresterClass() As String
        Get
            If (m_class Is Nothing) Then
                LookupNameplate()
            End If
            Return m_class
        End Get
    End Property

    ReadOnly Property Certification() As String
        Get
            If (m_certification Is Nothing) Then
                LookupNameplate()
            End If
            Return m_certification
        End Get
    End Property

    ReadOnly Property MCOV() As String
        Get
            If (m_mcov Is Nothing) Then
                LookupNameplate()
            End If
            Return m_mcov
        End Get
    End Property

    ReadOnly Property Rating() As String
        Get
            If (m_rating Is Nothing) Then
                LookupNameplate()
            End If
            Return m_rating
        End Get
    End Property


    Private Sub LookupNameplate()
		Dim cmd As SqlCommand
		Dim rdr As SqlDataReader

		If (m_catNum.BuildPath = CatNum.BuildPaths.UltraSIL OrElse m_catNum.BuildPath = CatNum.BuildPaths.DualRatedPolymerNew) Then
			Dim sql As String = "SELECT * FROM " & ULTRASIL_BASIC_ARRESTER_DATA & " WHERE " & _
			  " Family = '" & m_catNum.Family & "' AND Rating_Min <= " & m_catNum.Rating & " And Rating_Max >= " & m_catNum.Rating & " AND Cantilever_Strength_Option='" & m_catNum.CantileverStrength & "' And Mounting_Type = "

			If (CBool(m_catNum.IsCubicleMount)) Then
				sql = sql & "'Cubical'"
			Else
				sql = sql & "'Other'"
			End If

			cmd = New SqlCommand(sql, PolymerStationConnection)
			rdr = cmd.ExecuteReader()

		Else

			Dim nameplateQuery As String = GetNameplateQuery()


			cmd = New SqlCommand(nameplateQuery, GRoomConnection)
			rdr = cmd.ExecuteReader()
		End If

		If (rdr.Read) Then
			m_pressureRelief = CInt(rdr!PressureRelief)
			m_class = CStr(rdr!Class)
			m_certification = CStr(rdr!Certification)
			If (m_catNum.BuildPath = CatNum.BuildPaths.DualRatedAZG Or m_catNum.BuildPath = CatNum.BuildPaths.DualRatedPolymer) Then
				m_mcov = CStr(rdr!MCOV_Arrester)
				m_rating = CStr(rdr!ArresterRating)
			Else
				rdr.Close()	'Must close the nameplate lookup so we can use the DB for MCOV
				m_mcov = CStr(m_catNum.MCOV)
				m_rating = m_catNum.Rating
			End If
		Else
			MsgBox("Could not look up nameplate data for this catalog number. No matching nameplates were found.", MsgBoxStyle.Critical)

		End If

		rdr.Close()
    End Sub

    Private Function GetNameplateQuery() As String
        Dim tableName As String = "None"
        Dim hasRating As Boolean = False
        If (CBool(m_catNum.IsCubicleMount) And m_catNum.Type = CatNum.Types.POLY) Then
            tableName = NAMEPLATE_CUBE_POLY
            hasRating = True
        ElseIf (CBool(m_catNum.IsCubicleMount) And m_catNum.Type = CatNum.Types.PORC) Then
            tableName = NAMEPLATE_CUBE_PORC
		ElseIf (m_catNum.BuildPath = CatNum.BuildPaths.DualRatedAZG) Then
            '6June2012 RPost I think this is wrong, should reference DR_Porc
            tableName = NAMEPLATE_DR_PORC  'changed from poly to PORC 1/21/26 RVA
        ElseIf (m_catNum.BuildPath = CatNum.BuildPaths.StdAZG Or m_catNum.BuildPath = CatNum.BuildPaths.StdAZG5r0 Or _
            m_catNum.BuildPath = CatNum.BuildPaths.StdAZG5r1) Then
            tableName = NAMEPLATE_PORC
            hasRating = True
        ElseIf (m_catNum.BuildPath = CatNum.BuildPaths.StdPolymer) Then
            tableName = NAMEPLATE_POLY
            hasRating = True
		ElseIf m_catNum.BuildPath = CatNum.BuildPaths.DualRatedPolymer Then
			tableName = NAMEPLATE_DR_POLY
		End If
        Dim nameplateQuery As String
        If (m_catNum.BuildPath = CatNum.BuildPaths.DualRatedPolymer Or m_catNum.BuildPath = CatNum.BuildPaths.DualRatedAZG) Then
            nameplateQuery = "SELECT * FROM " & tableName & " "
            nameplateQuery = nameplateQuery & "WHERE CatalogNumber = '" & m_catNum.CatNumString & "' "
        Else
            nameplateQuery = "SELECT * FROM " & tableName & " "
            nameplateQuery = nameplateQuery & "WHERE Family = '" & m_catNum.Family & "' "
            If (hasRating) Then
                nameplateQuery = nameplateQuery & "AND MinRating <= " & m_catNum.Rating
                nameplateQuery = nameplateQuery & " AND MaxRating >= " & m_catNum.Rating & " "
            End If
            End If

            Return nameplateQuery
    End Function

    Private Sub SetActiveModule(ByRef activeModuleIn As ArresterModule)
        m_activeModule = activeModuleIn
    End Sub




End Class
