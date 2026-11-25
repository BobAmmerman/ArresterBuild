Option Strict On
Imports System.Data.SqlClient
Public Class WorkOrder

	Private m_workOrderNum As String

	Public m_catNum As New CatNum

	Private m_startSerial As Long
	Private m_numUnits As Integer
	Private m_buildDate As Date
	'    Private m_numBuilt As Integer


	Structure BasicUnit
		Public Sub New(ByVal serNumIn As Long, ByVal unitNumIn As Integer)
			serNum = serNumIn
			unitNum = unitNumIn
		End Sub
		Public serNum As Long
		Public unitNum As Integer
	End Structure

	Sub New()
		m_catNum.SetWorkOrder(Me)
	End Sub

	Property WorkOrderNum() As String
		Get
			Return m_workOrderNum
		End Get
		Set(ByVal value As String)
			m_workOrderNum = value
		End Set
	End Property

	Property StartSerial() As Long
		Get
			Return m_startSerial
		End Get
		Set(ByVal value As Long)
			m_startSerial = value
		End Set
	End Property

	Property NumUnits() As Integer
		Get
			Return m_numUnits
		End Get
		Set(ByVal value As Integer)
			m_numUnits = value
		End Set
	End Property

	Property CatNum() As CatNum
		Get
			Return m_catNum
		End Get
		Set(ByVal value As CatNum)
			m_catNum = value
		End Set
	End Property

	Private Function checkSerialNumbers() As Boolean
		Dim successful As Boolean = True
		Dim Sql As String = "SELECT * FROM BuiltUnits "
		Dim SQLWhere As String = "WHERE SerialNumber = '" & m_startSerial & "' "
		If m_numUnits > 1 Then
			'more than one serial number in the list.
			Dim iLoop As Integer
			For iLoop = 1 To m_numUnits - 1
				'add each serial number to the list
				SQLWhere = SQLWhere & "OR SerialNumber = '" & m_startSerial + iLoop & "' "
			Next iLoop
		End If
		Dim myReader As SqlDataReader = New SqlCommand(Sql & SQLWhere, GRoomConnection).ExecuteReader()
		Try
			If (myReader.HasRows()) Then
				Dim errString As String = "Duplicate serial numbers exist in the built units table:"
				While myReader.Read()
					errString = errString & CStr(myReader!SerialNumber) & ","
				End While
				MsgBox(errString, MsgBoxStyle.Critical)
				successful = False

			End If
		Finally
			myReader.Close()
		End Try

		Dim myReader2 As SqlDataReader = New SqlCommand("SELECT * FROM WO_SerialNumbers " & SQLWhere, GRoomConnection).ExecuteReader()
		Try
			If (myReader2.HasRows()) Then
				Dim errString As String = "The following serial numbers have already been assigned to a work order.:"
				While myReader2.Read()
					errString = errString & CStr(myReader2!SerialNumber) & ","
				End While
				MsgBox(errString, MsgBoxStyle.Critical)
				successful = False

			End If
		Finally
			myReader2.Close()
		End Try
		Return successful
	End Function

	Public Function Exists() As Boolean
		Return Exists(Nothing)
	End Function



	Public Function Exists(ByVal trans As SqlTransaction) As Boolean
		Dim retVal As Boolean = True
		Dim woLookupCount As String = "SELECT COUNT(*) FROM WO_Info WHERE WO_Number = '" & WorkOrderNum & "'"
		Dim woLookupCountCommand As New SqlCommand(woLookupCount, GRoomConnection)
		If (Not trans Is Nothing) Then
			woLookupCountCommand.Transaction = trans
		End If
		Dim count As Integer = CInt(woLookupCountCommand.ExecuteScalar())

		Dim woLookupQuery As String = "SELECT * FROM WO_Info WHERE WO_Number = '" & WorkOrderNum & "'"
		Dim woLookupCommand As New SqlCommand(woLookupQuery, GRoomConnection)
		If (Not trans Is Nothing) Then
			woLookupCommand.Transaction = trans
		End If
		Dim woLookupReader As SqlDataReader = woLookupCommand.ExecuteReader()
		If (count = 1) Then
			woLookupReader.Read()
			' 20Apr2012 DLimburg - added line to populate given number, this is done to handle and maintain shortcuts numbers
			m_catNum.GivenCatNum = CStr(woLookupReader!CatalogNumber)
			m_catNum.CatNumString = CStr(woLookupReader!CatalogNumber)
			m_numUnits = CInt(woLookupReader!NumArrestersRequired)
			retVal = True
			woLookupReader.Read()
		ElseIf (count > 1) Then
			MsgBox("Duplicate work orders found. Please contact the database administrator.", MsgBoxStyle.Critical, "Error")
			retVal = False
		Else
			retVal = False
			retVal = False
		End If
		woLookupReader.Close()
		Return retVal
	End Function

	Public Function Validate() As Boolean
		Dim finishedAndValid As Boolean = False
		Dim currentWorkOrder As String = m_workOrderNum
		If (currentWorkOrder.Length = 8 Or currentWorkOrder.Length = 7) Then
			finishedAndValid = True
		Else
			MsgBox("Work order must have either 7 or 8 characters!")
			finishedAndValid = False
		End If

		If finishedAndValid = True Then
			If (Not IsNumeric(currentWorkOrder)) Then
				'   If Integer.TryParse(, parseTest) = False Then
				MsgBox("The work order you've entered is not a valid number.")
				finishedAndValid = False
			End If
		End If


		Return finishedAndValid

	End Function

	Public Function AddWorkOrder() As Boolean
		Dim success As Boolean = True

		If (success = True And m_catNum.CatNumString.Length < 10 And m_catNum.CatNumString.Length <> 6) Then
			MsgBox("Catalog number is too short. It needs to be at least 10 charaters.", MsgBoxStyle.Critical)
			success = False
		End If

		If (Exists()) Then
			MsgBox("Work order already found in database. Cannot add it again.", MsgBoxStyle.Critical)
			success = False
		End If

		If (Not m_catNum.Validate()) Then
			MsgBox("The catalog number could not be verified.", MsgBoxStyle.Critical)
			success = False
		End If

		If (success And Not checkSerialNumbers()) Then
			success = False
		End If



		If (success) Then
			Dim trans As SqlClient.SqlTransaction = GRoomConnection.BeginTransaction(IsolationLevel.ReadCommitted)


			' 20Apr2012 DLimburg removed line below and replace with line that saves the given catalog number, this is done to maintain the shortcut
			'           Dim query As String = "INSERT INTO WO_Info VALUES('" & m_workOrderNum & "','" & m_catNum.CatNumString & "','" & System.DateTime.Now() & "' ,'" & m_numUnits & "','1','0','0','0','0','0','0')"
			' 30Apr2012 DLimburg undid this change as the tester program expects a full caatalog number and cannot work with a shortcut
			'           Dim query As String = "INSERT INTO WO_Info VALUES('" & m_workOrderNum & "','" & m_catNum.CatNumString & "','" & System.DateTime.Now() & "' ,'" & m_numUnits & "','1','0','0','0','0','0','0')"
			'                                                                                            & m_catNum.GivenCatNum &
			Dim query As String = "INSERT INTO WO_Info VALUES('" & m_workOrderNum & "','" & m_catNum.CatNumString & "','" & System.DateTime.Now() & "' ,'" & m_numUnits & "','1','0','0','0','0','0','0')"

			Dim woInfoInsertCmd As New SqlCommand(query, GRoomConnection)
			woInfoInsertCmd.Transaction = trans
			woInfoInsertCmd.ExecuteNonQuery()

			Dim counter As Integer
			For counter = 0 To m_numUnits - 1
				Dim command As New SqlCommand("INSERT INTO WO_SerialNumbers VALUES('" & m_workOrderNum & "','" & m_startSerial + counter & "')", GRoomConnection)
				command.Transaction = trans
				command.ExecuteNonQuery()
			Next counter
			If (success) Then
				trans.Commit()
			Else
				trans.Rollback()
			End If
			Return (success)
		End If
	End Function

	Public Function DeleteWorkOrder() As Boolean
		Dim success As Boolean = True
		If (Me.Validate() And Me.Exists()) Then


			Dim sureMsg As String = ""
			'units are already built, query user on deleting units.
			sureMsg = "The following Units have already been built for the work order." & vbCrLf & _
			"Do you wish to Unbuild all of these units?"

			Dim builtUnits As List(Of BasicUnit) = GetSerialNumbersUnitsBuilt()
			Dim UnitRow As BasicUnit
			For Each UnitRow In builtUnits
				sureMsg = sureMsg & vbCrLf & "Serial: " & UnitRow.serNum & " Unit:" & Chr(UnitRow.unitNum + 64) & vbCrLf
			Next UnitRow


			Dim Result As MsgBoxResult = MsgBox(sureMsg, vbYesNo, "Unbuild Units?")

			If Result = vbNo Then
				'user elected not to unbuild units.
				MsgBox("Units will not be unbuilt and Work order will not be deleted.")
				success = False
			End If

			Dim trans As SqlTransaction = GRoomConnection.BeginTransaction(IsolationLevel.ReadCommitted)
			For Each UnitRow In builtUnits
				If (success) Then
					success = UnbuildUnit(UnitRow.serNum, UnitRow.unitNum, trans)
				End If
			Next UnitRow


			'delete wo
			Dim deleteWOQuery As String = "DELETE FROM WO_Info " & _
			"WHERE WO_Number = '" & WorkOrderNum & "' "
			Dim deleteWOCommand As New SqlCommand(deleteWOQuery, GRoomConnection)
			deleteWOCommand.Transaction = trans

			If (deleteWOCommand.ExecuteNonQuery() <= 0) Then
				success = False
				MsgBox("Could not delete the work order from the work order info database.")
			End If

			'delete wo From wo_Serials
			Dim deleteWOSerialsQuery As String = "DELETE FROM WO_SerialNumbers " & _
			"WHERE WO_Number = '" & WorkOrderNum & "' "
			Dim deleteWOSerialsCommand As New SqlCommand(deleteWOSerialsQuery, GRoomConnection)
			deleteWOSerialsCommand.Transaction = trans

			If (deleteWOSerialsCommand.ExecuteNonQuery() <= 0) Then
				success = False
				MsgBox("Could not erase the serial number from the WO_Serials database.")
			End If


			If (success) Then
				trans.Commit()
				MsgBox("Work Order Successfully Deleted")
			Else
				trans.Rollback()
				MsgBox("Could not delete work order! No changes will be saved.")
			End If
		End If

	End Function

	Public Function UnbuildUnit(ByVal SerialNumber As Long, ByVal UnitNum As Integer, ByRef trans As SqlTransaction) As Boolean
		Dim success As Boolean = True
		Dim internalTransaction As Boolean = False
		If (Not dbMatch("SELECT * FROM ConsumedDisks " & _
		"WHERE SerialNumber = '" & SerialNumber & "' " & _
		"AND UnitNumber = " & UnitNum & " ", trans)) Then
			MsgBox("Could not find the serial number/unit number combination (" & SerialNumber & "/" & UnitNum & ") In the database. Cannot continue.")
			success = False
		Else
			If (trans Is Nothing) Then
				internalTransaction = True
				trans = GRoomConnection.BeginTransaction(IsolationLevel.ReadCommitted)
			End If

            Dim copyQuery As String = "INSERT INTO AvailableDisks SELECT Batch_Number, Disk_Number, SetupID, DiskType, TimeTested, Energy_1_Current, " & _
            "Energy_2_Current, Energy_3_Current, E1_Charge_Setpoint, E2_Charge_Setpoint, E3_Charge_Setpoint, V1mA_Volts, " & _
              "V1mA_Current, Vref_Volts, Watts_Watts, Watts_Temp, V10kA_Volts, V10kA_Current, SCP_1000A_Volts, Test_Failed, Category, " & _
              "CntrHold, ring_current, ring_volts, ring_energy " & _
              "FROM ConsumedDisks WHERE ConsumedDisks.SerialNumber = '" & SerialNumber & "' And ConsumedDisks.UnitNumber = " & UnitNum

			Dim copyQueryCommand As New SqlCommand(copyQuery, GRoomConnection)
			copyQueryCommand.Transaction = trans

			Dim eraseQuery As String = "DELETE FROM ConsumedDisks WHERE SerialNumber = '" & SerialNumber & "' AND UnitNumber = " & UnitNum
			Dim eraseQueryCommand As New SqlCommand(eraseQuery, GRoomConnection)
			eraseQueryCommand.Transaction = trans

			If (copyQueryCommand.ExecuteNonQuery() <= 0) Then
				success = False
				MsgBox("Error copying the disks back to the available disks table.")
			End If
			If (eraseQueryCommand.ExecuteNonQuery() <= 0) Then
				success = False
				MsgBox("Error erasing disks from the consumed disks table!")
			End If

			Dim unitCountField As String = "Unit" & UnitNum & "NumBuilt"
			Me.Lookup(SerialNumber, trans)
			Dim decrementQuery As String = "UPDATE WO_Info SET " & unitCountField & " = " & unitCountField & "-1 WHERE WO_Number = '" & WorkOrderNum & "'"

			Dim decrementQueryCommand As New SqlCommand(decrementQuery, GRoomConnection)
			decrementQueryCommand.Transaction = trans

			If (decrementQueryCommand.ExecuteNonQuery <= 0) Then
				success = False
				MsgBox("Error changing the quantity in the WO_Info Table!")
			End If

			Dim DelFromBuildUnitsQuery As String = "DELETE FROM BuiltUnits " & _
			"WHERE SerialNumber = '" & SerialNumber & "' " & _
			"AND Unit = " & UnitNum & " "

			Dim DelFromBuildUnitsCommand As New SqlCommand(DelFromBuildUnitsQuery, GRoomConnection)
			DelFromBuildUnitsCommand.Transaction = trans

			If (DelFromBuildUnitsCommand.ExecuteNonQuery <= 0) Then
				success = False
				MsgBox("Error erasing from the built units table!")
			End If

			If (internalTransaction) Then
				If (success) Then
					trans.Commit()
				Else
					trans.Rollback()
				End If

			End If
		End If

		Return success

	End Function

	Public Function GetUnitsBuilt(ByVal unitNumber As Integer) As Integer
		Dim retVal As Integer = 0
		Dim query As String = "SELECT * FROM WO_Info WHERE WO_Number = '" & m_workOrderNum & "'"
		Dim unitsLookupCommand As New SqlCommand(query, GRoomConnection)
		Dim unitsLookupReader As SqlDataReader = unitsLookupCommand.ExecuteReader()
		Try
			Try
				unitsLookupReader.Read()
				retVal = CInt(unitsLookupReader.Item("Unit" & unitNumber & "NumBuilt"))
			Catch ex As Exception
				MsgBox("Error reading the number of units built for unit " & unitNumber, MsgBoxStyle.Critical)
				retVal = 0
			End Try
		Finally
			unitsLookupReader.Close()
		End Try
		Return retVal

	End Function

	Public Function GetSerialNumbersUnBuilt(ByVal UnitModNumber As Integer) As Int64()
		'the user just selected a unit for the list.  Get the serial numbers that have
		'been built for it.
		Dim builtSerialsQuery As String
		builtSerialsQuery = "SELECT d.SerialNumber 'SerialNum' FROM WO_SerialNumbers d " & _
		" WHERE (d.WO_Number = '" & WorkOrderNum & "') EXCEPT SELECT  d.SerialNumber 'SerialNum' FROM " & _
		"BuiltUnits b INNER JOIN WO_SerialNumbers d ON b.SerialNumber = d.SerialNumber" & _
		" WHERE (b.Unit = " & UnitModNumber & " ) AND (d.WO_Number = '" & WorkOrderNum & "') Order by d.SerialNumber"

		Dim BuiltSerialsCommand As SqlCommand = New SqlCommand(builtSerialsQuery, GRoomConnection)
		Dim BuiltSerialsResults As SqlDataReader = BuiltSerialsCommand.ExecuteReader()

		Dim SerialNumsBuilt As New ArrayList
		If (BuiltSerialsResults.HasRows()) Then
			While BuiltSerialsResults.Read()
				Dim tmpParse As Int64
				If (Int64.TryParse(BuiltSerialsResults!SerialNum.ToString, tmpParse)) Then
					SerialNumsBuilt.Add(tmpParse)
				Else
					MsgBox("Error parsing serial number to integer!", MsgBoxStyle.Critical)
					MsgBox(BuiltSerialsResults!SerialNum.ToString, MsgBoxStyle.Information)
				End If

			End While
		End If
		BuiltSerialsResults.Close()

		Return CType(SerialNumsBuilt.ToArray(GetType(Int64)), Int64())
	End Function

	Public Function GetAllSerialNumbers() As List(Of Int64)
		'Changed from Arrary to List  -8May2012 RPost
		Dim AllSerialsQuery As String
		AllSerialsQuery = "SELECT SerialNumber FROM WO_SerialNumbers " & _
		"WHERE WO_Number = '" & WorkOrderNum & "'"

		Dim AllSerialsCommand As SqlCommand = New SqlCommand(AllSerialsQuery, GRoomConnection)
		Dim AllSerialsResults As SqlDataReader = AllSerialsCommand.ExecuteReader()

		Dim SerialNums As New List(Of Int64)
		If (AllSerialsResults.HasRows()) Then
			While AllSerialsResults.Read()
				Dim tmpParse As Int64
				If (Int64.TryParse(AllSerialsResults!SerialNumber.ToString, tmpParse)) Then '!SerialNum  typo fixed 8May2012 RPost
					SerialNums.Add(tmpParse)
				Else
					MsgBox("Error parsing serial number " & AllSerialsResults!SerialNum.ToString & " to integer!", MsgBoxStyle.Critical)
					MsgBox(AllSerialsResults!SerialNum.ToString, MsgBoxStyle.Information)
				End If
			End While
		End If
		AllSerialsResults.Close()

		Return SerialNums
	End Function

	Public Function GetSerialNumbersUnitsBuilt() As List(Of BasicUnit)
		Dim AllSerialsQuery As String
		AllSerialsQuery = "SELECT * FROM BuiltUnits WHERE SerialNumber IN (SELECT SerialNumber FROM WO_SerialNumbers " & _
		"WHERE WO_Number = '" & WorkOrderNum & "')"

		Dim BuiltSerialsCommand As SqlCommand = New SqlCommand(AllSerialsQuery, GRoomConnection)
		Dim BuiltSerialsResults As SqlDataReader = BuiltSerialsCommand.ExecuteReader()

		Dim BuiltUnits As New List(Of BasicUnit)
		If (BuiltSerialsResults.HasRows()) Then
			While BuiltSerialsResults.Read()
				BuiltUnits.Add(New BasicUnit(CType(BuiltSerialsResults!SerialNumber, Long), CInt(BuiltSerialsResults!Unit)))
			End While
		End If
		BuiltSerialsResults.Close()

		Return BuiltUnits
	End Function

	Public Function numRequired() As Integer
		Dim numReqdQuery As String = "SELECT NumArrestersRequired FROM WO_Info WHERE WO_Number = '" & WorkOrderNum & "'"
		Dim numReqdCommand As New SqlCommand(numReqdQuery, GRoomConnection)
		Dim numReqd As Integer
		Integer.TryParse(CStr(numReqdCommand.ExecuteScalar()), numReqd)
		Return numReqd
	End Function

	Public Function GetNumBuilt() As Integer
		Dim numBuiltQuery As String = "SELECT Unit1NumBuilt + Unit2NumBuilt + Unit3NumBuilt + Unit4NumBuilt + Unit5NumBuilt + Unit6NumBuilt" & _
		" FROM WO_Info WHERE WO_NUMBER = '" & WorkOrderNum & "'"
		Dim numBuiltCommand As New SqlCommand(numBuiltQuery, GRoomConnection)
		Dim numBuilt As Integer = 0
		Integer.TryParse(CStr(numBuiltCommand.ExecuteScalar()), numBuilt)
		Return numBuilt
	End Function

	Public Function Lookup(ByVal SerialNumber As Int64, ByRef trans As SqlTransaction) As Boolean
		Dim success As Boolean = True
		Dim snWoQuery As String = "SELECT * FROM WO_SerialNumbers WHERE SerialNumber = '" & SerialNumber & "'"
		Dim snWoCommand As New SqlCommand(snWoQuery, GRoomConnection)
		If (Not trans Is Nothing) Then
			snWoCommand.Transaction = trans
		End If
		Dim snWoReader As SqlDataReader = snWoCommand.ExecuteReader()
		If (snWoReader.HasRows()) Then
			snWoReader.Read()
			m_workOrderNum = CStr(snWoReader!WO_Number)
			m_catNum = New CatNum()
			m_catNum.WorkOrder = Me
			m_catNum.SerialNumber = SerialNumber
			success = True
		Else
			success = False
			MsgBox("Could not find a matching work order for this serial number: " & SerialNumber)
		End If
		snWoReader.Close()
		If (success) Then
			success = Exists(trans)
		End If
		Return success
	End Function

End Class
