Option Strict On
Imports System.Data.SqlClient
Public Class Disk

    Enum States
        Available = 0
        Consumed = 1
    End Enum
    Enum EdgeTypes
        ERR = -1
        Glass = 0
        Epoxy = 1
    End Enum

    'Fields
    Private m_batch As String
    Private m_diskNum As Integer
    Private m_vRef As Double = -2
    Private m_v10kA As Double = -2
    Private m_category As Integer = -2
    Private m_diskType As String = Nothing
    Private m_state As Integer = -2
    Private m_allowConsumedDisks As Boolean = False
    Private m_parentModule As ArresterModule
    Private m_validForUse As Boolean = True
    Private m_edgeType As Integer = -2



    Sub New(ByRef parent As ArresterModule, ByVal batchIn As String, ByVal diskNumIn As Integer)
        m_parentModule = parent
        m_batch = batchIn
        m_diskNum = diskNumIn
    End Sub
    Sub New(ByRef parent As ArresterModule, ByVal batchIn As String, ByVal diskNumIn As Integer, ByVal AllowConsumed As Boolean)
        m_parentModule = parent
        m_batch = batchIn
        m_diskNum = diskNumIn
        m_allowConsumedDisks = AllowConsumed
    End Sub

    Property Batch() As String
        Get
            Return m_batch
        End Get
        Set(ByVal value As String)
            m_batch = value
            LookupDisk()
        End Set
    End Property

    Property DiskNum() As Integer
        Get
            Return m_diskNum
        End Get
        Set(ByVal value As Integer)
            m_diskNum = value
            LookupDisk()
        End Set
    End Property

    ReadOnly Property VRef() As Double
        Get
            If (m_vRef = -2) Then
                LookupDisk()
            End If
            Return m_vRef
        End Get
    End Property

    ReadOnly Property V10kA() As Double
        Get
            If (m_v10kA = -2) Then
                LookupDisk()
            End If
            Return m_v10kA
        End Get
    End Property

    ReadOnly Property Category() As Integer
        Get
            If (m_category = -2) Then
                LookupDisk()
            End If
            Return m_category
        End Get
    End Property

    ReadOnly Property DiskType() As String
        Get
            If (m_diskType = Nothing) Then
                LookupDisk()
            End If
            Return m_diskType
        End Get
    End Property

    ReadOnly Property State() As Integer
        Get
            If (m_state = -2) Then
                LookupDisk()
            End If
            Return m_state
        End Get
    End Property

    ReadOnly Property ValidForUse() As Boolean
        Get
            If (m_validForUse = True) Then
                LookupDisk()
            End If
            Return m_validForUse
        End Get
    End Property

    ReadOnly Property EdgeType() As EdgeTypes
        Get
            If (DiskType.Substring(DiskType.Length - 1, 1) = "G") Then
                Return EdgeTypes.Glass
            ElseIf (DiskType.Substring(DiskType.Length - 1, 1) = "E") Then
                Return EdgeTypes.Epoxy
            Else
                Return EdgeTypes.ERR
            End If
        End Get
    End Property


    Private Sub LookupDisk()
        m_validForUse = True
        Dim AvailableQuery As String = "SELECT * FROM availableDisks WHERE Batch_Number = '" & m_batch & "' And Disk_Number = " & m_diskNum

        Dim AvailableQueryCommand As New SqlCommand(AvailableQuery, GRoomConnection)
        Dim AvailableQueryReader As SqlDataReader = AvailableQueryCommand.ExecuteReader()
        If (AvailableQueryReader.HasRows()) Then
            AvailableQueryReader.Read()
            m_diskType = CStr(AvailableQueryReader!DiskType)
            m_state = States.Available
            m_vRef = CDbl(AvailableQueryReader!Vref_Volts)
            m_v10kA = CDbl(AvailableQueryReader!V10kA_Volts)
            m_category = CInt(AvailableQueryReader!Category)
            ' Added code to handle special case where category For AT53, AT53 and AT54 are given as a three digit number  16Apr2012
            If (m_diskType.Substring(0, 3) = "AT5") Then
                Select Case (m_diskType.Substring(0, 4))
                    Case "AT52"
                        If m_category < 1000 Then m_category += 1000  'RVA 3/23/22 added condition to format
                    Case "AT53"
                        If m_category < 1000 Then m_category += 1000  'RVA 3/23/22 added condition to format
                    Case "AT54"
                        If m_category < 500 Then
                            m_category += 1000
                        End If
                End Select
            End If
            AvailableQueryReader.Close()
        Else
            AvailableQueryReader.Close()
            Dim ConsumedQuery As String = "SELECT * FROM consumedDisks WHERE Batch_Number = '" & m_batch & "' And Disk_Number = " & m_diskNum
            Dim ConsumedQueryCommand As New SqlCommand(ConsumedQuery, GRoomConnection)
            Dim ConsumedQueryReader As SqlDataReader = ConsumedQueryCommand.ExecuteReader()
            If (ConsumedQueryReader.HasRows()) Then
                If (m_allowConsumedDisks) Then

                    ConsumedQueryReader.Read()
                    m_diskType = CStr(ConsumedQueryReader!DiskType)
                    m_state = States.Consumed
                    m_vRef = CDbl(ConsumedQueryReader!Vref_Volts)
                    m_v10kA = CDbl(ConsumedQueryReader!V10kA_Volts)
                    m_category = CInt(ConsumedQueryReader!Category)
                    ' remove this code and replace with code imediately following as this is incorrect  16Apr2012
                    ' If (m_diskType.Substring(0, 4) = "AT52") Then
                    '   m_category += 1000
                    'End If
                    If (m_diskType.Substring(0, 3) = "AT5") Then
                        Select Case (m_diskType.Substring(0, 4))
                            Case "AT52"
                                m_category += 1000
                            Case "AT53"
                                m_category += 1000
                            Case "AT54"
                                If m_category < 500 Then
                                    m_category += 1000
                                End If
                        End Select
                    End If


                Else
                    MsgBox("This disk has already been consumed! Batch:" & m_batch & ", Disk Number:" & m_diskNum, MsgBoxStyle.Critical)
                    m_diskType = "ERR"
                    m_category = -1
                    m_state = -1
                    m_vRef = -1
                    m_v10kA = -1
                    m_validForUse = False
                End If

            Else
                MsgBox("Could not find a matching disk in the database for Batch:" & m_batch & ", Disk Number:" & m_diskNum, MsgBoxStyle.Critical)
                m_diskType = "ERR"
                m_category = -1
                m_state = -1
                m_vRef = -1
                m_validForUse = False
                m_v10kA = -1
        End If
        ConsumedQueryReader.Close()
        End If
    End Sub

    Public Function ConsumeDisk(ByRef trans As SqlClient.SqlTransaction) As Boolean
        Dim success As Boolean = True
        If (m_state = States.Available) Then
            'Dim insertQuery As String = "INSERT INTO consumedDisks SELECT '" & m_parentModule.CatNumber.SerialNumber & _
            '"', " & m_parentModule.UnitNumber & " , Batch_Number, Disk_Number, TimeTested, '" & Now() & "' , DiskType, SetupID, Energy_1_Current, " & _
            '"Energy_2_Current, Energy_3_Current, E1_Charge_Setpoint, " & _
            '"E2_Charge_Setpoint, E3_Charge_Setpoint, V1mA_Volts, V1mA_Current, Vref_Volts, Watts_Watts, Watts_Temp, " & _
            '"V10kA_Volts, V10kA_Current, SCP_1000A_Volts, Test_Failed, Category, CntrHold FROM AvailableDisks WHERE " & _
            '"Disk_Number = " & DiskNum & " And Batch_Number = '" & Batch & "';"


            '7May2012 RPost replaced UnitNumber with OldModNumber
            'Dim insertQuery As String = "INSERT INTO consumedDisks SELECT '" & m_parentModule.CatNumber.SerialNumber & _
            '"', " & m_parentModule.OldModNumber & " , Batch_Number, Disk_Number, TimeTested, '" & Now() & "' , DiskType, SetupID, Energy_1_Current, " & _
            '"Energy_2_Current, Energy_3_Current, E1_Charge_Setpoint, " & _
            '"E2_Charge_Setpoint, E3_Charge_Setpoint, V1mA_Volts, V1mA_Current, Vref_Volts, Watts_Watts, Watts_Temp, " & _
            '"V10kA_Volts, V10kA_Current, SCP_1000A_Volts, Test_Failed, Category, CntrHold FROM AvailableDisks WHERE " & _
            '"Disk_Number = " & DiskNum & " And Batch_Number = '" & Batch & "';"

            ' 2017-04-24 RVA - Added 3 fields for ringing test
            Dim insertQuery As String = "INSERT INTO consumedDisks (" & _
                "SerialNumber, UnitNumber, Batch_Number, Disk_Number," & _
                "TimeTested, TimeConsumed, DiskType, SetupID, Energy_1_Current, Energy_2_Current, Energy_3_Current," & _
                "E1_Charge_Setpoint, E2_Charge_Setpoint, E3_Charge_Setpoint," & _
                "V1mA_Volts, V1mA_Current, Vref_Volts, Watts_Watts, Watts_Temp, V10kA_Volts, V10kA_Current, SCP_1000A_Volts," & _
                "Test_Failed, Category, CntrHold, ring_current, ring_volts, ring_energy" & _
                ") SELECT '" & m_parentModule.CatNumber.SerialNumber & _
            "', " & m_parentModule.OldModNumber & " , Batch_Number, Disk_Number, TimeTested, '" & Now() & "' , DiskType, SetupID, Energy_1_Current, " & _
            "Energy_2_Current, Energy_3_Current, E1_Charge_Setpoint, " & _
            "E2_Charge_Setpoint, E3_Charge_Setpoint, V1mA_Volts, V1mA_Current, Vref_Volts, Watts_Watts, Watts_Temp, " & _
            "V10kA_Volts, V10kA_Current, SCP_1000A_Volts, Test_Failed, Category, CntrHold, ring_current, ring_volts, ring_energy " & _
            "FROM AvailableDisks WHERE " & _
            "Disk_Number = " & DiskNum & " And Batch_Number = '" & Batch & "';"

            Dim insertCommand As New SqlCommand(insertQuery, GRoomConnection)
            insertCommand.Transaction = trans


            Dim deleteQuery As String = "DELETE FROM AvailableDisks Where Disk_Number = " & DiskNum & " And Batch_Number = '" & Batch & "'"

            Dim deleteCommand As New SqlCommand(deleteQuery, GRoomConnection)
            deleteCommand.Transaction = trans
            If ((insertCommand.ExecuteNonQuery() <> 1) Or (deleteCommand.ExecuteNonQuery()) <> 1) Then
                MsgBox("Error moving Batch: " & Batch & ", Disk: " & DiskNum & " From available disks to consumed disks.", MsgBoxStyle.Critical)
                success = False
            End If
        Else
            MsgBox("Batch: " & Batch & ", Disk:" & DiskNum & " is already consumed!")
            success = False
        End If
        If (success) Then
            m_state = States.Consumed
        End If
        Return success
    End Function


    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim compDisk As Disk = DirectCast(obj, Disk)

        If ( _
            m_batch = compDisk.m_batch _
            And m_diskNum = compDisk.m_diskNum _
            ) Then
            Return True
        Else
            Return MyBase.Equals(obj)
        End If


    End Function




End Class
