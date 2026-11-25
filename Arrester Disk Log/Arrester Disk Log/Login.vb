Option Strict On
Imports System
Imports System.Timers
Imports System.Data.SqlClient

Public Class Login
    Public ActiveLogin As Boolean
    Public EmployeeID As Integer
    Public Password As String
    Public FirstName As String
    Public LastName As String
    Public Admin As Boolean
    Private logoutTimer As System.Timers.Timer
    Private Alive As Boolean
    Private loginTimeoutSeconds As Integer



    Public Sub New()
        loginTimeoutSeconds = 10
        Alive = True
        logoutTimer = New System.Timers.Timer()
        logoutTimer.Interval = loginTimeoutSeconds * 1000
        logoutTimer.AutoReset = True
        AddHandler logoutTimer.Elapsed, AddressOf OnLogoutTimerTick
        Admin = False
        ActiveLogin = False
    End Sub

    Public Function Login(ByVal EmployeeIDIn As Integer, ByVal PasswordIn As String) As Boolean

        Dim retVal As Boolean

        If (ngd_ProductionConnection.State = ConnectionState.Open) Then



            retVal = False
            Admin = False

            Dim matchedUsersQuery As String = "SELECT COUNT(*) FROM Employees WHERE EmployeeNumber = " & EmployeeIDIn & " And Password = '" & PasswordIn & "'"
            Dim matchedUsersCmd As New SqlCommand(matchedUsersQuery, ngd_ProductionConnection)
            Dim matchedUsers As Integer = CInt(matchedUsersCmd.ExecuteScalar())
            Dim loginQuery As String = "SELECT * FROM Employees WHERE EmployeeNumber = " & EmployeeIDIn & " And Password = '" & PasswordIn & "'"
            'loginResults = loginTableAdapter.LoginVerifyQuery(EmployeeIDIn, PasswordIn)
            Dim loginCommand As New SqlCommand(loginQuery, ngd_ProductionConnection)
            Dim loginResults As SqlDataReader = loginCommand.ExecuteReader()


            If matchedUsers <> 1 Then
                ActiveLogin = False
                retVal = False
            Else
                loginResults.Read()
                EmployeeID = CInt(loginResults!EmployeeNumber)
                Password = CStr(loginResults!Password)
                FirstName = CStr(loginResults!FirstName)
                LastName = CStr(loginResults!LastName)
                ActiveLogin = True
                retVal = True
                If CStr(loginResults!Access_Level) = "ADMINISTRATOR" Then
                    Admin = True
                End If
            End If
            loginResults.Close()
            logoutTimer.Start()

        Else
            MsgBox("Need database connection to log in!")
            retVal = False
        End If



        Return retVal

    End Function

    Public Sub Logout()
        ActiveLogin = False
        EmployeeID = Nothing
        Password = Nothing
        FirstName = Nothing
        LastName = Nothing
        Admin = False
    End Sub

    Private Sub OnLogoutTimerTick(ByVal source As Object, ByVal e As ElapsedEventArgs)
        If Alive = True Then
            Alive = False
        ElseIf Alive = False Then
            ActiveLogin = False
        End If
    End Sub


End Class
