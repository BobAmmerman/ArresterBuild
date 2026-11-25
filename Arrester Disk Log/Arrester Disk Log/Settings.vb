Option Strict On
Public Class Settings

    Dim loginForm As frmLogin

    Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.Focus()
    End Sub

    Sub New(ByRef loginFormIn As frmLogin)

        ' This call is required by the designer.
        InitializeComponent()

        loginForm = loginFormIn

        ' Add any initialization after the InitializeComponent() call.
        Me.Focus()

        loginForm.WindowState = FormWindowState.Minimized
        Me.BringToFront()
        Me.Activate()

    End Sub

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OKButton.Click
        My.Settings.Password = DBPassword.Text
        My.Settings.Username = DBUserName.Text
        My.Settings.DBServer = DBHostname.Text
        My.Settings.AutoPrintReport = AutoPrintCheckbox.Checked
        My.Settings.EpoxySubGlass = GlassEpoxySubCheckbox.Checked

        My.Settings.GRoomConnectionString = "Data Source=" & My.Settings.DBServer & ";Initial Catalog=GRoom;User ID=" & My.Settings.Username & ";Password=" & My.Settings.Password & ";"
        My.Settings.PolymerStationConnectionString = "Data Source=" & My.Settings.DBServer & ";Initial Catalog=PolymerStation;User ID=" & My.Settings.Username & ";Password=" & My.Settings.Password & ";"
        My.Settings.ngd_productionConnectionString = "Data Source=" & My.Settings.DBServer & ";Initial Catalog=ngd-production;User ID=" & My.Settings.Username & ";Password=" & My.Settings.Password & ";"

        If (GRoomConnection.State = ConnectionState.Open) Then
            GRoomConnection.Close()
        End If
        If (ngd_ProductionConnection.State = ConnectionState.Open) Then
            ngd_ProductionConnection.Close()
        End If
        If (PolymerStationConnection.State = ConnectionState.Open) Then
            PolymerStationConnection.Close()
        End If
        Dim successDB As Boolean = True
        Try

            GRoomConnection.ConnectionString = My.Settings.GRoomConnectionString
            PolymerStationConnection.ConnectionString = My.Settings.PolymerStationConnectionString
            ngd_ProductionConnection.ConnectionString = My.Settings.ngd_productionConnectionString
            GRoomConnection.Open()
            ngd_ProductionConnection.Open()
            PolymerStationConnection.Open()
        Catch ex As Exception
            MsgBox("The database cannot be reached using the supplied settings. Please check the settings and try again.")
            successDB = False
        End Try



        Dim loginTimeoutMilli As Integer
        Integer.TryParse(LoginTimeoutTextBox.Text, loginTimeoutMilli)
        My.Settings.LoginTimeoutMs = loginTimeoutMilli

        If (successDB And Not loginForm Is Nothing) Then
            loginForm.WindowState = FormWindowState.Normal
            Me.Close()
        ElseIf (successDB) Then
            Me.Close()
        End If

    End Sub

    Private Sub CancelSettingsButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CancelSettingsButton.Click
        Me.Close()
    End Sub

    Private Sub Settings_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        AutoPrintCheckbox.Checked = My.Settings.AutoPrintReport
        GlassEpoxySubCheckbox.Checked = My.Settings.EpoxySubGlass
        DBPassword.Text = My.Settings.Password
        DBUserName.Text = My.Settings.Username
        DBHostname.Text = My.Settings.DBServer
        LoginTimeoutTextBox.Text = CStr(My.Settings.LoginTimeoutMs / 1000)
    End Sub
End Class