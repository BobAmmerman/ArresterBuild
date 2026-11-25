Option Strict On
Public Class frmLogin

    Private loginObject As New Login()
    Private mainForm As New frmMain(loginObject, Me)


    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click

        Dim employeeNum As Integer
        Integer.TryParse(ClockNumTextBox.Text, employeeNum)
        If loginObject.Login(employeeNum, PasswordTextBox.Text) Then
            mainForm.UserName.Text = loginObject.FirstName & " " & loginObject.LastName
            If (loginObject.Admin = False) Then
                mainForm.AdminMenu.Visible = False
            End If
            PasswordTextBox.Text = ""
            ClockNumTextBox.Focus()
            mainForm.Enabled = True
            mainForm.Show()
            Me.Hide()

        Else
            MsgBox("Invalid Login, Please Try Again.")
            PasswordTextBox.Text = ""
            PasswordTextBox.Focus()
        End If

    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

End Class
