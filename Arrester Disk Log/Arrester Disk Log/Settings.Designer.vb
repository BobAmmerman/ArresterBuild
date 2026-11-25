<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Settings
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.DBUserName = New System.Windows.Forms.TextBox()
        Me.DBPassword = New System.Windows.Forms.TextBox()
        Me.AutoPrintCheckbox = New System.Windows.Forms.CheckBox()
        Me.DBPasswordLabel = New System.Windows.Forms.Label()
        Me.UsernameLabel = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.DBHostname = New System.Windows.Forms.TextBox()
        Me.GlassEpoxySubCheckbox = New System.Windows.Forms.CheckBox()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.CancelSettingsButton = New System.Windows.Forms.Button()
        Me.LoginTimeoutTextBox = New System.Windows.Forms.TextBox()
        Me.LoginTimeoutLabel = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'DBUserName
        '
        Me.DBUserName.Location = New System.Drawing.Point(15, 64)
        Me.DBUserName.Name = "DBUserName"
        Me.DBUserName.Size = New System.Drawing.Size(175, 20)
        Me.DBUserName.TabIndex = 0
        '
        'DBPassword
        '
        Me.DBPassword.Location = New System.Drawing.Point(15, 103)
        Me.DBPassword.Name = "DBPassword"
        Me.DBPassword.Size = New System.Drawing.Size(177, 20)
        Me.DBPassword.TabIndex = 1
        Me.DBPassword.UseSystemPasswordChar = True
        '
        'AutoPrintCheckbox
        '
        Me.AutoPrintCheckbox.AutoSize = True
        Me.AutoPrintCheckbox.Location = New System.Drawing.Point(15, 173)
        Me.AutoPrintCheckbox.Name = "AutoPrintCheckbox"
        Me.AutoPrintCheckbox.Size = New System.Drawing.Size(152, 17)
        Me.AutoPrintCheckbox.TabIndex = 2
        Me.AutoPrintCheckbox.Text = "Automatically Print Reports"
        Me.AutoPrintCheckbox.UseVisualStyleBackColor = True
        '
        'DBPasswordLabel
        '
        Me.DBPasswordLabel.AutoSize = True
        Me.DBPasswordLabel.Location = New System.Drawing.Point(12, 87)
        Me.DBPasswordLabel.Name = "DBPasswordLabel"
        Me.DBPasswordLabel.Size = New System.Drawing.Size(53, 13)
        Me.DBPasswordLabel.TabIndex = 3
        Me.DBPasswordLabel.Text = "Password"
        '
        'UsernameLabel
        '
        Me.UsernameLabel.AutoSize = True
        Me.UsernameLabel.Location = New System.Drawing.Point(12, 48)
        Me.UsernameLabel.Name = "UsernameLabel"
        Me.UsernameLabel.Size = New System.Drawing.Size(55, 13)
        Me.UsernameLabel.TabIndex = 4
        Me.UsernameLabel.Text = "Username"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(89, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Server Hostname"
        '
        'DBHostname
        '
        Me.DBHostname.Location = New System.Drawing.Point(15, 25)
        Me.DBHostname.Name = "DBHostname"
        Me.DBHostname.Size = New System.Drawing.Size(175, 20)
        Me.DBHostname.TabIndex = 8
        '
        'GlassEpoxySubCheckbox
        '
        Me.GlassEpoxySubCheckbox.AutoSize = True
        Me.GlassEpoxySubCheckbox.Location = New System.Drawing.Point(15, 196)
        Me.GlassEpoxySubCheckbox.Name = "GlassEpoxySubCheckbox"
        Me.GlassEpoxySubCheckbox.Size = New System.Drawing.Size(177, 17)
        Me.GlassEpoxySubCheckbox.TabIndex = 9
        Me.GlassEpoxySubCheckbox.Text = "Allow Glass/Epoxy Substitutions"
        Me.GlassEpoxySubCheckbox.UseVisualStyleBackColor = True
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(117, 219)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(75, 23)
        Me.OKButton.TabIndex = 10
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'CancelSettingsButton
        '
        Me.CancelSettingsButton.Location = New System.Drawing.Point(12, 219)
        Me.CancelSettingsButton.Name = "CancelSettingsButton"
        Me.CancelSettingsButton.Size = New System.Drawing.Size(75, 23)
        Me.CancelSettingsButton.TabIndex = 11
        Me.CancelSettingsButton.Text = "Cancel"
        Me.CancelSettingsButton.UseVisualStyleBackColor = True
        '
        'LoginTimeoutTextBox
        '
        Me.LoginTimeoutTextBox.Location = New System.Drawing.Point(15, 147)
        Me.LoginTimeoutTextBox.Name = "LoginTimeoutTextBox"
        Me.LoginTimeoutTextBox.Size = New System.Drawing.Size(100, 20)
        Me.LoginTimeoutTextBox.TabIndex = 12
        '
        'LoginTimeoutLabel
        '
        Me.LoginTimeoutLabel.AutoSize = True
        Me.LoginTimeoutLabel.Location = New System.Drawing.Point(12, 126)
        Me.LoginTimeoutLabel.Name = "LoginTimeoutLabel"
        Me.LoginTimeoutLabel.Size = New System.Drawing.Size(125, 13)
        Me.LoginTimeoutLabel.TabIndex = 13
        Me.LoginTimeoutLabel.Text = "Login Timeout (Seconds)"
        '
        'Settings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(202, 254)
        Me.Controls.Add(Me.LoginTimeoutLabel)
        Me.Controls.Add(Me.LoginTimeoutTextBox)
        Me.Controls.Add(Me.CancelSettingsButton)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.GlassEpoxySubCheckbox)
        Me.Controls.Add(Me.DBHostname)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.UsernameLabel)
        Me.Controls.Add(Me.DBPasswordLabel)
        Me.Controls.Add(Me.AutoPrintCheckbox)
        Me.Controls.Add(Me.DBPassword)
        Me.Controls.Add(Me.DBUserName)
        Me.Name = "Settings"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Settings"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DBUserName As System.Windows.Forms.TextBox
    Friend WithEvents DBPassword As System.Windows.Forms.TextBox
    Friend WithEvents AutoPrintCheckbox As System.Windows.Forms.CheckBox
    Friend WithEvents DBPasswordLabel As System.Windows.Forms.Label
    Friend WithEvents UsernameLabel As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DBHostname As System.Windows.Forms.TextBox
    Friend WithEvents GlassEpoxySubCheckbox As System.Windows.Forms.CheckBox
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Friend WithEvents CancelSettingsButton As System.Windows.Forms.Button
    Friend WithEvents LoginTimeoutTextBox As System.Windows.Forms.TextBox
    Friend WithEvents LoginTimeoutLabel As System.Windows.Forms.Label
End Class
