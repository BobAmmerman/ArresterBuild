<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UnbuildModule
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(UnbuildModule))
        Me.SerialNumberTextBox = New System.Windows.Forms.TextBox()
        Me.SerialNumberLabel = New System.Windows.Forms.Label()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.CancelUnbuildButton = New System.Windows.Forms.Button()
        Me.ModuleLabel = New System.Windows.Forms.Label()
        Me.ModulesListBox = New System.Windows.Forms.ListBox()
        Me.SuspendLayout()
        '
        'SerialNumberTextBox
        '
        Me.SerialNumberTextBox.Location = New System.Drawing.Point(31, 31)
        Me.SerialNumberTextBox.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.SerialNumberTextBox.MaxLength = 12
        Me.SerialNumberTextBox.Name = "SerialNumberTextBox"
        Me.SerialNumberTextBox.Size = New System.Drawing.Size(164, 22)
        Me.SerialNumberTextBox.TabIndex = 0
        '
        'SerialNumberLabel
        '
        Me.SerialNumberLabel.AutoSize = True
        Me.SerialNumberLabel.Location = New System.Drawing.Point(27, 11)
        Me.SerialNumberLabel.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.SerialNumberLabel.Name = "SerialNumberLabel"
        Me.SerialNumberLabel.Size = New System.Drawing.Size(102, 17)
        Me.SerialNumberLabel.TabIndex = 2
        Me.SerialNumberLabel.Text = "Serial Number:"
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(289, 213)
        Me.OKButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(100, 28)
        Me.OKButton.TabIndex = 2
        Me.OKButton.Text = "Okay"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'CancelUnbuildButton
        '
        Me.CancelUnbuildButton.Location = New System.Drawing.Point(28, 213)
        Me.CancelUnbuildButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.CancelUnbuildButton.Name = "CancelUnbuildButton"
        Me.CancelUnbuildButton.Size = New System.Drawing.Size(100, 28)
        Me.CancelUnbuildButton.TabIndex = 3
        Me.CancelUnbuildButton.Text = "Cancel"
        Me.CancelUnbuildButton.UseVisualStyleBackColor = True
        '
        'ModuleLabel
        '
        Me.ModuleLabel.AutoSize = True
        Me.ModuleLabel.Location = New System.Drawing.Point(27, 59)
        Me.ModuleLabel.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.ModuleLabel.Name = "ModuleLabel"
        Me.ModuleLabel.Size = New System.Drawing.Size(58, 17)
        Me.ModuleLabel.TabIndex = 6
        Me.ModuleLabel.Text = "Module:"
        '
        'ModulesListBox
        '
        Me.ModulesListBox.FormattingEnabled = True
        Me.ModulesListBox.ItemHeight = 16
        Me.ModulesListBox.Location = New System.Drawing.Point(28, 79)
        Me.ModulesListBox.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ModulesListBox.Name = "ModulesListBox"
        Me.ModulesListBox.Size = New System.Drawing.Size(360, 100)
        Me.ModulesListBox.TabIndex = 1
        '
        'UnbuildModule
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(420, 256)
        Me.Controls.Add(Me.ModulesListBox)
        Me.Controls.Add(Me.ModuleLabel)
        Me.Controls.Add(Me.CancelUnbuildButton)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.SerialNumberLabel)
        Me.Controls.Add(Me.SerialNumberTextBox)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "UnbuildModule"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "UnbuildUnit"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents SerialNumberTextBox As System.Windows.Forms.TextBox
    Friend WithEvents SerialNumberLabel As System.Windows.Forms.Label
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Friend WithEvents CancelUnbuildButton As System.Windows.Forms.Button
    Friend WithEvents ModuleLabel As System.Windows.Forms.Label
    Friend WithEvents ModulesListBox As System.Windows.Forms.ListBox
End Class
